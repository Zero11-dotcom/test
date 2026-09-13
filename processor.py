"""
processor.py — MQL5 backtest log parsing + analytics engine.

Pipeline
--------
1. parse_report()        : raw MQL5 Strategy-Tester .xlsx  ->  round-trip trades
2. analyze(params)       : user capital sizing, Exness position distribution,
                           X-MA virtual-mode state machine, KPIs, heatmap series
3. build_trade_log()     : full processed trade log (Live/Virtual, split chunks)

Financial model (documented in README):
- Every raw trade has a return on the report's own equity:
      r_t = net_pnl_raw_t / balance_before_raw_t        (from the Balance column)
- The user's account compounds the same returns on their Initial Capital:
      pnl_user_t = r_t * equity_user_before_t
  This reproduces the strategy exactly, scale-invariant to the report's deposit.
- Position Distribution ON  -> lots = equity / (0.15 * entry_price), split into
  chunks of <= 200 lots (same entry/exit). This is the same rule the EA itself
  used (verified: 3000 / (0.15 * 32632.95) = 0.61 lots), so P&L is unchanged;
  the toggle governs lot bookkeeping + the Exness split.
- Position Distribution OFF -> lots scale proportionally to the user's equity
  vs the report's equity (no 200-lot split applied).
"""

from __future__ import annotations

import math
import os
from dataclasses import dataclass, field

import numpy as np
import pandas as pd

# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #

CONFIG = {
    "DATA_ROOT": r"G:\EXNESS_backtesting_reports",
    "GENERAL_SUBDIR": os.path.join("GENERAL REPORT", "full mql5 report"),
    "PROFITABLE_SUBDIR": os.path.join(
        "ONLY PROFITABLE HOUR + PROFITABLE WEEKDAYS", "full mql5 report"
    ),
    "MARGIN_PCT": 0.15,           # Equity = 15% of position value (Exness)
    "MAX_LOT_PER_POSITION": 200,  # Exness max lot size per position
    "LOT_STEP": 0.01,
    "SIZING_PRICE": None,         # None -> use each trade's entry price.
                                  # Set a number (e.g. 130000) to fix it.
    "MAX_SERIES_POINTS": 2400,    # chart downsample target
}

# The 14 strategy variations. `file_id` maps to <file_id>_fullreport.xlsx and
# <file_id>_onlyprofitable.xlsx inside the two mode folders.
VARIATIONS = [
    {"id": "m5_w_m15_0p1_BTC",   "label": "01 · M5_base__htfM15__thr0p1 · BTC", "group": "Base × HTF × Threshold", "symbol": "BTC"},
    {"id": "m3_w_m15_0p1_BTC",   "label": "02 · M3_base__htfM15__thr0p1 · BTC", "group": "Base × HTF × Threshold", "symbol": "BTC"},
    {"id": "m5_w_m10_0p1_BTC",   "label": "03 · M5_base__htfM10__thr0p1 · BTC", "group": "Base × HTF × Threshold", "symbol": "BTC"},
    {"id": "m5_w_m30_0p1_BTC",   "label": "04 · M5_base__htfM30__thr0p1 · BTC", "group": "Base × HTF × Threshold", "symbol": "BTC"},
    {"id": "m5_w_m15_0p01_BTC",  "label": "05 · M5_base__htfM15__thr0p01 · BTC", "group": "Base × HTF × Threshold", "symbol": "BTC"},
    {"id": "m5_w_m10_0p1_ETH",   "label": "06 · M5_base__htfM10__thr0p1 · ETH", "group": "Base × HTF × Threshold", "symbol": "ETH"},
    {"id": "m5_w_m15_0p1_ETH",   "label": "07 · M5_base__htfM15__thr0p1 · ETH", "group": "Base × HTF × Threshold", "symbol": "ETH"},
    {"id": "m5_w_m10_0p01_ETH",  "label": "08 · M5_base__htfM10__thr0p01 · ETH", "group": "Base × HTF × Threshold", "symbol": "ETH"},
    {"id": "m5_w_m15_0p01_ETH",  "label": "09 · M5_base__htfM15__thr0p01 · ETH", "group": "Base × HTF × Threshold", "symbol": "ETH"},
    {"id": "m20_0p1_BTC",        "label": "10 · m20 BTC - 0.1",                "group": "Single Timeframe",       "symbol": "BTC"},
    {"id": "m15_0p1_BTC",        "label": "11 · m15 BTC - 0.1",                "group": "Single Timeframe",       "symbol": "BTC"},
    {"id": "m1_0p1_BTC",         "label": "12 · m1 BTC - 0.1",                 "group": "Single Timeframe",       "symbol": "BTC"},
    {"id": "m30_0p1_ETH",        "label": "13 · M30 ETH - 0.1",                "group": "Single Timeframe",       "symbol": "ETH"},
    {"id": "m15_0p1_ETH",        "label": "14 · M15 ETH - 0.1",                "group": "Single Timeframe",       "symbol": "ETH"},
]

MODES = {
    "general": CONFIG["GENERAL_SUBDIR"],
    "profitable": CONFIG["PROFITABLE_SUBDIR"],
}

DEAL_COLS = ["Time", "Deal", "Symbol", "Type", "Direction", "Volume", "Price",
             "Order", "Commission", "Swap", "Profit", "Balance", "Comment"]


# --------------------------------------------------------------------------- #
# Parsing
# --------------------------------------------------------------------------- #

def _num(v) -> float:
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(" ", "").replace(",", "")
    mult = 1.0
    if s and s[-1].upper() in "KMB":
        mult = {"K": 1e3, "M": 1e6, "B": 1e9}[s[-1].upper()]
        s = s[:-1]
    return float(s) * mult


def resolve_path(variation_id: str, mode: str) -> str:
    suffix = "_fullreport.xlsx" if mode == "general" else "_onlyprofitable.xlsx"
    fname = variation_id + suffix
    path = os.path.join(CONFIG["DATA_ROOT"], MODES[mode], fname)
    if not os.path.exists(path):
        raise FileNotFoundError(f"Report file not found: {path}")
    return path


def parse_report(path: str) -> dict:
    """Parse an MQL5 Strategy Tester xlsx into round-trip trades.

    Works with both layouts found in the data folder:
      * full report  (Settings / Results / Orders / Deals sections)
      * deals-only   ("OD_" files — Deals table at the top)

    Uses openpyxl in read-only mode directly: some of these workbooks carry
    wrong sheet-dimension metadata that truncates pandas' reader.
    """
    import openpyxl

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        rows = None
        hdr = None
        for ws in wb.worksheets:
            cached = list(ws.iter_rows(values_only=True))
            hdr = None
            for i, row in enumerate(cached):
                c0 = str(row[0]).strip() if row and row[0] is not None else ""
                c1 = str(row[1]).strip() if row and len(row) > 1 and row[1] is not None else ""
                if c0 == "Time" and c1 == "Deal":
                    hdr = i
                    break
            if hdr is not None:
                rows = cached
                break
        if hdr is None:
            raise ValueError(f"No Deals table found in {path}")

        width = max(len(r) for r in rows[hdr:]) if rows else 0
        deals_rows = []
        for row in rows[hdr + 1:]:
            vals = list(row) + [None] * (width - len(row))
            if vals[0] is None or not str(vals[0]).strip():
                continue  # footer / blank
            deals_rows.append(vals)
        deals = pd.DataFrame(deals_rows, columns=DEAL_COLS[:width])

        symbol = None
        deposit = 0.0
        trades = []
        open_pos = None       # accumulating position
        balance_before = None # account balance before the current open position

        for vals in deals_rows:
            r = dict(zip(DEAL_COLS, vals))
            t = pd.to_datetime(r["Time"], format="%Y.%m.%d %H:%M:%S", errors="coerce")
            if pd.isna(t):
                continue
            d_type = str(r.get("Type") or "").strip().lower()
            direction = str(r.get("Direction") or "").strip().lower()
            vol = _num(r.get("Volume"))
            price = _num(r.get("Price"))
            comm = _num(r.get("Commission"))
            balance = _num(r.get("Balance"))
            comment = r.get("Comment")
            comment = str(comment) if comment is not None else ""

            if d_type == "balance":
                deposit = balance
                balance_before = balance
                continue
            if symbol is None and r.get("Symbol"):
                symbol = str(r["Symbol"])

            if direction == "in":
                if open_pos is not None:            # position increase -> merge
                    open_pos["volume"] += vol
                    open_pos["commission"] += comm
                    open_pos["price_sum"] += vol * price
                else:
                    open_pos = {
                        "entry_time": t,
                        "price_sum": vol * price,
                        "volume": vol,
                        "commission": comm,
                        "direction": d_type,        # buy / sell
                        "comment": comment,
                    }
            elif direction in ("out", "in/out"):
                if open_pos is None:                # stray close (shouldn't happen)
                    continue
                entry_price = open_pos["price_sum"] / open_pos["volume"]
                net_pnl = balance - (balance_before if balance_before is not None else deposit)
                trades.append({
                    "entry_time": open_pos["entry_time"],
                    "exit_time": t,
                    "symbol": symbol or "",
                    "direction": open_pos["direction"],
                    "volume_raw": open_pos["volume"],
                    "entry_price": entry_price,
                    "exit_price": price,
                    "net_pnl_raw": net_pnl,
                    "balance_before_raw": balance_before if balance_before is not None else deposit,
                    "balance_after_raw": balance,
                    "comment": open_pos["comment"],
                })
                balance_before = balance
                if direction == "in/out":           # reversal: reopen the other way
                    open_pos = {
                        "entry_time": t,
                        "price_sum": vol * price,
                        "volume": vol,
                        "commission": 0.0,
                        "direction": "sell" if d_type == "buy" else "buy",
                        "comment": "",
                    }
                else:
                    open_pos = None
    finally:
        wb.close()

    df = pd.DataFrame(trades)
    if df.empty:
        raise ValueError(f"No closed trades parsed from {path}")
    return {"trades": df, "deposit": deposit, "symbol": symbol or ""}


# --------------------------------------------------------------------------- #
# Sizing helpers
# --------------------------------------------------------------------------- #

def split_lots(total_lots: float, max_lot: float, step: float) -> list:
    """Split a total lot size into position chunks <= max_lot (Exness limit)."""
    total_lots = math.floor(total_lots / step) * step
    chunks = []
    remaining = round(total_lots, 2)
    while remaining > 1e-9:
        c = min(max_lot, remaining)
        c = math.floor(c / step) * step
        if c <= 0:
            break
        chunks.append(round(c, 2))
        remaining = round(remaining - c, 2)
    return chunks or [0.0]


# --------------------------------------------------------------------------- #
# Core analysis
# --------------------------------------------------------------------------- #

@dataclass
class Params:
    capital: float = 100_000.0
    variation: str = "m5_w_m15_0p1_BTC"
    mode: str = "general"                 # general | profitable
    position_distribution: bool = True    # Exness limit ON/OFF
    ma_filter: bool = False
    ma_period: int = 20
    ma_condition: str = "touch"           # touch | close_below

    def validate(self):
        if self.mode not in MODES:
            raise ValueError(f"mode must be one of {list(MODES)}")
        if not any(v["id"] == self.variation for v in VARIATIONS):
            raise ValueError(f"unknown variation '{self.variation}'")
        self.capital = float(self.capital)
        if self.capital <= 0:
            raise ValueError("Initial capital must be > 0")
        self.ma_period = int(self.ma_period)
        if self.ma_filter and self.ma_period < 2:
            raise ValueError("MA period X must be >= 2")
        if self.ma_condition not in ("touch", "close_below"):
            raise ValueError("ma_condition must be 'touch' or 'close_below'")


def analyze(p: Params) -> dict:
    p.validate()
    path = resolve_path(p.variation, p.mode)
    parsed = parse_report(path)
    df = parsed["trades"].reset_index(drop=True)

    equity = p.capital          # user account equity (official / live)
    underlying = p.capital      # strategy equity incl. virtual trades
    live_equity = p.capital     # official equity (live trades only)

    margin_pct = CONFIG["MARGIN_PCT"]
    max_lot = CONFIG["MAX_LOT_PER_POSITION"]
    step = CONFIG["LOT_STEP"]

    # --- pass 1: sizing on user equity + MA-filter state machine ------------
    recs = []
    for t in df.itertuples(index=False):
        bal_before_raw = t.balance_before_raw
        if bal_before_raw and bal_before_raw > 0 and equity > 0:
            scale = equity / bal_before_raw          # equity-proportional sizing
            ret = t.net_pnl_raw / bal_before_raw
        else:
            scale, ret = 0.0, 0.0                    # blown account / dead report balance

        if p.position_distribution:
            px = CONFIG["SIZING_PRICE"] or t.entry_price
            total_lots = equity / (margin_pct * px) if px else 0.0
            total_lots = max(0.0, math.floor(total_lots / step) * step)
            chunks = split_lots(total_lots, max_lot, step)
        else:
            total_lots = t.volume_raw * scale
            chunks = [round(total_lots, 2)] if total_lots > 0 else [0.0]

        pnl_user = t.net_pnl_raw * scale
        equity = equity + pnl_user            # advance to entry equity of next trade
        recs.append({
            "entry_time": t.entry_time, "exit_time": t.exit_time,
            "symbol": t.symbol, "direction": t.direction,
            "volume_raw": t.volume_raw, "entry_price": t.entry_price,
            "exit_price": t.exit_price, "net_pnl_raw": t.net_pnl_raw,
            "chunks": chunks, "total_lots": sum(chunks),
            "pnl_user": pnl_user, "scale": scale,
            "comment": t.comment,
            "mode": "live",  # assigned below
        })

    # underlying equity curve (all trades, live + virtual)
    pnl_arr = np.array([r["pnl_user"] for r in recs])
    underlying_curve = p.capital + np.cumsum(pnl_arr)

    if p.ma_filter:
        ma = pd.Series(underlying_curve).rolling(p.ma_period, min_periods=p.ma_period).mean().to_numpy()
        state = "live"
        switches = []
        for i in range(len(recs)):
            recs[i]["mode"] = state
            m = ma[i]
            if not np.isnan(m):
                if state == "live":
                    hit = underlying_curve[i] <= m if p.ma_condition == "touch" else underlying_curve[i] < m
                    if hit:
                        switches.append({"index": i, "time": recs[i]["exit_time"], "type": "live->virtual"})
                        state = "virtual"
                else:  # virtual: resume when underlying equity closes back above MA
                    if underlying_curve[i] > m:
                        switches.append({"index": i, "time": recs[i]["exit_time"], "type": "virtual->live"})
                        state = "live"
    else:
        ma = np.full(len(recs), np.nan)
        switches = []
        for r in recs:
            r["mode"] = "live"

    # --- pass 2: official (live) equity --------------------------------------
    live_eq_curve = []
    official_pnls, official_times_exit, official_times_entry = [], [], []
    live_flags = []
    eq_live = p.capital
    eq_under = p.capital
    for i, r in enumerate(recs):
        if r["mode"] == "live":
            eq_live = eq_live + r["pnl_user"]     # may go negative = account blown
            official_pnls.append(r["pnl_user"])
            official_times_exit.append(r["exit_time"])
            official_times_entry.append(r["entry_time"])
        eq_under += r["pnl_user"]
        live_eq_curve.append(eq_live)
        live_flags.append(r["mode"])

    trades_df = pd.DataFrame(recs)
    trades_df["mode"] = live_flags
    trades_df["equity_after"] = underlying_curve
    trades_df["live_equity_after"] = live_eq_curve
    trades_df["ma"] = ma

    analysis = {
        "params": {
            "capital": p.capital, "variation": p.variation,
            "variation_label": next(v["label"] for v in VARIATIONS if v["id"] == p.variation),
            "mode": p.mode, "position_distribution": p.position_distribution,
            "ma_filter": p.ma_filter, "ma_period": p.ma_period,
            "ma_condition": p.ma_condition,
        },
        "meta": {
            "symbol": parsed["symbol"],
            "report_deposit": parsed["deposit"],
            "source_file": os.path.basename(path),
            "first_trade": df["entry_time"].min().strftime("%Y-%m-%d %H:%M"),
            "last_trade": df["exit_time"].max().strftime("%Y-%m-%d %H:%M"),
            "margin_pct": CONFIG["MARGIN_PCT"],
            "max_lot": CONFIG["MAX_LOT_PER_POSITION"],
        },
        "trades": trades_df,
        "ma_series": ma,
        "switches": switches,
    }
    _attach_kpis(analysis, official_pnls, official_times_exit, p)
    _attach_series(analysis, p)
    _attach_heatmap(analysis, official_pnls, official_times_exit)
    _attach_perf(analysis, official_pnls, official_times_entry)
    return analysis


# --------------------------------------------------------------------------- #
# KPIs
# --------------------------------------------------------------------------- #

def _max_dd(equity: list, times: list) -> dict:
    peak = -math.inf
    peak_t = times[0] if times else None
    best = {"abs": 0.0, "pct": 0.0, "peak_date": None, "trough_date": None}
    for e, t in zip(equity, times):
        if e > peak:
            peak, peak_t = e, t
        dd = peak - e
        if dd > best["abs"]:
            best = {
                "abs": round(dd, 2),
                "pct": round(dd / peak * 100, 2) if peak > 0 else 100.0,
                "peak_date": peak_t.strftime("%Y-%m-%d") if peak_t is not None else None,
                "trough_date": t.strftime("%Y-%m-%d"),
            }
    return best


def _period_dd(analysis, p: Params, freq: str) -> dict:
    """Worst peak-to-trough drawdown inside any calendar month ('M') / year ('Y')."""
    tdf = analysis["trades"]
    official = tdf[tdf["mode"] == "live"]
    if official.empty:
        return {"abs": 0.0, "pct": 0.0, "peak_date": None, "trough_date": None, "period": None}
    eq = [p.capital] + [float(v) for v in official["live_equity_after"]]
    tm = [official["exit_time"].min().replace(day=1, hour=0, minute=0)] + list(official["exit_time"])
    s = pd.Series(eq, index=pd.to_datetime(tm))
    best = {"abs": 0.0, "pct": 0.0, "peak_date": None, "trough_date": None, "period": None}
    for _, grp in s.groupby(pd.Grouper(freq=freq)):
        vals, idx = grp.values, grp.index
        if len(vals) < 2:
            continue
        dd = _max_dd(list(vals), list(idx))
        if dd["pct"] > best["pct"] or (dd["pct"] == best["pct"] and dd["abs"] > best["abs"]):
            dd["period"] = idx[0].strftime("%b %Y") if freq == "ME" else str(idx[0].year)
            best = dd
    return best


def _attach_kpis(analysis, official_pnls, official_times_exit, p: Params):
    tdf = analysis["trades"]
    official = tdf[tdf["mode"] == "live"]
    wins = official[official["pnl_user"] > 0]
    gross_win = float(wins["pnl_user"].sum())
    gross_loss = float(official[official["pnl_user"] <= 0]["pnl_user"].sum())
    net = float(official["pnl_user"].sum())
    n_live = len(official)
    n_virtual = int((tdf["mode"] == "virtual").sum())

    dd_m = _period_dd(analysis, p, "ME")
    dd_y = _period_dd(analysis, p, "YE")

    if n_live:
        dur = (official["exit_time"] - official["entry_time"]).mean()
        dur_h = dur.total_seconds() / 3600
        avg_duration = f"{int(dur_h)}h {int(round((dur_h % 1) * 60))}m"
    else:
        avg_duration = "—"

    analysis["kpis"] = {
        "net_pnl": round(net, 2),
        "net_pnl_pct": round(net / p.capital * 100, 2),
        "final_equity": round(p.capital + net, 2),
        "total_trades": n_live + n_virtual,
        "live_trades": n_live,
        "virtual_trades": n_virtual,
        "win_rate": round(len(wins) / n_live * 100, 2) if n_live else 0.0,
        "wins": len(wins),
        "losses": n_live - len(wins),
        "profit_factor": round(gross_win / abs(gross_loss), 2) if gross_loss else None,
        "gross_profit": round(gross_win, 2),
        "gross_loss": round(gross_loss, 2),
        "avg_trade": round(net / n_live, 2) if n_live else 0.0,
        "best_trade": round(float(official["pnl_user"].max()), 2) if n_live else 0.0,
        "worst_trade": round(float(official["pnl_user"].min()), 2) if n_live else 0.0,
        "max_dd_monthly": dd_m,
        "max_dd_yearly": dd_y,
        "mode_switches": len(analysis["switches"]),
        "ma_period": p.ma_period if p.ma_filter else None,
        "avg_trade_duration": avg_duration,
    }


# --------------------------------------------------------------------------- #
# Chart series (downsampled for the frontend)
# --------------------------------------------------------------------------- #

def _downsample(n: int, keep: set, target: int) -> list:
    if n <= target:
        return list(range(n))
    stride = max(1, n // target)
    idx = set(range(0, n, stride)) | {n - 1} | keep
    return sorted(i for i in idx if 0 <= i < n)


def _attach_series(analysis, p: Params):
    tdf = analysis["trades"]
    n = len(tdf)
    keep = {0, n - 1} | {s["index"] for s in analysis["switches"]} | {
        i for s in analysis["switches"] for i in (s["index"] + 1, s["index"] - 1) if 0 <= i < n
    }
    idx = _downsample(n, keep, CONFIG["MAX_SERIES_POINTS"])
    sampled = len(idx) < n

    labels = [tdf["exit_time"].iloc[i].strftime("%Y-%m-%d %H:%M") for i in idx]
    flags = [tdf["mode"].iloc[i] for i in idx]

    analysis["series"] = {
        "sampled": sampled,
        "labels": labels,
        "original": [round(float(tdf["equity_after"].iloc[i]), 2) for i in idx],
        "original_ma": [round(float(v), 2) if not np.isnan(v) else None
                        for v in np.asarray(analysis["ma_series"])[idx]],
        "live_flags": flags,
        # real-only: live trades stitched (virtual trades removed)
        "real_labels": [tdf["exit_time"].iloc[i].strftime("%Y-%m-%d %H:%M")
                        for i in range(n) if tdf["mode"].iloc[i] == "live"],
        "real_values": [round(float(v), 2) for v in
                        tdf.loc[tdf["mode"] == "live", "live_equity_after"]],
    }
    # full-resolution official equity used by drawdown calc / heatmap
    official = tdf[tdf["mode"] == "live"]
    analysis["live_equity_series"] = {
        "values": [float(v) for v in official["live_equity_after"]],
    }


# --------------------------------------------------------------------------- #
# Calendar heatmap (monthly / yearly)
# --------------------------------------------------------------------------- #

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def _attach_heatmap(analysis, official_pnls, official_times_exit):
    tdf = analysis["trades"]
    official = tdf[tdf["mode"] == "live"]
    if official.empty or not official_times_exit:
        analysis["heatmap"] = {"years": [], "grand_total": 0.0}
        return
    s = pd.Series(official["pnl_user"].values,
                  index=pd.DatetimeIndex(official["exit_time"]))
    start = s.index.min().to_period("M")
    end = s.index.max().to_period("M")
    monthly = s.groupby(s.index.to_period("M")).sum()

    years = []
    for year in range(start.year, end.year + 1):
        row = []
        for m in range(1, 13):
            per = pd.Period(f"{year}-{m:02d}", freq="M")
            v = float(monthly.get(per, 0.0))
            in_range = pd.Period(f"{year}-01") <= per <= end
            row.append({
                "month": m, "pnl": round(v, 2),
                "state": "active" if in_range else "out",
            })
        total = round(sum(c["pnl"] for c in row), 2)
        years.append({"year": year, "months": row, "total": total})

    analysis["heatmap"] = {
        "years": years,
        "grand_total": round(float(s.sum()), 2),
        "max_abs": round(max(abs(float(monthly.values.min())),
                             abs(float(monthly.values.max()))) or 1.0, 2),
    }


# --------------------------------------------------------------------------- #
# Weekday / hour performance (official trades, by entry time)
# --------------------------------------------------------------------------- #

def _attach_perf(analysis, official_pnls, official_times_entry):
    tdf = analysis["trades"]
    official = tdf[tdf["mode"] == "live"].copy()
    if official.empty:
        analysis["perf"] = {"weekday": [], "hour": []}
        return
    entry = pd.DatetimeIndex(official["entry_time"])
    wd_names = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    weekday = []
    for wd in range(7):
        m = np.asarray(entry.dayofweek == wd)
        sub = official[m]
        wins = int((sub["pnl_user"] > 0).sum())
        n = int(m.sum())
        weekday.append({
            "day": wd_names[wd], "pnl": round(float(sub["pnl_user"].sum()), 2),
            "trades": n, "win_rate": round(wins / n * 100, 1) if n else 0.0,
        })
    hour = []
    for h in range(24):
        m = np.asarray(entry.hour == h)
        sub = official[m]
        hour.append({
            "hour": f"{h:02d}", "pnl": round(float(sub["pnl_user"].sum()), 2),
            "trades": int(m.sum()),
        })
    analysis["perf"] = {"weekday": weekday, "hour": hour}


# --------------------------------------------------------------------------- #
# Full trade log for export
# --------------------------------------------------------------------------- #

def build_trade_log(analysis) -> pd.DataFrame:
    """Explode chunks -> one row per position (200-lot splits share entry/exit)."""
    tdf = analysis["trades"]
    rows = []
    for i, t in enumerate(tdf.itertuples(index=False)):
        chunks = t.chunks if isinstance(t.chunks, list) else [t.total_lots]
        n_chunks = len(chunks)
        for ci, vol in enumerate(chunks, start=1):
            rows.append({
                "TradeNo": i + 1,
                "Chunk": f"{ci}/{n_chunks}",
                "Mode": t.mode.upper(),
                "Symbol": t.symbol,
                "Direction": t.direction.upper(),
                "EntryTime": t.entry_time.strftime("%Y-%m-%d %H:%M:%S"),
                "ExitTime": t.exit_time.strftime("%Y-%m-%d %H:%M:%S"),
                "Lots (chunk)": round(float(vol), 2),
                "Lots (total)": round(float(t.total_lots), 2),
                "EntryPrice": round(float(t.entry_price), 2),
                "ExitPrice": round(float(t.exit_price), 2),
                "Raw P&L (report)": round(float(t.net_pnl_raw), 2),
                "P&L (chunk USD)": round(float(t.pnl_user) * (vol / t.total_lots if t.total_lots else 0), 2),
                "P&L (total USD)": round(float(t.pnl_user), 2),
                "Equity After (underlying)": round(float(t.equity_after), 2),
                "Live Equity After": round(float(t.live_equity_after), 2),
                f"MA({analysis['params']['ma_period']})": round(float(t.ma), 2) if not np.isnan(t.ma) else "",
                "Comment": t.comment,
            })
    return pd.DataFrame(rows)


def summarize_for_api(analysis) -> dict:
    """JSON-ready subset (no DataFrames)."""
    out = dict(analysis)
    out.pop("trades", None)
    out.pop("ma_series", None)
    out["switches"] = [
        {"time": s["time"].strftime("%Y-%m-%d %H:%M"), "type": s["type"]}
        for s in analysis["switches"]
    ]
    return out


if __name__ == "__main__":
    # smoke test
    res = analyze(Params(capital=3000, variation="m5_w_m15_0p1_BTC",
                         mode="general", ma_filter=False))
    k = res["kpis"]
    print("trades:", k["total_trades"], "| net:", k["net_pnl"],
          "| expected from report ≈ -3000.91")
    res2 = analyze(Params(capital=100_000, variation="m30_0p1_ETH",
                          mode="profitable", ma_filter=True,
                          ma_period=20, ma_condition="touch"))
    k2 = res2["kpis"]
    print("m30 ETH profitable | live:", k2["live_trades"], "virtual:", k2["virtual_trades"],
          "| switches:", k2["mode_switches"], "| net:", k2["net_pnl"])
