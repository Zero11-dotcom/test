"""
exporter.py — builds the compiled report bundle:
  * Backtest_Report.pdf   (metrics + charts + heatmap, dark theme)
  * TradeLog.xlsx         (processed trade log incl. LIVE/VIRTUAL + split chunks)
bundled into a single ZIP for download.
"""

from __future__ import annotations

import io
import zipfile
from datetime import datetime

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.patches import FancyBboxPatch
import numpy as np
import pandas as pd

from processor import MONTHS, build_trade_log

# dark palette (matches the dashboard)
BG = "#0b0b10"
PANEL = "#14141c"
GRID = "#23232e"
TXT = "#e5e7eb"
SUB = "#8b8b98"
RED = "#ef4444"
CYAN = "#22d3ee"
AMBER = "#f59e0b"
GREEN = "#34d399"
VIRTUAL = "#f87171"

plt.rcParams.update({
    "text.color": TXT, "axes.labelcolor": SUB, "xtick.color": SUB,
    "ytick.color": SUB, "axes.edgecolor": GRID, "figure.facecolor": BG,
    "axes.facecolor": PANEL, "savefig.facecolor": BG,
})


def _usd(v, plus=False):
    if v is None:
        return "—"
    s = f"{v:,.2f}"
    if v > 0 and plus:
        s = "+" + s
    return "$" + s


def _pct(v, plus=True):
    if v is None:
        return "—"
    s = f"{v:,.2f}%"
    if v > 0 and plus:
        s = "+" + s
    return s


# --------------------------------------------------------------------------- #
# PDF
# --------------------------------------------------------------------------- #

def _page(w=11.69, h=8.27):
    return plt.Figure(figsize=(w, h))


def _header(fig, title, subtitle=""):
    fig.text(0.045, 0.955, "ZERO", fontsize=15, fontweight="bold", color=RED)
    fig.text(0.105, 0.955, "| BACKTEST ANALYTICS", fontsize=10, color=SUB)
    fig.text(0.045, 0.905, title, fontsize=17, fontweight="bold", color=TXT)
    if subtitle:
        fig.text(0.045, 0.868, subtitle, fontsize=9, color=SUB)
    fig.lines.append(plt.Line2D([0.045, 0.955], [0.885, 0.885],
                     transform=fig.transFigure, color=GRID, lw=1))


def _kpi_box(fig, x, y, w, h, label, value, sub, accent):
    fig.patches.append(FancyBboxPatch(
        (x, y), w, h, boxstyle="round,pad=0.004,rounding_size=0.012",
        transform=fig.transFigure, facecolor=PANEL, edgecolor=GRID, lw=1))
    fig.patches.append(plt.Rectangle((x, y + 0.004), 0.0035, h - 0.008,
        transform=fig.transFigure, facecolor=accent, edgecolor="none"))
    fig.text(x + 0.012, y + h - 0.032, label, fontsize=8, color=SUB)
    fig.text(x + 0.012, y + h - 0.075, value, fontsize=15, fontweight="bold", color=TXT)
    fig.text(x + 0.012, y + 0.022, sub, fontsize=7.5, color=SUB, wrap=True)


def _pdf_summary(pdf, a):
    fig = _page()
    p, k, m = a["params"], a["kpis"], a["meta"]
    cond = "—" if not p["ma_filter"] else (
        f"{p['ma_period']}×MA · " + ("Touching MA" if p["ma_condition"] == "touch" else "Close Below MA"))
    sub = (f"{p['variation_label']}   ·   {m['symbol']}   ·   "
           f"{'GENERAL MODE' if p['mode']=='general' else 'ONLY PROFITABLE HOUR + WEEKDAY'}   ·   "
           f"Exness Limit {'ON' if p['position_distribution'] else 'OFF'}   ·   MA Filter: {cond}")
    _header(fig, "Performance Summary", sub)

    # KPI grid 3x2
    x0, y0, w, h, gx, gy = 0.045, 0.40, 0.29, 0.185, 0.015, 0.045
    boxes = [
        ("NET P&L", _usd(k["net_pnl"], plus=True), f"{_pct(k['net_pnl_pct'])} on initial capital", GREEN if k["net_pnl"] >= 0 else RED),
        ("TOTAL TRADES", f"{k['total_trades']:,}", f"{k['live_trades']:,} live · {k['virtual_trades']:,} virtual", CYAN),
        ("WIN RATE", f"{k['win_rate']}%", f"{k['wins']} W / {k['losses']} L · PF {k['profit_factor']}", AMBER),
        ("MAX DRAWDOWN — MONTHLY", f"{k['max_dd_monthly']['pct']}%", 
         f"{_usd(k['max_dd_monthly']['abs'])} · {k['max_dd_monthly'].get('period','') or ''} "
         f"({k['max_dd_monthly'].get('peak_date','')} → {k['max_dd_monthly'].get('trough_date','')})", RED),
        ("MAX DRAWDOWN — YEARLY", f"{k['max_dd_yearly']['pct']}%",
         f"{_usd(k['max_dd_yearly']['abs'])} · {k['max_dd_yearly'].get('period','') or ''} "
         f"({k['max_dd_yearly'].get('peak_date','')} → {k['max_dd_yearly'].get('trough_date','')})", RED),
        ("MODE SWITCHES", str(k["mode_switches"]),
         f"live ↔ virtual · MA: {cond}", VIRTUAL),
    ]
    for i, (lab, val, s, acc) in enumerate(boxes):
        r, c = divmod(i, 3)
        _kpi_box(fig, x0 + c * (w + gx), y0 + (1 - r) * (h + gy), w, h, lab, val, s, acc)

    # config / period footer
    fig.text(0.045, 0.345, "REPORT CONFIGURATION", fontsize=10, fontweight="bold", color=RED)
    cfg_lines = [
        f"Initial capital {_usd(p['capital'])}  ·  Final equity {_usd(k['final_equity'])}  ·  "
        f"Period {m['first_trade']} → {m['last_trade']}",
        f"Source {m['source_file']} (report deposit {_usd(m['report_deposit'])})  ·  "
        f"Position sizing: Volume = Equity / ({m['margin_pct']:.2f} × Entry Price)",
        f"Max {m['max_lot']} lots per position → split into chunks with identical entry/exit  ·  "
        f"Exness limit {'ON' if p['position_distribution'] else 'OFF'}",
        f"Gross profit {_usd(k['gross_profit'])}  ·  Gross loss {_usd(k['gross_loss'])}  ·  "
        f"Avg trade {_usd(k['avg_trade'])}  ·  Best {_usd(k['best_trade'])}  ·  Worst {_usd(k['worst_trade'])}",
    ]
    fig.text(0.045, 0.265, "\n".join(cfg_lines), fontsize=8.2, color=SUB, linespacing=2.0)

    # mini equity chart on summary page
    ax = fig.add_axes([0.045, 0.055, 0.91, 0.135])
    s = a["series"]
    x = np.arange(len(s["original"]))
    ax.plot(x, s["original"], color=CYAN, lw=1.2)
    ax.set_facecolor(PANEL)
    ax.grid(color=GRID, lw=0.5, alpha=0.6)
    ax.tick_params(labelsize=7)
    ax.set_title("Official equity curve (live trades)", fontsize=8, color=SUB, loc="left", pad=3)
    pdf.savefig(fig)
    plt.close(fig)


def _pdf_equity(pdf, a):
    fig = _page()
    p, s = a["params"], a["series"]
    _header(fig, "Equity Curves",
            f"{p['variation_label']} · X-axis: trade close time · Y-axis: equity USD")
    ax = fig.add_axes([0.055, 0.09, 0.91, 0.74])
    x = np.arange(len(s["labels"]))
    ax.plot(x, s["original"], color="#64748b", lw=1.0, alpha=0.8, label="Original (no filter)")
    flags = s["live_flags"]
    if p["ma_filter"]:
        # live segments solid cyan, virtual dashed red
        xs, ys = [x[0]], [s["original"][0]]
        for i in range(1, len(x)):
            if flags[i] != flags[i - 1]:
                xs.append(x[i]); ys.append(s["original"][i])
                _plot_seg(ax, xs, ys, flags[i - 1])
                xs, ys = [x[i]], [s["original"][i]]
        _plot_seg(ax, xs + [x[-1]], ys + [s["original"][-1]], flags[-1])
        ma = [v if v is not None else np.nan for v in s["original_ma"]]
        ax.plot(x, ma, color=AMBER, lw=1.2, label=f"MA({p['ma_period']})")
        ax.plot([], [], color=CYAN, lw=1.4, label="Live trading")
        ax.plot([], [], color=VIRTUAL, lw=1.4, ls="--", label="Virtual trading")
    else:
        ax.plot(x, s["original"], color=CYAN, lw=1.2, label="Equity")
    ax.legend(loc="upper left", fontsize=8, frameon=False, labelcolor=TXT)
    ax.grid(color=GRID, lw=0.5, alpha=0.6)
    ax.set_ylabel("Equity (USD)", fontsize=8)
    step = max(1, len(x) // 10)
    ax.set_xticks(x[::step])
    ax.set_xticklabels([s["labels"][i][:10] for i in range(0, len(x), step)], fontsize=7)
    pdf.savefig(fig)
    plt.close(fig)


def _plot_seg(ax, xs, ys, flag):
    if flag == "live":
        ax.plot(xs, ys, color=CYAN, lw=1.4)
    else:
        ax.plot(xs, ys, color=VIRTUAL, lw=1.4, ls="--")


def _pdf_real(pdf, a):
    fig = _page()
    p, s = a["params"], a["series"]
    _header(fig, "Real-Only Equity (virtual trades excluded, stitched)",
            "Official account equity — the equity that would have been realized live")
    ax = fig.add_axes([0.055, 0.09, 0.91, 0.74])
    x = np.arange(len(s["real_labels"]))
    ax.plot(x, s["real_values"], color=GREEN, lw=1.3)
    ax.fill_between(x, p["capital"], s["real_values"], color=GREEN, alpha=0.06)
    ax.axhline(p["capital"], color=SUB, lw=0.7, ls=":")
    ax.grid(color=GRID, lw=0.5, alpha=0.6)
    ax.set_ylabel("Equity (USD)", fontsize=8)
    step = max(1, len(x) // 10)
    ax.set_xticks(x[::step])
    ax.set_xticklabels([s["real_labels"][i][:10] for i in range(0, len(x), step)], fontsize=7)
    pdf.savefig(fig)
    plt.close(fig)


def _pdf_heatmap(pdf, a):
    fig = _page()
    hm = a["heatmap"]
    _header(fig, "Calendar Heatmap — Monthly & Yearly P&L", a["params"]["variation_label"])
    years = hm["years"]
    ny = len(years)
    ax = fig.add_axes([0.05, 0.10, 0.83, 0.72])
    ax.set_facecolor(BG)
    max_abs = hm.get("max_abs", 1.0) or 1.0
    for yi, yrow in enumerate(years):
        y = ny - 1 - yi
        for c in yrow["months"]:
            x = c["month"] - 1
            v = c["pnl"]
            if c["state"] == "out":
                fc, txt, tc = "#101018", "", SUB
            elif abs(v) < 0.005:
                fc, txt, tc = "#16161f", "0", SUB
            else:
                al = 0.2 + 0.75 * min(1.0, abs(v) / max_abs)
                fc = (*matplotlib.colors.to_rgb(GREEN), al) if v > 0 else (*matplotlib.colors.to_rgb(RED), al)
                txt = f"{v/1000:,.0f}k" if abs(v) >= 1000 else f"{v:,.0f}"
                tc = "#0b0b10" if al > 0.72 else TXT
            ax.add_patch(plt.Rectangle((x, y), 0.94, 0.94,
                facecolor=fc, edgecolor=GRID if c["state"] != "out" else "#1a1a24", lw=1.0))
            if txt:
                ax.text(x + 0.47, y + 0.47, txt, ha="center", va="center", fontsize=6.2, color=tc)
        ax.text(-0.55, y + 0.47, str(yrow["year"]), ha="right", va="center", fontsize=9, color=TXT, fontweight="bold")
        tot = yrow["total"]
        ax.text(12.35, y + 0.47, f"{tot/1000:,.1f}k" if abs(tot) >= 1000 else f"{tot:,.0f}",
                ha="left", va="center", fontsize=8,
                color=GREEN if tot >= 0 else RED, fontweight="bold")
    ax.text(12.35, ny - 0.06, "YEAR", fontsize=8, color=SUB, fontweight="bold")
    for m in range(12):
        ax.text(m + 0.47, ny + 0.12, MONTHS[m], ha="center", va="center", fontsize=7.2, color=SUB)
    ax.text(-0.55, ny + 0.3, "MONTH", fontsize=8, color=SUB)
    ax.set_xlim(-1.4, 14.6); ax.set_ylim(-0.35, ny + 0.55)
    ax.axis("off")
    fig.text(0.05, 0.045, f"Grand total (live trades): {_usd(hm.get('grand_total', 0), plus=True)}   ·   "
             f"cell color intensity = |monthly P&L| (green = profit, red = loss). "
             f"Dimmed cells are outside the backtest range.", fontsize=8, color=SUB)
    pdf.savefig(fig)
    plt.close(fig)


def _pdf_perf(pdf, a):
    fig = _page()
    perf = a["perf"]
    _header(fig, "Performance Profile — Weekday & Hour of Entry", a["params"]["variation_label"])
    ax1 = fig.add_axes([0.06, 0.12, 0.40, 0.68])
    wd = perf["weekday"]
    vals = [w["pnl"] for w in wd]
    ax1.bar(range(7), vals, color=[GREEN if v >= 0 else RED for v in vals], width=0.62)
    ax1.set_xticks(range(7)); ax1.set_xticklabels([w["day"] for w in wd], fontsize=8)
    ax1.grid(color=GRID, lw=0.5, axis="y", alpha=0.6)
    ax1.axhline(0, color=SUB, lw=0.6)
    ax1.set_title("P&L by weekday (entry time)", fontsize=9, color=SUB, loc="left")
    ax2 = fig.add_axes([0.56, 0.12, 0.40, 0.68])
    hr = perf["hour"]
    vals = [h["pnl"] for h in hr]
    ax2.bar(range(24), vals, color=[GREEN if v >= 0 else RED for v in vals], width=0.7)
    ax2.set_xticks(range(0, 24, 2)); ax2.set_xticklabels([f"{h:02d}" for h in range(0, 24, 2)], fontsize=7)
    ax2.grid(color=GRID, lw=0.5, axis="y", alpha=0.6)
    ax2.axhline(0, color=SUB, lw=0.6)
    ax2.set_title("P&L by hour of entry", fontsize=9, color=SUB, loc="left")
    pdf.savefig(fig)
    plt.close(fig)


def make_pdf(a) -> bytes:
    buf = io.BytesIO()
    with PdfPages(buf) as pdf:
        _pdf_summary(pdf, a)
        _pdf_equity(pdf, a)
        _pdf_real(pdf, a)
        _pdf_heatmap(pdf, a)
        _pdf_perf(pdf, a)
        d = pdf.infodict()
        d["Title"] = f"ZERO Backtest Report — {a['params']['variation_label']}"
        d["CreationDate"] = datetime.now().ctime()
    return buf.getvalue()


# --------------------------------------------------------------------------- #
# XLSX
# --------------------------------------------------------------------------- #

def make_xlsx(a) -> bytes:
    buf = io.BytesIO()
    p, k, m = a["params"], a["kpis"], a["meta"]
    try:
        wargs = {"engine": "xlsxwriter",
                 "engine_kwargs": {"options": {"constant_memory": True}}}
    except ImportError:  # pragma: no cover
        wargs = {"engine": "openpyxl"}
    with pd.ExcelWriter(buf, **wargs) as w:
        summary = pd.DataFrame([
            ("Report", f"ZERO Backtest — {p['variation_label']}"),
            ("Generated", datetime.now().strftime("%Y-%m-%d %H:%M")),
            ("Mode", "General" if p["mode"] == "general" else "Only Profitable Hour + Weekday"),
            ("Symbol", m["symbol"]),
            ("Source file", m["source_file"]),
            ("Initial Capital (USD)", p["capital"]),
            ("Final Equity (USD)", k["final_equity"]),
            ("Net P&L (USD)", k["net_pnl"]),
            ("Net P&L (Instrument %)", f"{k['net_pnl_pct']}%"),
            ("Total Trades", k["total_trades"]),
            ("  · Live", k["live_trades"]),
            ("  · Virtual", k["virtual_trades"]),
            ("Win Rate", f"{k['win_rate']}%"),
            ("Profit Factor", k["profit_factor"]),
            ("Gross Profit", k["gross_profit"]),
            ("Gross Loss", k["gross_loss"]),
            ("Max DD Monthly (%)", k["max_dd_monthly"]["pct"]),
            ("Max DD Monthly (USD)", k["max_dd_monthly"]["abs"]),
            ("Max DD Monthly Period", f"{k['max_dd_monthly'].get('period','')} "
                                      f"({k['max_dd_monthly'].get('peak_date','')} → {k['max_dd_monthly'].get('trough_date','')})"),
            ("Max DD Yearly (%)", k["max_dd_yearly"]["pct"]),
            ("Max DD Yearly (USD)", k["max_dd_yearly"]["abs"]),
            ("Max DD Yearly Period", f"{k['max_dd_yearly'].get('period','')} "
                                     f"({k['max_dd_yearly'].get('peak_date','')} → {k['max_dd_yearly'].get('trough_date','')})"),
            ("Mode Switches", k["mode_switches"]),
            ("MA Filter", "OFF" if not p["ma_filter"] else
                          f"ON · {p['ma_period']}×MA · " +
                          ("Touching MA" if p["ma_condition"] == "touch" else "Close Below MA")),
            ("Position Distribution (Exness)", "ON" if p["position_distribution"] else "OFF"),
            ("Sizing Rule", f"Volume = Equity / ({m['margin_pct']} × Entry Price), split at {m['max_lot']} lots"),
        ], columns=["Field", "Value"])
        summary.to_excel(w, sheet_name="Summary", index=False)

        build_trade_log(a).to_excel(w, sheet_name="TradeLog", index=False)

        hm = a["heatmap"]
        rows = []
        for y in hm["years"]:
            rows.append([y["year"]] + [c["pnl"] for c in y["months"]] + [y["total"]])
        monthly = pd.DataFrame(rows, columns=["Year"] + MONTHS + ["Total"])
        monthly.to_excel(w, sheet_name="MonthlyPnL", index=False)

        sw = pd.DataFrame(a["switches"])
        (sw if not sw.empty else pd.DataFrame(columns=["time", "type"])).to_excel(
            w, sheet_name="ModeSwitches", index=False)
    return buf.getvalue()


# --------------------------------------------------------------------------- #
# ZIP bundle
# --------------------------------------------------------------------------- #

def make_bundle(a) -> tuple[bytes, str]:
    """Returns (zip_bytes, filename)."""
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    fname = f"ZERO_report_{a['params']['variation']}_{stamp}.zip"
    zbuf = io.BytesIO()
    with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("Backtest_Report.pdf", make_pdf(a))
        z.writestr(f"TradeLog_{a['params']['variation']}.xlsx", make_xlsx(a))
    return zbuf.getvalue(), fname
