# ZERO — Backtest Analysis Dashboard

A web-based backtest analysis dashboard that processes raw **MQL5 Strategy Tester
reports** (`.xlsx`, Exness) and outputs performance KPIs, interactive equity
curves with an X-MA virtual-mode filter, a monthly P&L calendar heatmap, and
one-click PDF / Excel exports.

Dark "ZERO" theme in Tailwind CSS, matching the supplied UI references.

```
backtest-dashboard/
├── app.py           Flask backend (API + static server)      ← run this
├── processor.py     parsing + sizing + virtual-mode engine
├── exporter.py      PDF (matplotlib) + XLSX (xlsxwriter) exports
└── static/
    └── index.html   frontend (Tailwind CDN + Chart.js)
```

## Quick start

```bash
# 1. dependencies (Python 3.10+)
pip install flask pandas numpy openpyxl matplotlib xlsxwriter

# 2. run
python app.py
# → open http://127.0.0.1:5000
```

The dashboard auto-compiles **Variation 1 · General Mode** on load.
Press **Compile** after changing any sidebar parameter; downloads appear at the
bottom card.

## Linking the frontend to the backend data pipeline

The frontend is a single `static/index.html` served by Flask, so everything is
same-origin — no CORS setup needed. The contract:

| Endpoint            | Method | Body (JSON)                                            | Returns |
|---------------------|--------|--------------------------------------------------------|---------|
| `/api/variations`   | GET    | —                                                      | the 14 variations + 2 modes (fills the dropdown) |
| `/api/analyze`      | POST   | params (below)                                          | KPIs, 3 chart series, heatmap, weekday/hour perf |
| `/api/export/xlsx`  | POST   | params                                                  | `TradeLog_<variation>.xlsx` download |
| `/api/export/pdf`   | POST   | params                                                  | `BacktestReport_<variation>.pdf` download |
| `/api/export`       | POST   | params                                                  | both files zipped |

Params object (exactly what the sidebar holds):

```json
{
  "capital": 100000,
  "variation": "m5_w_m15_0p1_BTC",
  "mode": "general",                  // "general" | "profitable"
  "position_distribution": true,       // Exness 200-lot split ON/OFF
  "ma_filter": true,                   // P&L X MA filter ON/OFF
  "ma_period": 50,                     // X
  "ma_condition": "touch"              // "touch" | "close_below"
}
```

Want to split the frontend from Flask (e.g. serve it via nginx/Vite)? Point its
`fetch()` calls at the backend host and add CORS:

```python
# app.py
from flask_cors import CORS
CORS(app, resources={r"/api/*": {"origins": "*"}})
```

### Data location

`processor.py → CONFIG["DATA_ROOT"]` points at
`G:\EXNESS_backtesting_reports`. Two mode folders are expected:

* `GENERAL REPORT\full mql5 report\<id>_fullreport.xlsx`
* `ONLY PROFITABLE HOUR + PROFITABLE WEEKDAYS\full mql5 report\<id>_onlyprofitable.xlsx`

where `<id>` is one of the 14 variation ids in `VARIATIONS`
(e.g. `m5_w_m15_0p1_BTC`, `m30_0p1_ETH`). Change `DATA_ROOT` if the reports move.

## How the numbers are computed

**Parsing** — the MQL5 *Deals* table (`Time/Deal/Symbol/Type/Direction/Volume/
Price/Commission/Swap/Profit/Balance/Comment`) is read; `in`/`out` deal pairs are
stitched into round-trip trades with entry/exit time & price, volume, and net P&L
(balance delta, so commissions & swaps are included). Parsed per-trade net P&L
always reconciles with the report's own final balance.

**User-capital scaling** — reports were generated with different deposits
($3,000 / $100,000). Every trade is replayed as a *return on equity*
`r = pnl_raw / balance_before_raw`, applied to the user's running equity:

```
pnl_user_t = r_t × equity_user_before_t
```

This reproduces the report's own equity curve exactly, scaled to the Initial
Capital input, and compounds position sizing dynamically.

**Position Distribution (Exness limit)** — ON: `Volume = Equity / (0.15 × Entry
Price)` (Equity = 15 % of notional), split into chunks ≤ 200 lots
(`530.15 → 200 + 200 + 130.15`) that share identical entry/exit times. The lot
bookkeeping appears in the exported trade log; P&L is unchanged by splitting.
OFF: lots scale proportionally to the user's equity vs the report's, no split.

**X-MA Virtual Mode state machine** — the MA is computed on the *underlying*
equity curve (all trades). Per trade, closing time order:

1. `touch`  → equity ≤ MA triggers virtual;  `close_below` → equity < MA.
2. Once virtual, live trading resumes only when equity **closes back above** the MA.
3. Virtual trades are plotted (red dashed) but **excluded** from official
   Net P&L, win rate, drawdowns, and the heatmap.

**Drawdowns** — max peak-to-trough inside any single calendar month / year of the
official (live-only) curve, reported in USD and % with peak → trough dates.

## Export contents

* **PDF** — 5 pages: performance summary (KPI grid), equity curves (live/virtual/MA),
  real-only curve, calendar heatmap (Jan 2021 → last date), weekday/hour profile.
* **XLSX** — `Summary` (all params + KPIs), `TradeLog` (one row per 200-lot chunk
  with `Mode = LIVE/VIRTUAL`, `Chunk k/n`, prices, P&L, equity after, MA value),
  `MonthlyPnL` (year × month matrix), `ModeSwitches` (timestamped log).
