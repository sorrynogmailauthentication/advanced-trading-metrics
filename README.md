# Playbook

Personal **intraday trading toolkit**: Excel + Interactive Brokers (TWS) + TradingView screener scraping, with a small local web UI for **relative volume (RVOL)** history.

This is a workspace for live monitors and automation around an opening-range / RVOL workflow—not a hosted product.

---

## Architecture

```text
┌─────────────────┐     scrape :00      ┌──────────────┐
│ TradingView (TV)│ ──────────────────► │ Excel Лист1  │
│ screener (RVOL) │                     │ (Symbol,RVOL)│
└─────────────────┘                     └──────┬───────┘
                                               │ :03 logger
┌─────────────────┐     OCR :00                ▼
│ TWS Monitor (IB)│ ──────────────────► ┌──────────────┐
│ Favorites RVOL  │                     │  PostgreSQL  │
└─────────────────┘                     │  rvol_ticks  │
                                        └──────┬───────┘
┌─────────────────┐                            │
│ TWS RTD → Excel │  (prices / ORB levels)     ▼
│                 │                     ┌──────────────┐
└─────────────────┘                     │ Local charts │
                                        │ :8787 IB+TV  │
                                        └──────────────┘
```

| Layer | Role |
|--------|------|
| **Excel** | Live hub: IB RTD on **Лист2**, TV screener dump on **Лист1** |
| **Ticker list** | **Лист2 column `L`** (see `tikers.py`) |
| **PostgreSQL** | Append-only minute history (`source=tv` / `source=ib`) |
| **Local web app** | Dual RVOL charts (IB above, TradingView below) |

---

## Scripts

| Script | What it does |
|--------|----------------|
| [`tikers.py`](tikers.py) | Loads unique tickers from **Ticker list** |
| [`RVOL_TV_excel.py`](RVOL_TV_excel.py) | Selenium scrape of a TradingView screener; at **:00** writes Symbol + Rel vol to **Лист1** `G:H` (watchlist only) |
| [`rvol_logger.py`](rvol_logger.py) | At **:03** reads Лист1 RVOL → Postgres `source=tv` |
| [`rvol_ib_ui.py`](rvol_ib_ui.py) | At **:00** OCR of TWS Favorites (FIN INSTR \| RVOL) → Postgres `source=ib` |
| [`rvol_ib_api.py`](rvol_ib_api.py) | Optional: compute 10m TOD RVOL via IB API (no UI); same `source=ib` |
| [`rvol_db.py`](rvol_db.py) | Postgres helpers; schema created automatically |
| [`rvol_app.py`](rvol_app.py) | FastAPI charts: `http://127.0.0.1:8787/t/TICKER` |
| [`ORB_snapshot_excel.py`](ORB_snapshot_excel.py) | At **9:36** (configurable) freezes live High/Low → ORB columns |
| [`ATR_Yahoo_pull.py`](ATR_Yahoo_pull.py) | Pulls ATR-style stats from Yahoo for the Excel ticker list |

---

## Requirements

- Windows (Excel + TWS + OCR path)
- Python 3.11+ (3.14 works with `ib_async` for the API script)
- Open **`ATR_Monitor.xlsx`** in Excel
- Local **PostgreSQL** database (e.g. `playbook`)
- Chrome profile for TradingView scrape (undetected-chromedriver)
- TWS running for IB RTD / optional API / OCR monitor

```bash
pip install -r requirements-rvol.txt
# plus: undetected-chromedriver, selenium, yfinance as needed for other scripts
```

---

## Configuration

Copy env template and edit locally (never commit secrets):

```bash
cp .env.example .env
```

Important variables:

| Variable | Purpose |
|----------|---------|
| `DATABASE_URL` | Postgres connection string |
| `SCREENER_URL` | Your TradingView screener URL |
| `RVOL_LOG_GRACE` | Second-of-minute for TV logger (default `3`) |
| `TWS_WINDOW_TITLE` | Exact TWS window title for OCR attach |
| `TWS_OCR_REGION` | Screen crop `left,top,width,height` after `--calibrate` |
| `IB_HOST` / `IB_PORT` / `IB_CLIENT_ID` | TWS socket API (no IB login in `.env`) |

---

## Typical session

1. Start Postgres; open `ATR_Monitor.xlsx` (tickers in **Лист2!L**).
2. Start TWS (RTD + Favorites monitor visible for OCR).
3. Terminals:
   ```bash
   python RVOL_TV_excel.py          # TV → Excel :00
   python rvol_logger.py            # Excel → DB  :03 (tv)
   python rvol_ib_ui.py --calibrate # once: select Favorites region
   python rvol_ib_ui.py             # OCR → DB    :00 (ib)
   uvicorn rvol_app:app --port 8787
   ```
4. Charts: [http://127.0.0.1:8787/t/AAPL](http://127.0.0.1:8787/t/AAPL)  
   Excel: `=ГИПЕРССЫЛКА("http://127.0.0.1:8787/t/"&L2;"RVOL")`

**Add a ticker:** put it in **Лист2!L**. No DB registration step.  
Restart `rvol_ib_ui.py` (or `--reload-tickers`) so OCR picks up new symbols. TV logger picks them up on the next cycle.

Schema is created automatically via `ensure_schema()`—no separate migrate script.

---

## RVOL web UI

- **Top chart:** IB (`source=ib`)  
- **Bottom chart:** TradingView (`source=tv`)  
- Shared time axis; missing minutes → empty bars  
- Default Y scale **0–3**, expands only if a visible bar exceeds 3  
- Zoom (wheel on bar), shared scrollbar; default view = last 60 minutes  
- Refresh every 60s; paused when the tab is hidden  

---

## ORB snapshot

`ORB_snapshot_excel.py` waits until a configured clock time, then copies live High/Low into static ORB columns on **Лист2** for breakout highlighting in Excel.

---

## Privacy & what not to publish

| Keep private | Why |
|--------------|-----|
| **`.env`** | Real DB passwords (already in `.gitignore`) |
| **`chrome_profile_1/`** | Browser cookies / TradingView session |
| **`tws_ocr_debug.png`**, dumps | Screenshots of live blotter |
| **Personal strategy notes** (`ORB.md`, etc.) | Your edge / process—already tracked; untrack before a public repo if you want them private |
| **Screener URL / broker window title** | Keep only in local `.env`, not in committed source |

Safe to share: scripts, `.env.example` with placeholders, this README.

Before a public push:

```bash
git status   # confirm .env and chrome_profile_* are not staged
# optional: git rm --cached ORB.md
```

If a real password ever landed in git history, rotate it and scrub history.

---

## Disclaimer

For personal research and tooling only. Not financial advice. Market data, scraping, and broker APIs are subject to their terms of use and your data subscriptions.
