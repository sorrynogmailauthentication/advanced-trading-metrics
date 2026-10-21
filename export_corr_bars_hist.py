"""
1m closes → ATR_Monitor.xlsx, одна разовая выгрузка за день:

  NASDAQ vs QQQ → Лист1 / колонки N:AH
  NYSE   vs SPY → Лист3 / колонки N:AA

  python export_corr_bars_hist.py --date 2026-10-13
  python export_corr_bars_hist.py --date 2026-10-13 --start 09:30 --end 16:00
"""

from __future__ import annotations

import argparse
import asyncio
import time
from datetime import date, datetime, time as dtime
from zoneinfo import ZoneInfo

import xlwings as xw
from ib_insync import IB, Stock

# Python 3.12+: нет default event loop в MainThread
try:
    asyncio.get_event_loop()
except RuntimeError:
    asyncio.set_event_loop(asyncio.new_event_loop())

# NASDAQ vs QQQ → Лист1 N:AH. Без TSM и HOOD.
NASDAQ_TICKERS = [
    "AAPL",
    "AMD",
    "AMZN",
    "AVGO",
    "CRWD",
    "GOOGL",
    "INTC",
    "META",
    "MRVL",
    "MSFT",
    "MU",
    "NFLX",
    "NVDA",
    "SNDK",
    "SPCX",
    "TSLA",
]

# NYSE vs SPY → Лист3 N:AA.
NYSE_TICKERS = [
    "PLTR",
    "PANW",
    "LLY",
    "WMT",
    "JPM",
    "AMAT",
    "XOM",
    "V",
    "BAC",
    "DELL",
    "CRM",
    "ORCL",
]

BOOK = "ATR_Monitor.xlsx"

# sheet, start_col, end_col, benchmark, tickers
BLOCKS = (
    ("Лист1", "N", "AE", "QQQ", NASDAQ_TICKERS),
    ("Лист3", "N", "AA", "SPY", NYSE_TICKERS),
)

HOST = "127.0.0.1"
PORT = 7496  # 7496=live TWS, 7497=paper/Gateway
CLIENT_ID = 92

ET = ZoneInfo("America/New_York")
DEFAULT_START = dtime(9, 30)
DEFAULT_END = dtime(9, 40)


def _parse_hhmm(text: str) -> dtime:
    return datetime.strptime(text.strip(), "%H:%M").time()


def _parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Export historic 1m IBKR bars to Excel Лист3")
    p.add_argument(
        "--date",
        required=True,
        metavar="YYYY-MM-DD",
        help="Session date in ET. Example: 2026-10-13",
    )
    p.add_argument(
        "--start",
        default=DEFAULT_START.strftime("%H:%M"),
        help=f"Window start ET HH:MM (default {DEFAULT_START.strftime('%H:%M')})",
    )
    p.add_argument(
        "--end",
        default=DEFAULT_END.strftime("%H:%M"),
        help=f"Window end ET HH:MM inclusive (default {DEFAULT_END.strftime('%H:%M')})",
    )
    return p.parse_args()


def _ordered_symbols(benchmark: str, tickers: list[str]) -> list[str]:
    names = [s.strip().upper() for s in tickers if str(s).strip()]
    if not names:
        raise SystemExit(f"Список для {benchmark} пуст — заполни его в export_corr_bars_hist.py")
    ordered: list[str] = []
    seen: set[str] = set()
    for s in [benchmark, *names]:
        if s not in seen:
            seen.add(s)
            ordered.append(s)
    return ordered


def _connect() -> IB:
    ib = IB()
    for port in (PORT, 7497, 7496):
        try:
            ib.connect(HOST, port, clientId=CLIENT_ID, readonly=True, timeout=5)
            print(f"IBKR connected: {HOST}:{port}")
            return ib
        except Exception as e:
            print(f"port {port}: {e}")
    raise RuntimeError("Не удалось подключиться к TWS/Gateway (7496/7497)")


def _ib_end_datetime(session_day: date, *, is_today: bool) -> str:
    if is_today:
        return ""
    return f"{session_day.strftime('%Y%m%d')} 16:05:00 US/Eastern"


def _norm_bar_ts(ts) -> datetime:
    if isinstance(ts, str):
        ts = datetime.fromisoformat(ts)
    if ts.tzinfo is None:
        ts = ts.replace(tzinfo=ET)
    else:
        ts = ts.astimezone(ET)
    return ts.replace(second=0, microsecond=0)


def _fetch_closes(
    ib: IB,
    symbol: str,
    *,
    session_day: date,
    end_dt: str,
    t0: dtime,
    t1: dtime,
    duration: str = "1 D",
) -> dict[datetime, float]:
    contract = Stock(symbol, "SMART", "USD")
    ib.qualifyContracts(contract)
    bars = ib.reqHistoricalData(
        contract,
        endDateTime=end_dt,
        durationStr=duration,
        barSizeSetting="1 min",
        whatToShow="TRADES",
        useRTH=True,
        formatDate=1,
        keepUpToDate=False,
    )
    out: dict[datetime, float] = {}
    for bar in bars:
        key = _norm_bar_ts(bar.date)
        if key.date() != session_day:
            continue
        if t0 <= key.time() <= t1:
            out[key] = float(bar.close)
    return out


def _fetch_all(
    ib: IB,
    ordered: list[str],
    *,
    session_day: date,
    end_dt: str,
    t0: dtime,
    t1: dtime,
    duration: str = "1 D",
) -> dict[str, dict[datetime, float]]:
    series: dict[str, dict[datetime, float]] = {}
    for i, sym in enumerate(ordered):
        try:
            print(f"[{i + 1}/{len(ordered)}] {sym} ...", end=" ", flush=True)
            series[sym] = _fetch_closes(
                ib,
                sym,
                session_day=session_day,
                end_dt=end_dt,
                t0=t0,
                t1=t1,
                duration=duration,
            )
            print(f"{len(series[sym])} bars")
        except Exception as e:
            print(f"FAIL: {e}")
            series[sym] = {}
        time.sleep(0.7)
    return series


def _row_for_minute(
    ts: datetime, ordered: list[str], series: dict[str, dict[datetime, float]]
) -> list:
    row = [ts.replace(tzinfo=None)]
    for sym in ordered:
        row.append(series.get(sym, {}).get(ts))
    return row


def _sheet(name: str):
    wb = xw.Book(BOOK)
    return wb.sheets[name]


def _old_last_row(sheet, start) -> int:
    if start.value is None:
        return 1
    try:
        return start.end("down").row
    except Exception:
        return 1


def _wipe_block(sheet, ordered: list[str], start_col: str, end_col: str) -> None:
    """Затираем только start:end (без clear_contents на огромном range)."""
    start = sheet.range(f"{start_col}1")
    old_last = _old_last_row(sheet, start)
    wipe_rows = max(old_last, 2)
    rng = sheet.range(f"{start_col}1:{end_col}{wipe_rows}")
    n_cols = rng.columns.count
    rng.value = [[None] * n_cols for _ in range(wipe_rows)]
    header = ["time", *ordered]
    if len(header) > n_cols:
        raise SystemExit(
            f"Слишком много колонок ({len(header)}): блок {start_col}:{end_col} "
            f"вмещает {n_cols} (time + бенчмарк + тикеры)"
        )
    start.value = [header]
    sheet.range(f"{start_col}1").number_format = "@"


def _write_block(sheet, ordered: list[str], rows: list[list], start_col: str) -> None:
    """Полная перезапись после wipe."""
    start = sheet.range(f"{start_col}1")
    start.value = [["time", *ordered]] + rows
    n_rows = 1 + len(rows)
    if n_rows >= 2:
        sheet.range(f"{start_col}2:{start_col}{n_rows}").number_format = "hh:mm"


def _export_block(
    ib: IB,
    *,
    sheet_name: str,
    start_col: str,
    end_col: str,
    benchmark: str,
    tickers: list[str],
    session_day: date,
    end_dt: str,
    t0: dtime,
    t1: dtime,
) -> int:
    ordered = _ordered_symbols(benchmark, tickers)
    print(f"{sheet_name} {benchmark}: {', '.join(ordered)}")
    series = _fetch_all(
        ib, ordered, session_day=session_day, end_dt=end_dt, t0=t0, t1=t1
    )
    all_times = sorted({t for d in series.values() for t in d})
    if not all_times:
        print(
            f"Нет баров {benchmark} за {session_day} "
            f"{t0.strftime('%H:%M')}-{t1.strftime('%H:%M')} ET."
        )
        return 0

    rows = [_row_for_minute(ts, ordered, series) for ts in all_times]
    sheet = _sheet(sheet_name)
    _wipe_block(sheet, ordered, start_col, end_col)
    _write_block(sheet, ordered, rows, start_col)
    print(f"Wrote {len(rows)} rows -> {BOOK}/{sheet_name} from {start_col}")
    return len(rows)


def run_oneshot(session_day: date, t0: dtime, t1: dtime) -> None:
    today_et = datetime.now(ET).date()
    is_today = session_day == today_et
    end_dt = _ib_end_datetime(session_day, is_today=is_today)

    print(
        f"One-shot {session_day.isoformat()} ET "
        f"{t0.strftime('%H:%M')}-{t1.strftime('%H:%M')} "
        f"(endDateTime={repr(end_dt) if end_dt else 'now'})"
    )

    ib = _connect()
    written = 0
    try:
        for sheet_name, start_col, end_col, benchmark, tickers in BLOCKS:
            written += _export_block(
                ib,
                sheet_name=sheet_name,
                start_col=start_col,
                end_col=end_col,
                benchmark=benchmark,
                tickers=tickers,
                session_day=session_day,
                end_dt=end_dt,
                t0=t0,
                t1=t1,
            )
    finally:
        ib.disconnect()

    if not written:
        raise SystemExit(
            f"Нет баров за {session_day} {t0.strftime('%H:%M')}-{t1.strftime('%H:%M')} ET."
        )


def main():
    args = _parse_args()
    t0 = _parse_hhmm(args.start)
    t1 = _parse_hhmm(args.end)
    if t1 < t0:
        raise SystemExit("--end must be >= --start")
    run_oneshot(date.fromisoformat(args.date), t0, t1)


if __name__ == "__main__":
    main()
