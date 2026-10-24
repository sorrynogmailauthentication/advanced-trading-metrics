"""
1m closes + volume → ATR_Monitor.xlsx, одна разовая выгрузка за день.
Volume — той же минуты, тем же порядком колонок (time, бенчмарк, тикеры),
как отдаёт API (лоты, без умножения).

Колонка time — чистое время Excel (1899-12-30 hh:mm), как на листе RVOL.
Дата сессии пишется отдельно в CI3.

  NASDAQ vs QQQ → Лист1 / closes N:AE, volume с BO
  NYSE   vs SPY → Лист3 / closes N:AA, volume с BC

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

# sheet, close_start, close_end, vol_start, benchmark, tickers
BLOCKS = (
    ("Лист1", "N", "AE", "CE", "QQQ", NASDAQ_TICKERS),
    ("Лист3", "N", "AA", "BO", "SPY", NYSE_TICKERS),
)

HOST = "127.0.0.1"
PORT = 7496  # 7496=live TWS, 7497=paper/Gateway
CLIENT_ID = 92
# Сколько reqHistoricalData одновременно. Поднимать, пока IBKR не начнёт отклонять.
WAVE_SIZE = 20

ET = ZoneInfo("America/New_York")
DEFAULT_START = dtime(9, 30)
DEFAULT_END = dtime(9, 40)
# End(xlDown) с заголовка при пустой строке 2 уезжает на 1048576.
MAX_DATA_ROW = 1500


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


def _col_idx(col: str) -> int:
    n = 0
    for c in col.upper():
        n = n * 26 + (ord(c) - ord("A") + 1)
    return n


def _idx_col(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(ord("A") + r) + s
    return s


def _end_col(start_col: str, n_cols: int) -> str:
    return _idx_col(_col_idx(start_col) + n_cols - 1)


def _series_from_bars(
    bars,
    *,
    session_day: date,
    t0: dtime,
    t1: dtime,
) -> dict[datetime, tuple[float, float]]:
    """Минута → (close, volume) одного и того же 1m бара."""
    out: dict[datetime, tuple[float, float]] = {}
    for bar in bars:
        key = _norm_bar_ts(bar.date)
        if key.date() != session_day:
            continue
        if t0 <= key.time() <= t1:
            vol = float(bar.volume)
            out[key] = (float(bar.close), int(vol) if vol.is_integer() else vol)
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
) -> dict[str, dict[datetime, tuple[float, float]]]:
    # Ошибка запроса (в т.ч. pacing 162) приходит исключением по этому символу,
    # а не пустым набором баров.
    ib.RaiseRequestErrors = True
    contracts = [Stock(sym, "SMART", "USD") for sym in ordered]
    qualified = ib.run(
        asyncio.gather(
            *(ib.qualifyContractsAsync(c) for c in contracts),
            return_exceptions=True,
        )
    )
    series: dict[str, dict[datetime, tuple[float, float]]] = {}
    ready: list[Stock] = []
    for sym, contract, result in zip(ordered, contracts, qualified):
        if isinstance(result, Exception) or not result:
            err = result if isinstance(result, Exception) else "contract not qualified"
            print(f"{sym} FAIL: {err}")
            series[sym] = {}
            continue
        ready.append(contract)

    n_waves = (len(ready) + WAVE_SIZE - 1) // WAVE_SIZE if ready else 0
    for wave, start in enumerate(range(0, len(ready), WAVE_SIZE), start=1):
        chunk = ready[start : start + WAVE_SIZE]
        names = ", ".join(c.symbol for c in chunk)
        print(f"wave {wave}/{n_waves} ({len(chunk)}): {names}", flush=True)
        fetched = ib.run(
            asyncio.gather(
                *(
                    ib.reqHistoricalDataAsync(
                        contract,
                        endDateTime=end_dt,
                        durationStr=duration,
                        barSizeSetting="1 min",
                        whatToShow="TRADES",
                        useRTH=True,
                        formatDate=1,
                        keepUpToDate=False,
                    )
                    for contract in chunk
                ),
                return_exceptions=True,
            )
        )
        for contract, bars in zip(chunk, fetched):
            sym = contract.symbol
            if isinstance(bars, Exception):
                print(f"  {sym} FAIL: {bars}")
                series[sym] = {}
                continue
            series[sym] = _series_from_bars(
                bars, session_day=session_day, t0=t0, t1=t1
            )
            print(f"  {sym} {len(series[sym])} bars")
    return series


def _row_for_minute(
    ts: datetime,
    ordered: list[str],
    series: dict[str, dict[datetime, tuple[float, float]]],
    field: int,
) -> list:
    row = [datetime(1899, 12, 30, ts.hour, ts.minute)]
    for sym in ordered:
        bar = series.get(sym, {}).get(ts)
        row.append(None if bar is None else bar[field])
    return row


def _sheet(name: str):
    wb = xw.Book(BOOK)
    return wb.sheets[name]


def _old_last_row(sheet, start_col: str) -> int:
    """Последняя занятая строка колонки, без прыжка на конец листа."""
    start = sheet.range(f"{start_col}1")
    if start.value is None:
        return 1
    try:
        last = start.end("down").row
    except Exception:
        return 1
    if 1 <= last <= MAX_DATA_ROW:
        return last
    vals = sheet.range(f"{start_col}1:{start_col}{MAX_DATA_ROW}").value
    if not isinstance(vals, list):
        return 1
    last_filled = 1
    for i, v in enumerate(vals, start=1):
        if v not in (None, ""):
            last_filled = i
    return last_filled


def _pause_excel(sheet) -> tuple:
    app = sheet.book.app
    state = (app.screen_updating, app.enable_events, app.calculation)
    app.screen_updating = False
    app.enable_events = False
    app.calculation = "manual"
    return state


def _resume_excel(sheet, state: tuple) -> None:
    screen_updating, enable_events, calculation = state
    app = sheet.book.app
    app.calculation = calculation
    app.enable_events = enable_events
    app.screen_updating = screen_updating


def _write_block(
    sheet,
    ordered: list[str],
    rows: list[list],
    start_col: str,
    end_col: str,
    *,
    volume: bool = False,
) -> None:
    """Одна запись блока. Хвост старых строк ниже новых данных очищаем отдельно."""
    header = ["time", *ordered]
    n_cols = _col_idx(end_col) - _col_idx(start_col) + 1
    if len(header) > n_cols:
        raise SystemExit(
            f"Слишком много колонок ({len(header)}): блок {start_col}:{end_col} "
            f"вмещает {n_cols} (time + бенчмарк + тикеры)"
        )
    n_rows = 1 + len(rows)
    old_last = _old_last_row(sheet, start_col)
    sheet.range(f"{start_col}1").value = [header] + rows
    sheet.range(f"{start_col}1").number_format = "@"
    if n_rows >= 2:
        sheet.range(f"{start_col}2:{start_col}{n_rows}").number_format = "hh:mm"
        if volume and n_cols >= 2:
            data_col = _idx_col(_col_idx(start_col) + 1)
            sheet.range(f"{data_col}2:{end_col}{n_rows}").number_format = "#,##0"
    if old_last > n_rows:
        sheet.range(f"{start_col}{n_rows + 1}:{end_col}{old_last}").value = None


def _export_block(
    ib: IB,
    *,
    sheet_name: str,
    start_col: str,
    end_col: str,
    vol_start: str,
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

    close_rows = [_row_for_minute(ts, ordered, series, 0) for ts in all_times]
    vol_rows = [_row_for_minute(ts, ordered, series, 1) for ts in all_times]
    vol_end = _end_col(vol_start, 1 + len(ordered))
    sheet = _sheet(sheet_name)
    state = _pause_excel(sheet)
    try:
        _write_block(sheet, ordered, close_rows, start_col, end_col)
        _write_block(sheet, ordered, vol_rows, vol_start, vol_end, volume=True)
    finally:
        _resume_excel(sheet, state)
    print(
        f"Wrote {len(close_rows)} rows -> {BOOK}/{sheet_name} "
        f"closes {start_col}, volume {vol_start}:{vol_end}"
    )
    return len(close_rows)


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
        for sheet_name, start_col, end_col, vol_start, benchmark, tickers in BLOCKS:
            written += _export_block(
                ib,
                sheet_name=sheet_name,
                start_col=start_col,
                end_col=end_col,
                vol_start=vol_start,
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
    started = time.perf_counter()
    try:
        run_oneshot(date.fromisoformat(args.date), t0, t1)
    finally:
        elapsed = time.perf_counter() - started
        print(f"Done in {elapsed:.1f}s (wave {WAVE_SIZE})")


if __name__ == "__main__":
    main()
