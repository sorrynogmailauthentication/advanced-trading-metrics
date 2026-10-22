"""
RTH 1m volume → rvol_volume.sqlite.

Копия export_rvol_volume.py: те же минуты 09:30..15:59 ET, но в базу,
чтобы повторный запуск не качал уже полный день.

День качаем снова, если по тикеру в базе больше 0 и меньше 390 баров
(обрыв, раннее закрытие, сегодняшний неполный день). 0 баров — дня нет,
тоже качаем. 390 и больше — пропускаем.

Тикеры без --symbol: QQQ, список NASDAQ, SPY, список NYSE
из export_corr_bars_hist.py.

  python export_rvol_volume_db.py
  python export_rvol_volume_db.py --symbol NVDA --days 90
  python export_rvol_volume_db.py --to-excel
  python export_rvol_volume_db.py --to-excel --symbol NVDA

--to-excel пишет лист с именем тикера: 90 последних непустых дней,
сетка как у export_rvol_volume.py. К IBKR не подключается.
"""

from __future__ import annotations

import argparse
import asyncio
import sqlite3
import time
from datetime import date, datetime, timedelta, time as dtime
from pathlib import Path
from zoneinfo import ZoneInfo

import xlwings as xw
from ib_insync import IB, Stock

from export_corr_bars_hist import NASDAQ_TICKERS, NYSE_TICKERS

try:
    asyncio.get_event_loop()
except RuntimeError:
    asyncio.set_event_loop(asyncio.new_event_loop())

LOOKBACK_DAYS = 90
EXCEL_DAYS = 90
FULL_DAY_BARS = 390  # 09:30..15:59

DB_PATH = Path(__file__).with_name("rvol_volume.sqlite")
BOOK = "ATR_Monitor.xlsx"
CLEAR_COL = "CZ"
RTH_START = dtime(9, 30)
RTH_END = dtime(15, 59)

HOST = "127.0.0.1"
PORT = 7496  # 7496=live TWS, 7497=paper/Gateway
CLIENT_ID = 94
PACE_SEC = 10
HIST_TIMEOUT = 20

ET = ZoneInfo("America/New_York")


def _all_symbols() -> list[str]:
    ordered: list[str] = []
    seen: set[str] = set()
    for raw in ["QQQ", *NASDAQ_TICKERS, "SPY", *NYSE_TICKERS]:
        symbol = str(raw).strip().upper()
        if symbol and symbol not in seen:
            seen.add(symbol)
            ordered.append(symbol)
    return ordered


def _parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Store RTH 1m volume in SQLite")
    p.add_argument(
        "--symbol",
        action="append",
        help="One ticker. Repeat for several. Default: all lists from export_corr_bars_hist.py",
    )
    p.add_argument(
        "--days",
        type=int,
        default=LOOKBACK_DAYS,
        help=f"Calendar lookback through today ET (default {LOOKBACK_DAYS})",
    )
    p.add_argument(
        "--to-excel",
        action="store_true",
        help=f"Write last {EXCEL_DAYS} non-empty days from the db to a sheet named as the ticker",
    )
    return p.parse_args()


def _connect_ib() -> IB:
    ib = IB()
    for port in (PORT, 7497, 7496):
        try:
            ib.connect(HOST, port, clientId=CLIENT_ID, readonly=True, timeout=5)
            ib.RequestTimeout = HIST_TIMEOUT
            print(f"IBKR connected: {HOST}:{port}", flush=True)
            return ib
        except Exception as e:
            print(f"port {port}: {e}")
    raise RuntimeError("Не удалось подключиться к TWS/Gateway (7496/7497)")


def _db() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS minute_volume (
            symbol TEXT NOT NULL,
            session_date TEXT NOT NULL,
            minute TEXT NOT NULL,
            volume REAL NOT NULL,
            PRIMARY KEY (symbol, session_date, minute)
        )
        """
    )
    conn.commit()
    return conn


def _bar_count(conn: sqlite3.Connection, symbol: str, session_day: date) -> int:
    row = conn.execute(
        """
        SELECT COUNT(*) FROM minute_volume
        WHERE symbol = ? AND session_date = ?
        """,
        (symbol, session_day.isoformat()),
    ).fetchone()
    return int(row[0])


def _replace_day(
    conn: sqlite3.Connection,
    symbol: str,
    session_day: date,
    volume: dict[dtime, float],
) -> None:
    day = session_day.isoformat()
    conn.execute(
        "DELETE FROM minute_volume WHERE symbol = ? AND session_date = ?",
        (symbol, day),
    )
    conn.executemany(
        """
        INSERT INTO minute_volume (symbol, session_date, minute, volume)
        VALUES (?, ?, ?, ?)
        """,
        [
            (symbol, day, minute.strftime("%H:%M"), vol)
            for minute, vol in sorted(volume.items())
        ],
    )
    conn.commit()


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


def _rth_minutes() -> list[dtime]:
    start = datetime.combine(date.min, RTH_START)
    end = datetime.combine(date.min, RTH_END)
    out: list[dtime] = []
    cur = start
    while cur <= end:
        out.append(cur.time())
        cur += timedelta(minutes=1)
    return out


def _session_days(today: date, lookback_days: int) -> list[date]:
    start = today - timedelta(days=lookback_days)
    days: list[date] = []
    cur = start
    while cur <= today:
        if cur.weekday() < 5:
            days.append(cur)
        cur += timedelta(days=1)
    return days


def _fetch_day_volume(
    ib: IB,
    contract,
    session_day: date,
    *,
    is_today: bool,
) -> dict[dtime, float]:
    end_dt = _ib_end_datetime(session_day, is_today=is_today)
    last_err: Exception | None = None
    bars = None
    for attempt in (1, 2):
        try:
            started = time.perf_counter()
            bars = ib.reqHistoricalData(
                contract,
                endDateTime=end_dt,
                durationStr="1 D",
                barSizeSetting="1 min",
                whatToShow="TRADES",
                useRTH=True,
                formatDate=1,
                keepUpToDate=False,
                timeout=HIST_TIMEOUT,
            )
            if not bars and (time.perf_counter() - started) >= HIST_TIMEOUT - 1:
                raise TimeoutError(
                    f"IBKR не ответил на historical data за {HIST_TIMEOUT}с "
                    f"({session_day.isoformat()})"
                )
            last_err = None
            break
        except Exception as e:
            last_err = e
            if attempt == 1 and "pacing" in str(e).lower():
                print("pacing, ждём 60с...")
                time.sleep(60)
                continue
            raise
    if last_err is not None:
        raise last_err

    out: dict[dtime, float] = {}
    for bar in bars or []:
        key = _norm_bar_ts(bar.date)
        if key.date() != session_day:
            continue
        if not (RTH_START <= key.time() <= RTH_END):
            continue
        out[key.time()] = float(bar.volume)
    return out


def _fill_symbol(
    ib: IB,
    conn: sqlite3.Connection,
    symbol: str,
    days: list[date],
    today: date,
) -> None:
    contract = Stock(symbol, "SMART", "USD")
    ib.qualifyContracts(contract)
    requested = 0
    for i, day in enumerate(days):
        have = _bar_count(conn, symbol, day)
        if have >= FULL_DAY_BARS:
            print(f"  {day.isoformat()} skip ({have} bars)")
            continue
        if requested:
            time.sleep(PACE_SEC)
        requested += 1
        why = "fetch" if have == 0 else f"refetch, have {have}"
        print(
            f"  [{i + 1}/{len(days)}] {day.isoformat()} {why} ...",
            end=" ",
            flush=True,
        )
        try:
            day_vol = _fetch_day_volume(
                ib, contract, day, is_today=(day == today)
            )
        except TimeoutError as e:
            print(f"FAIL: {e}")
            raise SystemExit(
                "HMDS не отдаёт бары. Перезапусти TWS и запусти скрипт снова."
            ) from e
        except Exception as e:
            print(f"FAIL: {e}")
            continue
        print(f"{len(day_vol)} bars")
        if not day_vol:
            continue
        _replace_day(conn, symbol, day, day_vol)


def _sheet(name: str):
    wb = xw.Book(BOOK)
    for s in wb.sheets:
        if s.name == name:
            return s
    return wb.sheets.add(name)


def _nonempty_days(conn: sqlite3.Connection, symbol: str, limit: int) -> list[date]:
    rows = conn.execute(
        """
        SELECT session_date FROM minute_volume
        WHERE symbol = ?
        GROUP BY session_date
        HAVING COUNT(*) > 0
        ORDER BY session_date DESC
        LIMIT ?
        """,
        (symbol, limit),
    ).fetchall()
    days = [date.fromisoformat(row[0]) for row in rows]
    days.reverse()
    return days


def _load_volume(
    conn: sqlite3.Connection, symbol: str, days: list[date]
) -> dict[tuple[date, dtime], float]:
    if not days:
        return {}
    placeholders = ",".join("?" for _ in days)
    rows = conn.execute(
        f"""
        SELECT session_date, minute, volume FROM minute_volume
        WHERE symbol = ? AND session_date IN ({placeholders})
        """,
        (symbol, *[day.isoformat() for day in days]),
    )
    out: dict[tuple[date, dtime], float] = {}
    for session_date, minute, volume in rows:
        hh, mm = minute.split(":")
        out[(date.fromisoformat(session_date), dtime(int(hh), int(mm)))] = float(volume)
    return out


def _write_excel(
    symbol: str,
    minutes: list[dtime],
    days: list[date],
    volume: dict[tuple[date, dtime], float],
) -> None:
    sheet = _sheet(symbol)
    clear_rows = 1 + len(minutes)
    sheet.range(f"A1:{CLEAR_COL}{clear_rows}").value = None

    header: list = ["time"]
    for day in days:
        header.append(datetime(day.year, day.month, day.day))

    rows: list[list] = []
    for minute in minutes:
        excel_time = datetime(1899, 12, 30, minute.hour, minute.minute)
        row: list = [excel_time]
        for day in days:
            row.append(volume.get((day, minute)))
        rows.append(row)

    sheet.range("A1").value = [header] + rows
    last_col = 1 + len(days)
    sheet.range((2, 1), (1 + len(minutes), 1)).number_format = "hh:mm"
    if days:
        sheet.range((1, 2), (1, last_col)).number_format = "yyyy-mm-dd"
        sheet.range((2, 2), (1 + len(minutes), last_col)).number_format = "#,##0"


def _export_excel(conn: sqlite3.Connection, symbols: list[str]) -> None:
    minutes = _rth_minutes()
    for symbol in symbols:
        days = _nonempty_days(conn, symbol, EXCEL_DAYS)
        if not days:
            print(f"{symbol}: в базе нет дней, лист не пишу")
            continue
        volume = _load_volume(conn, symbol, days)
        _write_excel(symbol, minutes, days, volume)
        print(
            f"Wrote {len(minutes)} minutes x {len(days)} days "
            f"-> {BOOK}/{symbol}"
        )


def main() -> None:
    args = _parse_args()
    if args.days < 1:
        raise SystemExit("--days must be >= 1")
    symbols = (
        [s.strip().upper() for s in args.symbol if s.strip()]
        if args.symbol
        else _all_symbols()
    )
    if not symbols:
        raise SystemExit("Список тикеров пуст")

    conn = _db()
    try:
        if args.to_excel:
            _export_excel(conn, symbols)
            return

        today = datetime.now(ET).date()
        days = _session_days(today, args.days)
        print(
            f"{len(symbols)} tickers, {days[0].isoformat()}..{days[-1].isoformat()} "
            f"({len(days)} будней), пауза {PACE_SEC}с между запросами к IBKR"
        )
        print(f"DB {DB_PATH}")
        ib = _connect_ib()
        try:
            for n, symbol in enumerate(symbols, start=1):
                print(f"[{n}/{len(symbols)}] {symbol}")
                _fill_symbol(ib, conn, symbol, days, today)
        finally:
            ib.disconnect()
    finally:
        conn.close()


if __name__ == "__main__":
    main()
