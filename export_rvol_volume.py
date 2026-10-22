"""
RTH 1m volume → ATR_Monitor.xlsx / лист с именем тикера.

Один тикер, последние 90 календарных дней (выходные пропускаем,
праздники без баров не попадают в колонки). 10-минутный RVOL
считается потом в Excel.

Сетка:
  колонка A — минута сессии 09:30..15:59 ET
  колонки B+ — день, слева старые, справа новые
  ячейка — volume этой минуты (пусто, если бара не было)

Блок A1:CZ391 перезаписывается. Формулы — ниже 392-й строки
или правее колонки CZ.

IBKR отдаёт 1m только по одному дню за запрос и не больше
60 запросов за 10 минут, поэтому между днями пауза 10с
(~10 минут на окно).

  python export_rvol_volume.py
  python export_rvol_volume.py --symbol NVDA --days 90
"""

from __future__ import annotations

import argparse
import asyncio
import time
from datetime import date, datetime, timedelta, time as dtime
from zoneinfo import ZoneInfo

import xlwings as xw
from ib_insync import IB, Stock

try:
    asyncio.get_event_loop()
except RuntimeError:
    asyncio.set_event_loop(asyncio.new_event_loop())

SYMBOL = "AAPL"
LOOKBACK_DAYS = 90

BOOK = "ATR_Monitor.xlsx"
# A1:CZ391 — данные. Правее и ниже не трогаем.
CLEAR_COL = "CZ"
RTH_START = dtime(9, 30)
RTH_END = dtime(15, 59)

HOST = "127.0.0.1"
PORT = 7496  # 7496=live TWS, 7497=paper/Gateway
CLIENT_ID = 93
PACE_SEC = 10
HIST_TIMEOUT = 20

ET = ZoneInfo("America/New_York")


def _parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Export RTH 1m volume to an Excel sheet named as the ticker")
    p.add_argument("--symbol", default=SYMBOL, help=f"One ticker (default {SYMBOL})")
    p.add_argument(
        "--days",
        type=int,
        default=LOOKBACK_DAYS,
        help=f"Calendar lookback through today ET (default {LOOKBACK_DAYS})",
    )
    return p.parse_args()


def _connect() -> IB:
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
            # Таймаут библиотека глотает и возвращает []. Праздник отвечает сразу.
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


def _sheet(name: str):
    wb = xw.Book(BOOK)
    for s in wb.sheets:
        if s.name == name:
            return s
    return wb.sheets.add(name)


def _write(sheet, minutes: list[dtime], days: list[date], volume: dict[tuple[date, dtime], float]) -> None:
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


def main() -> None:
    args = _parse_args()
    symbol = args.symbol.strip().upper()
    if not symbol:
        raise SystemExit("Пустой --symbol")
    if args.days < 1:
        raise SystemExit("--days must be >= 1")

    today = datetime.now(ET).date()
    minutes = _rth_minutes()
    days = _session_days(today, args.days)
    print(
        f"{symbol} RTH 1m volume, {days[0].isoformat()}..{days[-1].isoformat()} "
        f"({len(days)} будней), пауза {PACE_SEC}с"
    )

    ib = _connect()
    volume: dict[tuple[date, dtime], float] = {}
    kept: list[date] = []
    try:
        contract = Stock(symbol, "SMART", "USD")
        ib.qualifyContracts(contract)
        for i, day in enumerate(days):
            if i:
                time.sleep(PACE_SEC)
            print(f"[{i + 1}/{len(days)}] {day.isoformat()} ...", end=" ", flush=True)
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
            kept.append(day)
            for minute, vol in day_vol.items():
                volume[(day, minute)] = vol
    finally:
        ib.disconnect()

    if not kept:
        raise SystemExit(f"Нет RTH баров {symbol} за последние {args.days} дней.")

    sheet = _sheet(symbol)
    _write(sheet, minutes, kept, volume)
    print(
        f"Wrote {len(minutes)} minutes x {len(kept)} days "
        f"-> {BOOK}/{symbol}"
    )


if __name__ == "__main__":
    main()
