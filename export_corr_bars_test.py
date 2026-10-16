"""
1m closes → ATR_Monitor.xlsx / Лист1 / колонки M+
QQQ + тикеры из tikers.py

Live (как rvol_logger): на старте затираем блок, дальше на :03 каждой
минуты (ET) дописываем строку. Сессия с 09:30 ET.

  python export_corr_bars_test.py
  python export_corr_bars_test.py --start 09:30 --end 16:00

Разовая выгрузка за день в прошлом:

  python export_corr_bars_test.py --date 2026-10-13
  python export_corr_bars_test.py --date 2026-10-13 --start 09:30 --end 16:00
"""

from __future__ import annotations

import argparse
import asyncio
import time
from datetime import date, datetime, timedelta, time as dtime
from zoneinfo import ZoneInfo

import xlwings as xw
from ib_insync import IB, Stock

from tikers import TICKERS, load_tickers_from_excel

# Python 3.12+: нет default event loop в MainThread
try:
    asyncio.get_event_loop()
except RuntimeError:
    asyncio.set_event_loop(asyncio.new_event_loop())

BOOK = "ATR_Monitor.xlsx"
SHEET = "Лист1"
START_COL = "N"  # time, затем QQQ, затем тикеры
END_COL = "AH"  # wipe только M:AF, правее не трогаем

HOST = "127.0.0.1"
PORT = 7496  # 7496=live TWS, 7497=paper/Gateway
CLIENT_ID = 91

ET = ZoneInfo("America/New_York")
DEFAULT_START = dtime(9, 30)
DEFAULT_END_LIVE = dtime(16, 0)
DEFAULT_END_HIST = dtime(9, 40)
GRACE_SEC = 3  # :03 каждой минуты, как rvol_logger


def _parse_hhmm(text: str) -> dtime:
    return datetime.strptime(text.strip(), "%H:%M").time()


def _parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Export 1m IBKR bars to Excel for corr")
    p.add_argument(
        "--date",
        metavar="YYYY-MM-DD",
        help="One-shot: session date in ET (no live loop). Example: 2026-10-13",
    )
    p.add_argument(
        "--start",
        default=DEFAULT_START.strftime("%H:%M"),
        help=f"Window start ET HH:MM (default {DEFAULT_START.strftime('%H:%M')})",
    )
    p.add_argument(
        "--end",
        default=None,
        help="Window end ET HH:MM inclusive "
        f"(live default {DEFAULT_END_LIVE.strftime('%H:%M')}, "
        f"--date default {DEFAULT_END_HIST.strftime('%H:%M')})",
    )
    return p.parse_args()


def _ordered_symbols() -> list[str]:
    tickers = list(TICKERS) if TICKERS else load_tickers_from_excel()
    if not tickers:
        raise SystemExit("TICKERS пуст — открой ATR_Monitor.xlsx и колонку L на Лист2")
    ordered: list[str] = []
    seen: set[str] = set()
    for s in ["QQQ", *tickers]:
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
        time.sleep(0.6)
    return series


def _row_for_minute(
    ts: datetime, ordered: list[str], series: dict[str, dict[datetime, float]]
) -> list:
    row = [ts.replace(tzinfo=None)]
    for sym in ordered:
        row.append(series.get(sym, {}).get(ts))
    return row


def _sheet():
    wb = xw.Book(BOOK)
    return wb.sheets[SHEET]


def _old_last_row(sheet, start) -> int:
    if start.value is None:
        return 1
    try:
        return start.end("down").row
    except Exception:
        return 1


def _wipe_block(sheet, ordered: list[str]) -> None:
    """Затираем только M:AF при старте (без clear_contents на огромном range)."""
    start = sheet.range(f"{START_COL}1")
    old_last = _old_last_row(sheet, start)
    wipe_rows = max(old_last, 2)
    rng = sheet.range(f"{START_COL}1:{END_COL}{wipe_rows}")
    n_cols = rng.columns.count
    rng.value = [[None] * n_cols for _ in range(wipe_rows)]
    header = ["time", *ordered]
    if len(header) > n_cols:
        raise SystemExit(
            f"Слишком много колонок ({len(header)}): блок {START_COL}:{END_COL} "
            f"вмещает {n_cols} (time + QQQ + тикеры)"
        )
    start.value = [header]
    sheet.range(f"{START_COL}1").number_format = "@"


def _append_row(sheet, excel_row: int, row: list) -> None:
    sheet.range(f"{START_COL}{excel_row}").value = [row]
    sheet.range(f"{START_COL}{excel_row}").number_format = "hh:mm"


def _write_block(sheet, ordered: list[str], rows: list[list]) -> None:
    """Полная перезапись после wipe (backfill / --date)."""
    start = sheet.range(f"{START_COL}1")
    start.value = [["time", *ordered]] + rows
    n_rows = 1 + len(rows)
    if n_rows >= 2:
        sheet.range(f"{START_COL}2:{START_COL}{n_rows}").number_format = "hh:mm"


def wait_until_minute_grace(grace: int = GRACE_SEC) -> datetime:
    """Ждём :grace секунд текущей/следующей минуты по ET. Возвращает target."""
    now = datetime.now(ET)
    minute_start = now.replace(second=0, microsecond=0)
    if now.second < grace:
        target = minute_start + timedelta(seconds=grace)
    else:
        target = minute_start + timedelta(minutes=1, seconds=grace)
    delay = (target - datetime.now(ET)).total_seconds()
    if delay > 0:
        print(f"Ждём до {target.strftime('%H:%M:%S')} ET ({delay:.1f}с)...")
        time.sleep(delay)
    return target


def run_oneshot(session_day: date, t0: dtime, t1: dtime) -> None:
    today_et = datetime.now(ET).date()
    is_today = session_day == today_et
    end_dt = _ib_end_datetime(session_day, is_today=is_today)
    ordered = _ordered_symbols()

    print(
        f"One-shot {session_day.isoformat()} ET "
        f"{t0.strftime('%H:%M')}-{t1.strftime('%H:%M')} "
        f"(endDateTime={repr(end_dt) if end_dt else 'now'})"
    )

    ib = _connect()
    try:
        series = _fetch_all(
            ib, ordered, session_day=session_day, end_dt=end_dt, t0=t0, t1=t1
        )
    finally:
        ib.disconnect()

    all_times = sorted({t for d in series.values() for t in d})
    if not all_times:
        raise SystemExit(
            f"Нет баров за {session_day} {t0.strftime('%H:%M')}-{t1.strftime('%H:%M')} ET."
        )

    rows = [_row_for_minute(ts, ordered, series) for ts in all_times]
    sheet = _sheet()
    _wipe_block(sheet, ordered)
    _write_block(sheet, ordered, rows)
    print(f"Wrote {len(rows)} rows -> {BOOK}/{SHEET} from {START_COL}")


def run_live(t0: dtime, t1: dtime) -> None:
    ordered = _ordered_symbols()
    session_day = datetime.now(ET).date()
    print(
        f"Live corr bars → {BOOK}/{SHEET} {START_COL}+. "
        f"С {t0.strftime('%H:%M')} до {t1.strftime('%H:%M')} ET, "
        f"снимок на :{GRACE_SEC:02d}. Ctrl+C стоп."
    )

    sheet = _sheet()
    _wipe_block(sheet, ordered)
    print("Блок очищен, заголовок записан.")

    ib = _connect()
    last_written: datetime | None = None
    next_excel_row = 2

    try:
        # backfill с открытия до последней полной минуты (если уже после t0)
        now = datetime.now(ET)
        if now.date() == session_day and now.time() >= t0:
            # на :03 пишем текущую минуту; для backfill — до предыдущей закрытой
            # если сейчас >= :03 текущей минуты, текущая уже «доступна»
            if now.second >= GRACE_SEC:
                fill_end = now.replace(second=0, microsecond=0)
            else:
                fill_end = now.replace(second=0, microsecond=0) - timedelta(minutes=1)
            if fill_end.time() >= t0:
                t_fill = min(fill_end.time(), t1)
                print(
                    f"Backfill {session_day} {t0.strftime('%H:%M')}-"
                    f"{t_fill.strftime('%H:%M')} ET..."
                )
                series = _fetch_all(
                    ib,
                    ordered,
                    session_day=session_day,
                    end_dt="",
                    t0=t0,
                    t1=t_fill,
                )
                times = sorted({t for d in series.values() for t in d})
                rows = [_row_for_minute(ts, ordered, series) for ts in times]
                if rows:
                    _write_block(sheet, ordered, rows)
                    last_written = times[-1]
                    next_excel_row = 2 + len(rows)
                    print(f"Backfill: {len(rows)} строк, last={last_written.strftime('%H:%M')}")

        while True:
            target = wait_until_minute_grace(GRACE_SEC)
            bar_ts = target.replace(second=0, microsecond=0)

            if bar_ts.date() != session_day:
                # новый день — перезапуск логики дня
                session_day = bar_ts.date()
                _wipe_block(sheet, ordered)
                last_written = None
                next_excel_row = 2
                print(f"Новый день {session_day}, блок очищен.")

            if bar_ts.time() < t0:
                print(f"{bar_ts.strftime('%H:%M')} ET — до открытия, skip")
                continue
            if bar_ts.time() > t1:
                print(f"{bar_ts.strftime('%H:%M')} ET — после окна, стоп")
                break
            if last_written is not None and bar_ts <= last_written:
                print(f"{bar_ts.strftime('%H:%M')} уже записан, skip")
                continue

            try:
                series = _fetch_all(
                    ib,
                    ordered,
                    session_day=session_day,
                    end_dt="",
                    t0=bar_ts.time(),
                    t1=bar_ts.time(),
                    duration="600 S",
                )
                # если короткий duration не отдал бар — fallback 1 D
                if not any(bar_ts in d for d in series.values()):
                    series = _fetch_all(
                        ib,
                        ordered,
                        session_day=session_day,
                        end_dt="",
                        t0=bar_ts.time(),
                        t1=bar_ts.time(),
                        duration="1 D",
                    )
                row = _row_for_minute(bar_ts, ordered, series)
                _append_row(sheet, next_excel_row, row)
                print(
                    f"{datetime.now(ET).strftime('%H:%M:%S')} ET → "
                    f"row {next_excel_row} {bar_ts.strftime('%H:%M')} "
                    f"QQQ={row[1]}"
                )
                last_written = bar_ts
                next_excel_row += 1
            except Exception as e:
                print(f"Ошибка снимка {bar_ts.strftime('%H:%M')}: {e}")
                time.sleep(1)
    finally:
        ib.disconnect()


def main():
    args = _parse_args()
    t0 = _parse_hhmm(args.start)
    if args.end:
        t1 = _parse_hhmm(args.end)
    elif args.date:
        t1 = DEFAULT_END_HIST
    else:
        t1 = DEFAULT_END_LIVE
    if t1 < t0:
        raise SystemExit("--end must be >= --start")

    if args.date:
        run_oneshot(date.fromisoformat(args.date), t0, t1)
    else:
        run_live(t0, t1)


if __name__ == "__main__":
    main()
