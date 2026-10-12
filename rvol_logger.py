"""
Каждую минуту читает RVOL с Лист1 (скрипт RVOL_TV_excel) и пишет
новые строки в Postgres для тикеров из Лист2!L (tikers.load_tickers_from_excel).

Не перезаписывает историю — только INSERT.

Нужно: открытый ATR_Monitor.xlsx, локальный Postgres, DATABASE_URL в .env.
Запуск:  python rvol_logger.py
"""
from __future__ import annotations

import os
import re
import time
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional, Tuple

import xlwings as xw
from dotenv import load_dotenv

load_dotenv()

from rvol_db import ensure_schema, insert_ticks  # noqa: E402
from tikers import (  # noqa: E402
    _EXCEL_BOOK,
    load_tickers_from_excel,
)

RVOL_SHEET = os.getenv("RVOL_SHEET", "Лист1")
# Секунда внутри минуты, когда снимаем (не drift от sleep(60))
# Скрейпер пишет на :00; логгер даёт GRACE секунд на парсинг
GRACE_SEC = int(os.getenv("RVOL_LOG_GRACE", "3"))

def _norm_header(h: Any) -> str:
    if h is None:
        return ""
    return re.sub(r"\s+", " ", str(h).strip().lower())


def _parse_rvol(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = str(value).strip().replace("%", "").replace(" ", "").replace(",", ".")
    if not text or text.upper() in ("N/A", "#N/A", "—", "-"):
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _read_sheet1_rvol_map(wb) -> Dict[str, float]:
    """symbol -> rvol с Лист1 (блок скринера с G1)."""
    sheet = wb.sheets[RVOL_SHEET]
    headers = sheet.range("G1").expand("right").value
    if headers is None:
        return {}
    if not isinstance(headers, list):
        headers = [headers]

    headers_norm = [_norm_header(h) for h in headers]
    ticker_idx = 0
    for i, h in enumerate(headers_norm):
        if "symbol" in h or "ticker" in h or h == "тикер":
            ticker_idx = i
            break

    try:
        rvol_idx = headers_norm.index("rel vol at time")
    except ValueError as e:
        raise RuntimeError(
            f"Не найден столбец 'Rel vol at time' в заголовках Лист1 G1: {headers}"
        ) from e


    last_row = sheet.range("G" + str(sheet.cells.last_cell.row)).end("up").row
    if last_row < 2:
        return {}

    n_cols = len(headers)
    data = sheet.range((2, 7), (last_row, 6 + n_cols)).value
    if data is None:
        return {}
    if not isinstance(data[0], list):
        data = [data]

    out: Dict[str, float] = {}
    for row in data:
        if not row or len(row) <= max(ticker_idx, rvol_idx):
            continue
        raw_ticker = row[ticker_idx]
        if raw_ticker is None:
            continue
        ticker = str(raw_ticker).strip().split("\n")[0].strip().upper()
        rvol = _parse_rvol(row[rvol_idx])
        if ticker and rvol is not None:
            out[ticker] = rvol
    return out


def wait_until_minute_grace(grace: int = GRACE_SEC) -> None:
    """Ждём :grace секунд текущей/следующей минуты (например :05)."""
    now = datetime.now()
    minute_start = now.replace(second=0, microsecond=0)
    if now.second < grace:
        target = minute_start + timedelta(seconds=grace)
    else:
        target = minute_start + timedelta(minutes=1, seconds=grace)
    delay = (target - datetime.now()).total_seconds()
    if delay > 0:
        print(f"Ждём до {target.strftime('%H:%M:%S')} ({delay:.1f}с)...")
        time.sleep(delay)


def snapshot_once() -> int:
    tickers = load_tickers_from_excel()
    if not tickers:
        print("Список тикеров пуст (Лист2!L).")
        return 0

    wb = xw.Book(_EXCEL_BOOK)
    rvol_map = _read_sheet1_rvol_map(wb)
    # метка минуты (снимок в :03 → ts = :00 этой минуты)
    ts = datetime.now().replace(second=0, microsecond=0)

    rows: List[Tuple[datetime, str, float]] = []
    missing = []
    for t in tickers:
        key = t.upper()
        if key in rvol_map:
            rows.append((ts, key, rvol_map[key]))
        else:
            missing.append(t)

    n = insert_ticks(rows)
    print(
        f"{datetime.now().strftime('%H:%M:%S')} → {ts.strftime('%H:%M')} "
        f"записано {n}/{len(tickers)}"
        + (f", нет на Лист1: {', '.join(missing)}" if missing else "")
    )
    return n


def main() -> None:
    print(
        f"RVOL logger → Postgres. Снимок на :{GRACE_SEC:02d} каждой минуты. Ctrl+C стоп."
    )
    ensure_schema()
    while True:
        try:
            wait_until_minute_grace()
            snapshot_once()
        except Exception as e:
            print(f"Ошибка снимка: {e}")
            time.sleep(1)


if __name__ == "__main__":
    main()
