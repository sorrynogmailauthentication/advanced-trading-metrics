"""
ORB snapshot: ждёт 09:36:00 и копирует live High/Low → ORB High/Low
в уже открытом Excel (TWS RTD должен работать).

Архитектура книги:
  Лист1 — RVOL/скрипт (не трогаем)
  Лист2 — основной монитор с данными Interactive Brokers (RTD)

Запуск вручную до 9:36. Книга ATR_Monitor.xlsx должна быть открыта.
"""
from __future__ import annotations

import time
from datetime import datetime, time as dtime
from typing import Any, List, Optional

import xlwings as xw

# --- Лист2: основной лист с IB RTD (как в tikers.py) ---
WORKBOOK_NAME = "ATR_Monitor.xlsx"
SHEET_NAME = "Лист2"

FIRST_ROW = 2
LAST_ROW: Optional[int] = None  # None = до последней непустой строки тикера
TICKER_COL = "L"  # тикеры на Лист2

# Подставьте буквы колонок с RTD High/Low на Лист2
LIVE_HIGH_COL = "E"  # RTD High
LIVE_LOW_COL = "D"   # RTD Low
ORB_HIGH_COL = "F"  # снимок High @ 9:36 (статические значения)
ORB_LOW_COL = "G"   # снимок Low @ 9:36

# Локальное время ПК. Для NYSE при ПК в UTC+3 обычно 16:36:00.
SNAP_TIME = dtime(16, 36, 0)

# Повтор, если в момент снимка часть ячеек ещё пустая / ошибка RTD
RETRY_SECONDS = 3
RETRY_PAUSE = 0.5


def wait_until(target: dtime) -> None:
    now = datetime.now()
    target_dt = datetime.combine(now.date(), target)
    if now >= target_dt:
        print(f"Уже позже {target.strftime('%H:%M:%S')} — снимаем сразу.")
        return

    secs = (target_dt - now).total_seconds()
    print(f"Ждём до {target.strftime('%H:%M:%S')} ({secs:.0f} сек)...")
    if secs > 2:
        time.sleep(secs - 1)
    while datetime.now().time() < target:
        time.sleep(0.05)
    print(f"Снимок: {datetime.now().strftime('%H:%M:%S.%f')[:-3]}")


def last_data_row(sheet, ticker_col: str, first_row: int, last_row: Optional[int]) -> int:
    if last_row is not None:
        return last_row
    cell = sheet.range(f"{ticker_col}{first_row}")
    if cell.value is None:
        return first_row
    return cell.expand("down").last_cell.row


def as_column(values: Any, n_rows: int) -> List[Any]:
    """xlwings: одна ячейка → scalar; столбец → list."""
    if values is None:
        return [None] * n_rows
    if not isinstance(values, list):
        return [values]
    flat: List[Any] = []
    for v in values:
        if isinstance(v, list):
            flat.append(v[0] if v else None)
        else:
            flat.append(v)
    if len(flat) < n_rows:
        flat.extend([None] * (n_rows - len(flat)))
    return flat[:n_rows]


def is_number(v: Any) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def read_high_low(sheet, first_row: int, end_row: int):
    n = end_row - first_row + 1
    highs = as_column(
        sheet.range(f"{LIVE_HIGH_COL}{first_row}:{LIVE_HIGH_COL}{end_row}").value, n
    )
    lows = as_column(
        sheet.range(f"{LIVE_LOW_COL}{first_row}:{LIVE_LOW_COL}{end_row}").value, n
    )
    return highs, lows


def merge_prefer_numeric(prev: List[Any], new: List[Any]) -> List[Any]:
    out = []
    for p, n in zip(prev, new):
        if is_number(n):
            out.append(n)
        elif is_number(p):
            out.append(p)
        else:
            out.append(n if n is not None else p)
    return out


def write_column(sheet, col: str, first_row: int, values: List[Any]) -> None:
    sheet.range(f"{col}{first_row}").value = [[v] for v in values]


def snapshot_orb() -> None:
    wait_until(SNAP_TIME)

    try:
        wb = xw.Book(WORKBOOK_NAME)
    except Exception as e:
        print(
            f"Ошибка: откройте '{WORKBOOK_NAME}' в Excel перед запуском.\n{e}"
        )
        return

    sheet = wb.sheets[SHEET_NAME]
    end_row = last_data_row(sheet, TICKER_COL, FIRST_ROW, LAST_ROW)
    n = end_row - FIRST_ROW + 1

    highs, lows = read_high_low(sheet, FIRST_ROW, end_row)

    deadline = time.time() + RETRY_SECONDS
    while time.time() < deadline:
        missing = sum(
            1 for h, l in zip(highs, lows) if not (is_number(h) and is_number(l))
        )
        if missing == 0:
            break
        time.sleep(RETRY_PAUSE)
        h2, l2 = read_high_low(sheet, FIRST_ROW, end_row)
        highs = merge_prefer_numeric(highs, h2)
        lows = merge_prefer_numeric(lows, l2)

    write_column(sheet, ORB_HIGH_COL, FIRST_ROW, highs)
    write_column(sheet, ORB_LOW_COL, FIRST_ROW, lows)

    ok = sum(1 for h, l in zip(highs, lows) if is_number(h) and is_number(l))
    print(
        f"Готово: строки {FIRST_ROW}–{end_row} → "
        f"{ORB_HIGH_COL}/{ORB_LOW_COL} (чисел: {ok}/{n})"
    )


if __name__ == "__main__":
    snapshot_orb()
