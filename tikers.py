import re

import xlwings as xw

_EXCEL_BOOK = "ATR_Monitor.xlsx"
_SHEET_NAME = "Лист2"
_TICKER_COLUMN = "L"


def _cell_to_ticker(value):
    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""
    return re.split(r"\n|\s{2,}", text)[0].strip()


def _find_open_book(book_name: str):
    """Уже открытая книга без xw.Book() — тот активирует окно Excel."""
    for app in xw.apps:
        for book in app.books:
            try:
                if book.name == book_name or str(book.fullname).endswith(book_name):
                    return book
            except Exception:
                continue
    return None


def load_tickers_from_excel(*, activate: bool = True):
    """
    activate=False — не вызывать xw.Book() (не поднимает Excel/не дёргает z-order).
    Книга должна быть уже открыта.
    """
    if activate:
        wb = xw.Book(_EXCEL_BOOK)
    else:
        wb = _find_open_book(_EXCEL_BOOK)
        if wb is None:
            raise RuntimeError(
                f"Книга '{_EXCEL_BOOK}' не найдена среди открытых. "
                "Откройте её заранее (режим без activate)."
            )

    sheet = wb.sheets[_SHEET_NAME]
    last_row = sheet.range(
        f"{_TICKER_COLUMN}{sheet.cells.last_cell.row}"
    ).end("up").row
    if last_row < 1:
        return []

    raw = sheet.range(f"{_TICKER_COLUMN}2:{_TICKER_COLUMN}{last_row}").value
    if raw is None:
        return []
    if not isinstance(raw, list):
        cells = [raw]
    elif raw and isinstance(raw[0], list):
        cells = [row[0] for row in raw]
    else:
        cells = raw

    tickers = []
    seen = set()
    for cell in cells:
        ticker = _cell_to_ticker(cell)
        if ticker and ticker not in seen:
            seen.add(ticker)
            tickers.append(ticker)
    return tickers


try:
    TICKERS = load_tickers_from_excel()
except Exception as e:
    print(
        f"Ошибка: откройте '{_EXCEL_BOOK}' в Excel и заполните колонку "
        f"{_TICKER_COLUMN}. ({e})"
    )
    TICKERS = []
