"""
Скрейп Rel Volume (10 MIN) из TWS → Postgres (source='ib').

TWS на Java: UIA обычно видит только рамку окна (см. tws_ui_dump.txt).
Поэтому основной путь — OCR выделенного региона Favorites.

Окно: TWS_WINDOW_TITLE из .env (заголовок окна TWS)
MONITOR → Favorites: FIN INSTR | RVOL сразу справа от тикера

Запуск:
  python rvol_ib_ui.py --calibrate   # выделить мышкой таблицу Favorites
  python rvol_ib_ui.py --once        # один снимок
  python rvol_ib_ui.py               # цикл на :00
  python rvol_ib_ui.py --dump        # диагностика UIA (обычно пусто)

Регион хранится в .env: TWS_OCR_REGION=left,top,width,height
"""
from __future__ import annotations

import argparse
import os
import re
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from dotenv import load_dotenv

load_dotenv()

from rvol_db import ensure_schema, insert_ticks  # noqa: E402
from tikers import load_tickers_from_excel  # noqa: E402

WINDOW_TITLE = os.getenv("TWS_WINDOW_TITLE", "Main Interactive Brokers")
SOURCE = "ib"
ENV_PATH = Path(__file__).resolve().parent / ".env"
DUMP_PATH = Path(__file__).resolve().parent / "tws_ui_dump.txt"
DEBUG_SHOT = Path(__file__).resolve().parent / "tws_ocr_debug.png"
# Тикеры читаем из Excel ОДИН раз при старте — COM/xlwings ворует фокус.
_TICKER_CACHE: List[str] = []


def load_watchlist(*, force_reload: bool = False) -> List[str]:
    """Без Excel в минутном цикле — только кэш в памяти."""
    global _TICKER_CACHE
    if _TICKER_CACHE and not force_reload:
        return list(_TICKER_CACHE)
    tickers = load_tickers_from_excel(activate=False)
    _TICKER_CACHE = list(tickers)
    print(f"Watchlist загружен из Excel: {len(_TICKER_CACHE)} тикеров (дальше без COM)")
    return list(_TICKER_CACHE)


def wait_until_minute_start() -> None:
    now = datetime.now()
    if now.second == 0 and now.microsecond < 50_000:
        return
    target = now.replace(second=0, microsecond=0) + timedelta(minutes=1)
    delay = (target - datetime.now()).total_seconds()
    if delay > 0:
        print(f"Ждём :00 → {target.strftime('%H:%M:%S')} ({delay:.1f}с)...")
        time.sleep(delay)


def _parse_rvol(text: str) -> Optional[float]:
    if text is None:
        return None
    t = str(text).strip().replace("%", "").replace(",", "").replace(" ", "")
    if not t or t.upper() in ("N/A", "NA", "—", "-", "--"):
        return None
    t = re.sub(r"[xX]$", "", t)
    try:
        return float(t)
    except ValueError:
        return None


def _clean_ticker(text: str) -> str:
    if not text:
        return ""
    raw = re.split(r"[\n\r]", str(text).strip())[0].strip()
    parts = re.split(r"\s+", raw)
    if not parts:
        return ""
    sym = re.sub(r"[^A-Za-z0-9.\-]", "", parts[0]).upper()
    return sym


def parse_region(s: Optional[str]) -> Optional[Tuple[int, int, int, int]]:
    if not s:
        return None
    parts = [p.strip() for p in s.split(",")]
    if len(parts) != 4:
        return None
    left, top, width, height = map(int, parts)
    return left, top, width, height


def save_region_to_env(region: Tuple[int, int, int, int]) -> None:
    left, top, width, height = region
    line = f"TWS_OCR_REGION={left},{top},{width},{height}"
    text = ENV_PATH.read_text(encoding="utf-8") if ENV_PATH.exists() else ""
    if re.search(r"^TWS_OCR_REGION=.*$", text, flags=re.M):
        text = re.sub(r"^TWS_OCR_REGION=.*$", line, text, flags=re.M)
    else:
        if text and not text.endswith("\n"):
            text += "\n"
        text += line + "\n"
    ENV_PATH.write_text(text, encoding="utf-8")
    os.environ["TWS_OCR_REGION"] = f"{left},{top},{width},{height}"
    print(f"Сохранено в .env: {line}")


def calibrate_region() -> Tuple[int, int, int, int]:
    """Полноэкранный оверлей: зажми ЛКМ и выдели таблицу Favorites."""
    import tkinter as tk

    from PIL import ImageGrab

    print("Калибровка: зажмите ЛКМ и выделите область таблицы Favorites (FIN INSTR + RVOL).")
    print("Отпустите кнопку — регион сохранится. Esc — отмена.")

    root = tk.Tk()
    root.attributes("-fullscreen", True)
    root.attributes("-alpha", 0.25)
    root.attributes("-topmost", True)
    root.configure(bg="black")
    canvas = tk.Canvas(root, cursor="cross", bg="black", highlightthickness=0)
    canvas.pack(fill=tk.BOTH, expand=True)

    state = {"x0": 0, "y0": 0, "rect": None, "done": None}

    def on_press(e):
        state["x0"], state["y0"] = e.x_root, e.y_root
        if state["rect"]:
            canvas.delete(state["rect"])
        state["rect"] = canvas.create_rectangle(
            e.x, e.y, e.x, e.y, outline="lime", width=2
        )

    def on_drag(e):
        if state["rect"]:
            x0 = state["x0"] - root.winfo_rootx()
            y0 = state["y0"] - root.winfo_rooty()
            canvas.coords(state["rect"], x0, y0, e.x, e.y)

    def on_release(e):
        x1, y1 = state["x0"], state["y0"]
        x2, y2 = e.x_root, e.y_root
        left, top = min(x1, x2), min(y1, y2)
        width, height = abs(x2 - x1), abs(y2 - y1)
        if width < 20 or height < 20:
            print("Слишком маленький регион, попробуйте ещё раз.")
            return
        state["done"] = (left, top, width, height)
        root.destroy()

    def on_escape(_e):
        state["done"] = None
        root.destroy()

    canvas.bind("<ButtonPress-1>", on_press)
    canvas.bind("<B1-Motion>", on_drag)
    canvas.bind("<ButtonRelease-1>", on_release)
    root.bind("<Escape>", on_escape)
    root.mainloop()

    if not state["done"]:
        raise SystemExit("Калибровка отменена")

    region = state["done"]
    # контрольный скрин
    left, top, width, height = region
    img = ImageGrab.grab(bbox=(left, top, left + width, top + height))
    img.save(DEBUG_SHOT)
    print(f"Превью региона: {DEBUG_SHOT}")
    save_region_to_env(region)
    return region


def capture_region(region: Tuple[int, int, int, int]):
    import mss
    import numpy as np
    from PIL import Image

    left, top, width, height = region
    with mss.mss() as sct:
        shot = sct.grab({"left": left, "top": top, "width": width, "height": height})
        img = Image.frombytes("RGB", shot.size, shot.bgra, "raw", "BGRX")
    return img


def ocr_boxes(img) -> List[Tuple[str, float, float, float, float]]:
    """
    Возвращает список (text, cx, cy, x1, x2) в координатах картинки.
    """
    import numpy as np
    from rapidocr_onnxruntime import RapidOCR

    engine = getattr(ocr_boxes, "_engine", None)
    if engine is None:
        engine = RapidOCR()
        ocr_boxes._engine = engine

    result, _ = engine(np.array(img))
    out: List[Tuple[str, float, float, float, float]] = []
    if not result:
        return out
    for item in result:
        # item: [box, text, score] where box is 4 points
        box, text, score = item[0], item[1], item[2]
        if score is not None and float(score) < 0.4:
            continue
        xs = [p[0] for p in box]
        ys = [p[1] for p in box]
        x1, x2 = min(xs), max(xs)
        y1, y2 = min(ys), max(ys)
        cx, cy = (x1 + x2) / 2, (y1 + y2) / 2
        out.append((str(text).strip(), cx, cy, x1, x2))
    return out


def _is_rvol_header(text: str) -> bool:
    """RVOL / RVol... / RVO… (часто обрезано)."""
    n = re.sub(r"[\s.·…]+", "", text.strip().upper())
    if n.startswith("VOLUME"):
        return False
    # RVOL, RVO, RVOL10MIN, RVOL...
    if n.startswith("RVOL") or n.startswith("RVO"):
        return True
    return False


def _find_rvol_header(
    boxes: List[Tuple[str, float, float, float, float]],
) -> Optional[Tuple[float, float, str]]:
    """
    Возвращает (anchor_x, y, text).
    Для обрезанного 'RVol...' якорь = левая треть бокса (ближе к числам колонки),
    а не геометрический центр широкой подписи.
    """
    hits = []
    for text, cx, cy, x1, x2 in boxes:
        if _is_rvol_header(text):
            # якорь ближе к левому краю заголовка
            anchor = x1 + max(8.0, (x2 - x1) * 0.25)
            hits.append((cy, x1, anchor, text))
    if not hits:
        return None
    hits.sort(key=lambda h: (h[0], h[1]))
    top_y = hits[0][0]
    top_hits = [h for h in hits if abs(h[0] - top_y) < 30]
    top_hits.sort(key=lambda h: h[1])
    cy, _x1, anchor, text = top_hits[0]
    print(f"Заголовок RVOL распознан как {text!r}, anchor_x≈{anchor:.0f}")
    return anchor, cy, text


def parse_rvol_from_ocr(
    boxes: List[Tuple[str, float, float, float, float]],
    watchlist: List[str],
) -> Dict[str, float]:
    """
    Layout: FIN INSTR | RVOL | LAST | …
    Главное правило: ближайшее число СРАЗУ СПРАВА от тикера на строке.
    Заголовок RVol... — мягкий бонус по X, не жёсткий отсев.
    """
    wanted = {t.upper() for t in watchlist}
    header = _find_rvol_header(boxes)
    rvol_x = header[0] if header else None
    if rvol_x is None:
        print("Заголовок RVOL не найден — только «сосед справа»")

    skip = {
        "RVOL",
        "RVO",
        "FIN",
        "MIN",
        "INSTR",
        "VOLUME",
        "LAST",
        "CHANGE",
        "CHNG",
        "MONITOR",
        "FAVORITES",
    }
    tickers: List[Tuple[str, float, float, float]] = []  # sym, cx, cy, x2
    numbers: List[Tuple[float, float, float]] = []  # val, cx, cy

    for text, cx, cy, x1, x2 in boxes:
        if _is_rvol_header(text):
            continue
        sym = _clean_ticker(text)
        if sym in skip:
            continue
        if sym in wanted or (
            sym
            and re.fullmatch(r"[A-Z][A-Z0-9.\-]{0,11}", sym)
            and 1 <= len(sym) <= 6
        ):
            tickers.append((sym, cx, cy, x2))

        v = _parse_rvol(text)
        # RVOL 0.74 / 1.92; не LAST 746
        if v is None or not (0 < v < 80):
            continue
        numbers.append((v, cx, cy))

    if not tickers:
        print("⚠️ Тикеры в OCR не найдены")
        return {}
    if not numbers:
        print("⚠️ Числа RVOL-диапазона в OCR не найдены")
        return {}

    # оценка медианы X «первого числа справа» по строкам — уточняет колонку
    first_right_xs: List[float] = []
    for _sym, tx, ty, tx2 in tickers:
        row_nums = [
            (nx - tx2, nx)
            for v, nx, ny in numbers
            if abs(ny - ty) <= 18 and nx >= tx2 - 5
        ]
        if row_nums:
            row_nums.sort()
            first_right_xs.append(row_nums[0][1])
    col_x = None
    if first_right_xs:
        first_right_xs.sort()
        col_x = first_right_xs[len(first_right_xs) // 2]
        print(f"Колонка чисел справа от тикеров @ x≈{col_x:.0f}")
    # если заголовок далеко от реальной колонки чисел — доверяем числам
    if rvol_x is not None and col_x is not None and abs(rvol_x - col_x) > 50:
        print(
            f"Якорь заголовка ({rvol_x:.0f}) далеко от колонки чисел ({col_x:.0f}) — "
            "берём колонку чисел"
        )
        rvol_x = col_x
    elif rvol_x is None:
        rvol_x = col_x

    result: Dict[str, float] = {}
    for sym, tx, ty, tx2 in tickers:
        if wanted and sym not in wanted:
            continue
        best = None
        best_score = 1e18
        for v, nx, ny in numbers:
            dy = abs(ny - ty)
            if dy > 20:
                continue
            dx = nx - tx2
            if dx < -8:
                continue
            # сосед справа важнее всего; близость к колонке RVOL — бонус
            score = dx * 8 + dy * 3
            if rvol_x is not None:
                score += abs(nx - rvol_x) * 0.5
            # типичный RVOL чаще < 20 — лёгкий приоритет
            if v >= 20:
                score += 30
            if score < best_score:
                best_score = score
                best = v
        if best is not None:
            result[sym] = best

    print(f"Сматчено тикеров: {len(result)}/{len(wanted) if wanted else len(tickers)}")
    return result





def dump_uia() -> None:
    from pywinauto import Application

    print(f"UIA dump → {DUMP_PATH}")
    app = Application(backend="uia").connect(title=WINDOW_TITLE, timeout=10)
    win = app.top_window()
    lines: List[str] = []

    def walk(ctrl, depth=0):
        if depth > 6:
            return
        try:
            info = ctrl.element_info
            lines.append(
                f"{'  ' * depth}[{info.control_type}] name={(info.name or '')[:80]!r}"
            )
            for ch in ctrl.children():
                walk(ch, depth + 1)
        except Exception:
            return

    walk(win)
    DUMP_PATH.write_text("\n".join(lines), encoding="utf-8")
    print(
        f"Строк: {len(lines)}. Если только TitleBar/кнопки — Java UI закрыт, "
        "используйте OCR (--calibrate)."
    )


def snapshot_once(region: Tuple[int, int, int, int], save_debug: bool = True) -> int:
    watchlist = load_watchlist()
    if not watchlist:
        print("Watchlist пуст — загрузите тикеры при старте (Excel уже открыт)")
        return 0

    # только mss BitBlt региона — без win32 focus / без Excel COM
    img = capture_region(region)
    if save_debug:
        img.save(DEBUG_SHOT)

    boxes = ocr_boxes(img)
    print(f"OCR: {len(boxes)} текстовых блоков")
    rvol_map = parse_rvol_from_ocr(boxes, watchlist)

    ts = datetime.now().replace(second=0, microsecond=0)
    rows = []
    missing = []
    for t in watchlist:
        key = t.upper()
        if key in rvol_map:
            rows.append((ts, key, rvol_map[key]))
        else:
            missing.append(t)

    n = insert_ticks(rows, source=SOURCE)
    print(
        f"{datetime.now().strftime('%H:%M:%S')} → {ts.strftime('%H:%M')} "
        f"IB RVOL записано {n}/{len(watchlist)}"
        + (f", нет: {', '.join(missing)}" if missing else "")
    )
    if n == 0:
        print(f"Проверьте превью {DEBUG_SHOT} и что в регионе видны FIN INSTR + RVOL.")
    return n


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--once", action="store_true")
    parser.add_argument("--dump", action="store_true", help="UIA dump (обычно бесполезен для TWS)")
    parser.add_argument("--calibrate", action="store_true", help="Выделить регион Favorites")
    parser.add_argument(
        "--reload-tickers",
        action="store_true",
        help="Перечитать Лист2!L из Excel один раз (может моргнуть фокусом)",
    )
    args = parser.parse_args()

    ensure_schema()

    if args.dump:
        dump_uia()
        return

    if args.calibrate:
        calibrate_region()
        load_dotenv(override=True)
        region = parse_region(os.getenv("TWS_OCR_REGION"))
        assert region
        load_watchlist(force_reload=True)
        print("Пробный OCR...")
        snapshot_once(region)
        return

    region = parse_region(os.getenv("TWS_OCR_REGION"))
    if not region:
        print("Нет TWS_OCR_REGION в .env. Сначала: python rvol_ib_ui.py --calibrate")
        raise SystemExit(1)

    # Excel COM только здесь — потом минутный цикл без него
    load_watchlist(force_reload=True)

    if args.once:
        if args.reload_tickers:
            load_watchlist(force_reload=True)
        snapshot_once(region)
        return

    print(
        f"TWS OCR RVOL → Postgres. Регион={region}. Снимок на :00. "
        "Без Excel/focus в цикле. Ctrl+C стоп."
    )
    while True:
        try:
            wait_until_minute_start()
            snapshot_once(region, save_debug=False)
        except KeyboardInterrupt:
            print("\nСтоп.")
            break
        except Exception as e:
            print(f"Ошибка: {e}")
            time.sleep(2)


if __name__ == "__main__":
    main()
