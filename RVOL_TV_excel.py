"""
Скрейпер TradingView → Excel (Лист1): только Symbol + Rel vol at time,
только тикеры из tikers.load_tickers_from_excel() (Лист2!L).

Цикл с :00; логгер читает на :03.
"""
import os
import re
import time
from datetime import datetime, timedelta

from dotenv import load_dotenv
import xlwings as xw
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from tikers import load_tickers_from_excel

load_dotenv()

CHROMIUM_VERSION = 153
SCREENER_URL = os.getenv(
    "SCREENER_URL",
    "https://www.tradingview.com/screener/YOUR_SCREENER_ID/",
)

REFRESH_SETTLE_SEC = 1.5
COLD_START_WAIT_SEC = 15
SCROLL_PAUSE_SEC = 1.5
SCROLL_STABLE_ROUNDS = 10

RVOL_HEADER = "Rel vol at time"
OUT_HEADERS = ["Symbol", RVOL_HEADER]

# Быстрый съём только двух колонок из DOM (без обхода каждой ячейки в Python)
EXTRACT_JS = """
const table = arguments[0];
const rvolWanted = (arguments[1] || '').toLowerCase();

const ths = Array.from(table.querySelectorAll('thead tr th'));
const headers = [];
for (const th of ths) {
  const u = th.querySelector('[class*="upperLine"]');
  let t = (u ? u.textContent : th.textContent).trim();
  t = t.replace(/\\s+/g, ' ');
  if (t && t !== '+') headers.push(t);
}

let tickerIdx = 0;
let rvolIdx = -1;
headers.forEach((h, i) => {
  const l = h.toLowerCase();
  if (l.includes('symbol') || l.includes('ticker')) tickerIdx = i;
  if (l === rvolWanted) rvolIdx = i;
});

if (rvolIdx < 0) {
  return { ok: false, headers: headers, rows: [] };
}

const rows = [];
for (const tr of table.querySelectorAll('tbody tr')) {
  const tds = tr.querySelectorAll('td');
  if (tds.length <= Math.max(tickerIdx, rvolIdx)) continue;
  let sym = (tds[tickerIdx].innerText || '').trim();
  sym = sym.split(/\\n/)[0].trim();
  const rvol = (tds[rvolIdx].innerText || '').trim();
  if (sym) rows.push([sym, rvol]);
}
return { ok: true, headers: headers, rows: rows };
"""


def wait_until_minute_start():
    now = datetime.now()
    if now.second == 0 and now.microsecond < 50_000:
        return
    target = now.replace(second=0, microsecond=0) + timedelta(minutes=1)
    delay = (target - datetime.now()).total_seconds()
    if delay > 0:
        print(f"Ждём начала минуты {target.strftime('%H:%M:%S')} ({delay:.1f}с)...")
        time.sleep(delay)


def get_driver():
    options = uc.ChromeOptions()
    user_data_dir = r"D:\VS Project\playbook\chrome_profile_1"
    if not os.path.exists(user_data_dir):
        os.makedirs(user_data_dir)
    options.add_argument(f"--user-data-dir={user_data_dir}")
    options.add_argument("--profile-directory=Mike")
    return uc.Chrome(options=options, version_main=CHROMIUM_VERSION)


def scroll_to_load_all_tickers(driver, pause=1, stable_rounds=6):
    xpath_table = "//table[contains(@class, 'table-')]"
    print("Прокручиваем таблицу до конца для загрузки всех тикеров...")
    prev_count = -1
    idle = 0

    while idle < stable_rounds:
        table = driver.find_element(By.XPATH, xpath_table)
        rows = table.find_elements(By.XPATH, ".//tbody//tr")
        count = len(rows)

        if count == prev_count:
            idle += 1
        else:
            idle = 0
            prev_count = count
            if count:
                print(f"  Загружено строк: {count}")

        if rows:
            driver.execute_script(
                "arguments[0].scrollIntoView({block: 'end', behavior: 'instant'});",
                rows[-1],
            )
        driver.execute_script(
            """
            var el = arguments[0];
            while (el && el !== document.body) {
                if (el.scrollHeight > el.clientHeight + 10) {
                    el.scrollTop = el.scrollHeight;
                }
                el = el.parentElement;
            }
            window.scrollTo(0, document.body.scrollHeight);
            """,
            table,
        )
        time.sleep(pause)

    print(f"Прокрутка завершена. Строк в таблице: {prev_count}")
    return driver.find_element(By.XPATH, xpath_table)


def clean_ticker(text):
    if not text:
        return ""
    return re.split(r"\n|\s{2,}", str(text).strip())[0].strip()


def update_live_excel(data):
    """Пишет только Symbol + Rel vol at time в G:H, не трогает A-F."""
    try:
        wb = xw.Book("ATR_Monitor.xlsx")
        sheet = wb.sheets["Лист1"]

        last_row = sheet.range("G" + str(sheet.cells.last_cell.row)).end("up").row
        if last_row < 1:
            last_row = 1
        sheet.range(f"G1:H{max(last_row, len(data) + 2)}").clear_contents()

        sheet.range("G1").value = OUT_HEADERS
        if data:
            sheet.range("G2").value = data

        print(f"✅ G:H обновлены (тикеров: {len(data)})")
    except Exception as e:
        print(f"❌ Ошибка при записи в Excel: {e}")


def collect_rvol_for_watchlist(driver, watchlist):
    """
    JS: все Symbol + Rel vol at time со страницы,
    затем фильтр по watchlist (порядок как в Excel L).
    """
    wanted = {t.upper(): t for t in watchlist}
    table = driver.find_element(By.XPATH, "//table[contains(@class, 'table-')]")
    payload = driver.execute_script(EXTRACT_JS, table, RVOL_HEADER)

    if not payload or not payload.get("ok"):
        headers = (payload or {}).get("headers")
        raise RuntimeError(
            f"Колонка '{RVOL_HEADER}' не найдена. Заголовки: {headers}"
        )

    found = {}
    for sym, rvol in payload.get("rows") or []:
        key = clean_ticker(sym).upper()
        if key in wanted:
            found[key] = rvol

    rows = []
    missing = []
    for t in watchlist:
        key = t.upper()
        if key in found:
            rows.append([key, found[key]])
        else:
            missing.append(t)
    return rows, missing


def scrape_screener_to_excel():
    driver = get_driver()
    driver.maximize_window()

    try:
        print(f"Открываем страницу скринера: {SCREENER_URL}...")
        driver.get(SCREENER_URL)

        wait = WebDriverWait(driver, 40)
        wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//table[contains(@class, 'table-')]")
            )
        )
        print(f"Холодный старт: ждём {COLD_START_WAIT_SEC}с...")
        time.sleep(COLD_START_WAIT_SEC)

        print("Первичная прокрутка...")
        scroll_to_load_all_tickers(
            driver, pause=SCROLL_PAUSE_SEC, stable_rounds=SCROLL_STABLE_ROUNDS
        )
        print("Пауза 5с после прокрутки...")
        time.sleep(5)
        print("Повторная прокрутка...")
        scroll_to_load_all_tickers(
            driver,
            pause=SCROLL_PAUSE_SEC,
            stable_rounds=max(4, SCROLL_STABLE_ROUNDS // 2),
        )
        print("Холодный старт готов. Дальше — цикл по :00.")

        iteration = 1
        print("Цикл :00 → только Rel vol at time для тикеров из Лист2!L")
        while True:
            wait_until_minute_start()
            started = datetime.now()
            print(f"\n--- #{iteration} @ {started.strftime('%H:%M:%S')} ---")

            watchlist = load_tickers_from_excel()
            if not watchlist:
                print("⚠️ Список тикеров пуст (Лист2!L)")
                iteration += 1
                continue

            try:
                refresh_btn = wait.until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//button[@data-qa-id='screener-refresh-button']")
                    )
                )
                print("Refresh...")
                refresh_btn.click()
            except Exception as e:
                print(f"⚠️ Refresh не нажат: {e}")

            time.sleep(REFRESH_SETTLE_SEC)

            try:
                rows, missing = collect_rvol_for_watchlist(driver, watchlist)
                update_live_excel(rows)
                elapsed = (datetime.now() - started).total_seconds()
                msg = f"Готово за {elapsed:.1f}с, записано {len(rows)}/{len(watchlist)}"
                if missing:
                    msg += f", нет в скринере: {', '.join(missing)}"
                print(msg)
            except Exception as e:
                print(f"❌ Сбор/запись: {e}")

            iteration += 1

    except KeyboardInterrupt:
        print("\nСкрипт остановлен пользователем.")
    except Exception as e:
        print(f"❌ Критическая ошибка: {e}")
    finally:
        print("Закрываем Chromium...")
        driver.quit()


if __name__ == "__main__":
    scrape_screener_to_excel()
