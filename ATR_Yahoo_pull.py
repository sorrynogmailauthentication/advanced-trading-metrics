import pandas as pd
import xlwings as xw
import yfinance as yf
from tikers import TICKERS

# Список из 20 ликвидных тикеров (исключён делистнутый SNDK, добавлены QCOM, ASML, AMAT)
tickers = TICKERS

# Попытка подключиться к уже открытому Excel-файлу
try:
    wb = xw.Book("ATR_Monitor.xlsx")
    sheet = wb.sheets["Лист1"]
except Exception:
    print(
        "Ошибка: Пожалуйста, откройте файл 'ATR_Monitor.xlsx' в Excel перед запуском!"
    )
    exit()

report_data = []
print("Получение данных из Yahoo Finance и расчет IBKR SMA ATR...")

for ticker in tickers:
    try:
        # Загрузка чистых сырых данных без авто-корректировок (для точного матча с IBKR)
        df = yf.download(
            ticker, period="3mo", auto_adjust=False, back_adjust=False
        )

        if df.empty:
            continue

        if isinstance(df.columns, pd.MultiIndex):
            df.columns = df.columns.get_level_values(0)

        # Расчет компонентов True Range
        high_low = df["High"] - df["Low"]
        high_prev_close = (df["High"] - df["Close"].shift(1)).abs()
        low_prev_close = (df["Low"] - df["Close"].shift(1)).abs()

        df["True_Range"] = pd.concat(
            [high_low, high_prev_close, low_prev_close], axis=1
        ).max(axis=1)

        # Стандартное сглаживание SMA (как по умолчанию в IBKR)
        df["ATR_14"] = df["True_Range"].rolling(window=14).mean()

        latest_row = df.iloc[-1]
        report_data.append(
            {
                "Тикер": ticker,
                "Дата": df.index[-1].strftime("%Y-%m-%d"),
                "Close": round(float(latest_row["Close"]), 2),
                "14-Day ATR (IBKR)": round(float(latest_row["ATR_14"]), 2),
            }
        )
    except Exception as e:
        print(f"Ошибка при обработке {ticker}: {e}")

# Превращаем результат в DataFrame
output_df = pd.DataFrame(report_data)

# Запись в Excel начиная с ячейки D4
# index=False исключает запись номеров строк pandas, оставляя только чистые данные
sheet.range("A1").options(index=False, header=True).value = output_df

print("Данные успешно обновлены в Excel, начиная с ячейки D4!")
