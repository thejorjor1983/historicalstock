import yfinance as yf
import pandas as pd


# =========================
# USER INPUTS
# =========================

ticker = input("Enter stock ticker, for example AAPL, NVDA, MSFT: ").upper()

start_date = input("Enter start date, for example 2020-01-01: ")
end_date = input("Enter end date, for example 2026-06-15: ")


# =========================
# DOWNLOAD HISTORICAL DATA
# =========================

print(f"Downloading historical data for {ticker}...")

data = yf.download(
    ticker,
    start=start_date,
    end=end_date,
    auto_adjust=False
)


# =========================
# CHECK IF DATA EXISTS
# =========================

if data.empty:
    print(f"No data found for {ticker}. Please check the ticker symbol or dates.")

else:
    # If yfinance returns multi-level columns, flatten them.
    if isinstance(data.columns, pd.MultiIndex):
        data.columns = data.columns.get_level_values(0)

    # =========================
    # CALCULATE INDICATORS
    # =========================

    # SMA 50 and SMA 200
    data["SMA_50"] = data["Close"].rolling(window=50).mean()
    data["SMA_200"] = data["Close"].rolling(window=200).mean()

    # RSI 14
    delta = data["Close"].diff()

    gain = delta.where(delta > 0, 0)
    loss = -delta.where(delta < 0, 0)

    avg_gain = gain.rolling(window=14).mean()
    avg_loss = loss.rolling(window=14).mean()

    rs = avg_gain / avg_loss
    data["RSI_14"] = 100 - (100 / (1 + rs))

    # MACD
    ema_12 = data["Close"].ewm(span=12, adjust=False).mean()
    ema_26 = data["Close"].ewm(span=26, adjust=False).mean()

    data["MACD"] = ema_12 - ema_26
    data["MACD_SIGNAL"] = data["MACD"].ewm(span=9, adjust=False).mean()
    data["MACD_HISTOGRAM"] = data["MACD"] - data["MACD_SIGNAL"]

    # =========================
    # INTERPRETATION COLUMNS
    # =========================

    data["PRICE_VS_SMA_50"] = data.apply(
        lambda row: "Above SMA 50"
        if pd.notna(row["SMA_50"]) and row["Close"] > row["SMA_50"]
        else "Below SMA 50",
        axis=1
    )

    data["SMA_50_VS_SMA_200"] = data.apply(
        lambda row: "SMA 50 above SMA 200"
        if pd.notna(row["SMA_50"]) and pd.notna(row["SMA_200"]) and row["SMA_50"] > row["SMA_200"]
        else "SMA 50 below SMA 200",
        axis=1
    )

    def interpret_rsi(rsi):
        if pd.isna(rsi):
            return ""
        elif rsi > 70:
            return "Overbought / possibly stretched"
        elif rsi >= 40:
            return "Reasonable range"
        elif rsi < 30:
            return "Oversold / weak"
        else:
            return "Weak momentum"

    data["RSI_INTERPRETATION"] = data["RSI_14"].apply(interpret_rsi)

    data["MACD_INTERPRETATION"] = data.apply(
        lambda row: "MACD above signal - positive momentum"
        if pd.notna(row["MACD"]) and pd.notna(row["MACD_SIGNAL"]) and row["MACD"] > row["MACD_SIGNAL"]
        else "MACD below signal - weak momentum",
        axis=1
    )

    # =========================
    # SCORE COLUMNS
    # =========================

    data["SCORE_PRICE_ABOVE_SMA_50"] = data.apply(
        lambda row: 1
        if pd.notna(row["SMA_50"]) and row["Close"] > row["SMA_50"]
        else 0,
        axis=1
    )

    data["SCORE_SMA_50_ABOVE_SMA_200"] = data.apply(
        lambda row: 1
        if pd.notna(row["SMA_50"]) and pd.notna(row["SMA_200"]) and row["SMA_50"] > row["SMA_200"]
        else 0,
        axis=1
    )

    data["SCORE_RSI_REASONABLE"] = data["RSI_14"].apply(
        lambda rsi: 1 if pd.notna(rsi) and 40 <= rsi <= 70 else 0
    )

    data["SCORE_MACD_ABOVE_SIGNAL"] = data.apply(
        lambda row: 1
        if pd.notna(row["MACD"]) and pd.notna(row["MACD_SIGNAL"]) and row["MACD"] > row["MACD_SIGNAL"]
        else 0,
        axis=1
    )

    data["TOTAL_SCORE"] = (
        data["SCORE_PRICE_ABOVE_SMA_50"]
        + data["SCORE_SMA_50_ABOVE_SMA_200"]
        + data["SCORE_RSI_REASONABLE"]
        + data["SCORE_MACD_ABOVE_SIGNAL"]
    )

    def get_decision(score):
        if score >= 3:
            return "POSITIVE TECHNICAL SETUP"
        elif score == 2:
            return "MIXED TECHNICAL SIGNALS"
        else:
            return "WEAK TECHNICAL SETUP"

    data["TECHNICAL_DECISION"] = data["TOTAL_SCORE"].apply(get_decision)

    # =========================
    # PREPARE FOR EXCEL
    # =========================

    # Make Date a normal Excel column instead of the index.
    data.reset_index(inplace=True)

    # Reorder columns so score is at the very right.
    preferred_columns = [
        "Date",
        "Open",
        "High",
        "Low",
        "Close",
        "Adj Close",
        "Volume",

        "SMA_50",
        "SMA_200",
        "RSI_14",
        "MACD",
        "MACD_SIGNAL",
        "MACD_HISTOGRAM",

        "PRICE_VS_SMA_50",
        "SMA_50_VS_SMA_200",
        "RSI_INTERPRETATION",
        "MACD_INTERPRETATION",

        "SCORE_PRICE_ABOVE_SMA_50",
        "SCORE_SMA_50_ABOVE_SMA_200",
        "SCORE_RSI_REASONABLE",
        "SCORE_MACD_ABOVE_SIGNAL",

        "TOTAL_SCORE",
        "TECHNICAL_DECISION",
    ]

    existing_columns = [
        column for column in preferred_columns
        if column in data.columns
    ]

    data = data[existing_columns]

    # =========================
    # SAVE TO EXCEL
    # =========================

    file_name = f"{ticker}_historical_data_with_indicators_and_score.xlsx"

    data.to_excel(file_name, index=False)

    print(f"Done! Historical data with indicators and score saved to: {file_name}")
