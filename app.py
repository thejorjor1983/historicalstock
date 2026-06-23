from io import BytesIO
from datetime import date

import pandas as pd
import streamlit as st
import yfinance as yf


# =========================
# INDICATOR HELPERS
# =========================

def calculate_rsi(close_prices, window=14):
    delta = close_prices.diff()

    gain = delta.where(delta > 0, 0)
    loss = -delta.where(delta < 0, 0)

    avg_gain = gain.rolling(window=window).mean()
    avg_loss = loss.rolling(window=window).mean()

    rs = avg_gain / avg_loss
    rsi = 100 - (100 / (1 + rs))

    return rsi


def interpret_rsi(rsi):
    if pd.isna(rsi):
        return ""
    if rsi > 70:
        return "Overbought / possibly stretched"
    if rsi >= 40:
        return "Reasonable range"
    if rsi < 30:
        return "Oversold / weak"
    return "Weak momentum"


def get_decision(score):
    if score >= 3:
        return "POSITIVE TECHNICAL SETUP"
    if score == 2:
        return "MIXED TECHNICAL SIGNALS"
    return "WEAK TECHNICAL SETUP"


# =========================
# MAIN ANALYSIS FUNCTION
# =========================

def create_historical_stock_dataframe(ticker, start_date, end_date, progress_callback=None):
    ticker = ticker.strip().upper()

    if not ticker:
        raise ValueError("Please enter a ticker.")

    if start_date >= end_date:
        raise ValueError("The end date must be after the start date.")

    if progress_callback:
        progress_callback(f"Downloading historical data for {ticker}...")

    data = yf.download(
        ticker,
        start=start_date,
        end=end_date,
        auto_adjust=False,
        progress=False,
    )

    if data.empty:
        raise ValueError(f"No data found for {ticker}. Please check the ticker symbol or dates.")

    if isinstance(data.columns, pd.MultiIndex):
        data.columns = data.columns.get_level_values(0)

    required_columns = {"Open", "High", "Low", "Close", "Volume"}
    missing_columns = required_columns - set(data.columns)
    if missing_columns:
        raise ValueError(f"Missing expected columns from Yahoo Finance: {', '.join(sorted(missing_columns))}")

    if progress_callback:
        progress_callback("Calculating indicators...")

    # =========================
    # CALCULATE INDICATORS
    # =========================

    data["SMA_50"] = data["Close"].rolling(window=50).mean()
    data["SMA_200"] = data["Close"].rolling(window=200).mean()

    data["RSI_14"] = calculate_rsi(data["Close"], 14)

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
        axis=1,
    )

    data["SMA_50_VS_SMA_200"] = data.apply(
        lambda row: "SMA 50 above SMA 200"
        if pd.notna(row["SMA_50"])
        and pd.notna(row["SMA_200"])
        and row["SMA_50"] > row["SMA_200"]
        else "SMA 50 below SMA 200",
        axis=1,
    )

    data["RSI_INTERPRETATION"] = data["RSI_14"].apply(interpret_rsi)

    data["MACD_INTERPRETATION"] = data.apply(
        lambda row: "MACD above signal - positive momentum"
        if pd.notna(row["MACD"])
        and pd.notna(row["MACD_SIGNAL"])
        and row["MACD"] > row["MACD_SIGNAL"]
        else "MACD below signal - weak momentum",
        axis=1,
    )

    # =========================
    # SCORE COLUMNS
    # =========================

    data["SCORE_PRICE_ABOVE_SMA_50"] = data.apply(
        lambda row: 1
        if pd.notna(row["SMA_50"]) and row["Close"] > row["SMA_50"]
        else 0,
        axis=1,
    )

    data["SCORE_SMA_50_ABOVE_SMA_200"] = data.apply(
        lambda row: 1
        if pd.notna(row["SMA_50"])
        and pd.notna(row["SMA_200"])
        and row["SMA_50"] > row["SMA_200"]
        else 0,
        axis=1,
    )

    data["SCORE_RSI_REASONABLE"] = data["RSI_14"].apply(
        lambda rsi: 1 if pd.notna(rsi) and 40 <= rsi <= 70 else 0
    )

    data["SCORE_MACD_ABOVE_SIGNAL"] = data.apply(
        lambda row: 1
        if pd.notna(row["MACD"])
        and pd.notna(row["MACD_SIGNAL"])
        and row["MACD"] > row["MACD_SIGNAL"]
        else 0,
        axis=1,
    )

    data["TOTAL_SCORE"] = (
        data["SCORE_PRICE_ABOVE_SMA_50"]
        + data["SCORE_SMA_50_ABOVE_SMA_200"]
        + data["SCORE_RSI_REASONABLE"]
        + data["SCORE_MACD_ABOVE_SIGNAL"]
    )

    data["TECHNICAL_DECISION"] = data["TOTAL_SCORE"].apply(get_decision)

    # =========================
    # PREPARE FOR EXCEL / DISPLAY
    # =========================

    data.reset_index(inplace=True)

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

    existing_columns = [column for column in preferred_columns if column in data.columns]
    return data[existing_columns]


def dataframe_to_excel_bytes(dataframe):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name="Historical Data")

        worksheet = writer.sheets["Historical Data"]
        for column_cells in worksheet.columns:
            max_length = 0
            column_letter = column_cells[0].column_letter

            for cell in column_cells:
                value = "" if cell.value is None else str(cell.value)
                max_length = max(max_length, len(value))

            worksheet.column_dimensions[column_letter].width = min(max_length + 2, 45)

    output.seek(0)
    return output.getvalue()


# =========================
# STREAMLIT APPLICATION
# =========================

st.set_page_config(
    page_title="Historical Stock Data with Indicators",
    page_icon="📈",
    layout="wide",
)

st.title("Historical Stock Data with Indicators")
st.write("Enter a ticker and date range. The app will calculate SMA, RSI, MACD, scores, and export the result to Excel.")

with st.sidebar:
    st.header("Inputs")

    ticker = st.text_input("Ticker", value="AAPL").strip().upper()

    start_date = st.date_input(
        "From date",
        value=date(2020, 1, 1),
    )

    end_date = st.date_input(
        "To date",
        value=date.today(),
    )

    run_analysis = st.button("Download and prepare Excel", type="primary")

if run_analysis:
    log_messages = []

    def add_log(message):
        log_messages.append(message)

    try:
        with st.spinner("Working..."):
            result = create_historical_stock_dataframe(
                ticker=ticker,
                start_date=start_date,
                end_date=end_date,
                progress_callback=add_log,
            )

            excel_bytes = dataframe_to_excel_bytes(result)

        st.success("Done. Your file is ready.")

        if log_messages:
            with st.expander("Log"):
                for message in log_messages:
                    st.write(message)

        latest_row = result.tail(1)

        st.subheader("Latest Technical Decision")
        if not latest_row.empty:
            col1, col2, col3, col4 = st.columns(4)

            col1.metric("Ticker", ticker)
            col2.metric("Total Score", int(latest_row["TOTAL_SCORE"].iloc[0]))
            col3.metric("Decision", str(latest_row["TECHNICAL_DECISION"].iloc[0]))
            col4.metric("Last Close", f"{float(latest_row['Close'].iloc[0]):,.2f}")

        st.subheader("Preview")
        st.dataframe(result.tail(20), use_container_width=True)

        file_name = f"{ticker}_historical_data_with_indicators_and_score.xlsx"

        st.download_button(
            label="Download Excel file",
            data=excel_bytes,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as error:
        st.error(str(error))
else:
    st.info("Enter a ticker and dates, then click the button in the sidebar.")
