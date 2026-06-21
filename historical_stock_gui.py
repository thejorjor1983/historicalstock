import io
from datetime import date

import pandas as pd
import streamlit as st
import yfinance as yf


# =========================
# PAGE SETTINGS
# =========================

st.set_page_config(
    page_title="Historical Stock Data Generator",
    page_icon="📈",
    layout="wide",
)

st.title("📈 Historical Stock Data with Technical Indicators")

st.write(
    "Enter a stock ticker, choose a date range, and download an Excel file "
    "with historical prices, indicators, interpretations, and a technical score."
)


# =========================
# HELPER FUNCTIONS
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


def build_excel_file(dataframe):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name="Historical Data")

    output.seek(0)
    return output


def download_stock_data(ticker, start_date, end_date):
    data = yf.download(
        ticker,
        start=start_date,
        end=end_date,
        auto_adjust=False,
        progress=False,
    )

    if data.empty:
        return data

    if isinstance(data.columns, pd.MultiIndex):
        data.columns = data.columns.get_level_values(0)

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
        if (
            pd.notna(row["SMA_50"])
            and pd.notna(row["SMA_200"])
            and row["SMA_50"] > row["SMA_200"]
        )
        else "SMA 50 below SMA 200",
        axis=1,
    )

    data["RSI_INTERPRETATION"] = data["RSI_14"].apply(interpret_rsi)

    data["MACD_INTERPRETATION"] = data.apply(
        lambda row: "MACD above signal - positive momentum"
        if (
            pd.notna(row["MACD"])
            and pd.notna(row["MACD_SIGNAL"])
            and row["MACD"] > row["MACD_SIGNAL"]
        )
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
        if (
            pd.notna(row["SMA_50"])
            and pd.notna(row["SMA_200"])
            and row["SMA_50"] > row["SMA_200"]
        )
        else 0,
        axis=1,
    )

    data["SCORE_RSI_REASONABLE"] = data["RSI_14"].apply(
        lambda rsi: 1 if pd.notna(rsi) and 40 <= rsi <= 70 else 0
    )

    data["SCORE_MACD_ABOVE_SIGNAL"] = data.apply(
        lambda row: 1
        if (
            pd.notna(row["MACD"])
            and pd.notna(row["MACD_SIGNAL"])
            and row["MACD"] > row["MACD_SIGNAL"]
        )
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

    # Make Date a normal column
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

    existing_columns = [
        column for column in preferred_columns
        if column in data.columns
    ]

    return data[existing_columns]


# =========================
# USER INPUTS
# =========================

col1, col2, col3 = st.columns(3)

with col1:
    ticker = st.text_input(
        "Ticker",
        value="AAPL",
        help="Example: AAPL, NVDA, MSFT, RKLB",
    ).upper().strip()

with col2:
    start_date = st.date_input(
        "From date",
        value=date(2020, 1, 1),
    )

with col3:
    end_date = st.date_input(
        "To date",
        value=date.today(),
    )


# =========================
# RUN BUTTON
# =========================

if st.button("Generate Excel File", type="primary"):
    if not ticker:
        st.error("Please enter a stock ticker.")

    elif start_date >= end_date:
        st.error("The start date must be before the end date.")

    else:
        with st.spinner(f"Downloading historical data for {ticker}..."):
            try:
                result = download_stock_data(
                    ticker=ticker,
                    start_date=start_date,
                    end_date=end_date,
                )

                if result.empty:
                    st.error(
                        f"No data found for {ticker}. "
                        "Please check the ticker symbol or dates."
                    )

                else:
                    st.success(f"Done! Data generated for {ticker}.")

                    st.subheader("Preview")
                    st.dataframe(result.tail(20), use_container_width=True)

                    excel_file = build_excel_file(result)

                    file_name = f"{ticker}_historical_data_with_indicators_and_score.xlsx"

                    st.download_button(
                        label="Download Excel File",
                        data=excel_file,
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                    latest_row = result.iloc[-1]

                    st.subheader("Latest Technical Summary")

                    summary_col1, summary_col2, summary_col3, summary_col4 = st.columns(4)

                    with summary_col1:
                        st.metric("Close", f"${latest_row['Close']:.2f}")

                    with summary_col2:
                        st.metric("RSI 14", f"{latest_row['RSI_14']:.2f}")

                    with summary_col3:
                        st.metric("Total Score", f"{latest_row['TOTAL_SCORE']}/4")

                    with summary_col4:
                        st.metric("Decision", latest_row["TECHNICAL_DECISION"])

            except (ValueError, KeyError, TypeError, AttributeError) as error:
                st.error(f"Something went wrong: {error}")
