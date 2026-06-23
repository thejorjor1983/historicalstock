import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext

import pandas as pd
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

def create_historical_stock_excel(ticker, start_date, end_date, output_file, progress_callback=None):
    ticker = ticker.strip().upper()
    start_date = start_date.strip()
    end_date = end_date.strip()

    if not ticker:
        raise ValueError("Please enter a ticker.")
    if not start_date:
        raise ValueError("Please enter a start date, for example 2020-01-01.")
    if not end_date:
        raise ValueError("Please enter an end date, for example 2026-06-15.")
    if not output_file:
        raise ValueError("Please choose where to save the Excel file.")

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
    # PREPARE FOR EXCEL
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
    data = data[existing_columns]

    if progress_callback:
        progress_callback("Saving Excel file...")

    output_path = Path(output_file)
    data.to_excel(output_path, index=False)

    if progress_callback:
        progress_callback(f"Done! File saved to: {output_path}")

    return output_path


# =========================
# GUI APPLICATION
# =========================

class HistoricalStockGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Historical Stock Data with Indicators")
        self.root.geometry("820x520")
        self.root.minsize(720, 450)

        self.ticker_var = tk.StringVar()
        self.start_date_var = tk.StringVar(value="2020-01-01")
        self.end_date_var = tk.StringVar(value="2026-06-15")
        self.output_file_var = tk.StringVar()

        self.run_button = None
        self.browse_button = None
        self.status_label = None
        self.output_box = None

        self.create_widgets()

    def create_widgets(self):
        main_frame = tk.Frame(self.root, padx=12, pady=12)
        main_frame.pack(fill="both", expand=True)
        main_frame.columnconfigure(1, weight=1)

        tk.Label(main_frame, text="Ticker:").grid(row=0, column=0, sticky="w", pady=5)
        tk.Entry(main_frame, textvariable=self.ticker_var, width=20).grid(
            row=0, column=1, sticky="w", pady=5
        )

        tk.Label(main_frame, text="From date:").grid(row=1, column=0, sticky="w", pady=5)
        tk.Entry(main_frame, textvariable=self.start_date_var, width=20).grid(
            row=1, column=1, sticky="w", pady=5
        )
        tk.Label(main_frame, text="Example: 2020-01-01").grid(row=1, column=2, sticky="w", padx=8)

        tk.Label(main_frame, text="To date:").grid(row=2, column=0, sticky="w", pady=5)
        tk.Entry(main_frame, textvariable=self.end_date_var, width=20).grid(
            row=2, column=1, sticky="w", pady=5
        )
        tk.Label(main_frame, text="Example: 2026-06-15").grid(row=2, column=2, sticky="w", padx=8)

        tk.Label(main_frame, text="Save Excel as:").grid(row=3, column=0, sticky="w", pady=5)
        tk.Entry(main_frame, textvariable=self.output_file_var).grid(
            row=3, column=1, sticky="ew", pady=5
        )
        self.browse_button = tk.Button(main_frame, text="Browse", command=self.choose_output_file)
        self.browse_button.grid(row=3, column=2, sticky="w", padx=8)

        button_frame = tk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=3, sticky="w", pady=(12, 8))

        self.run_button = tk.Button(
            button_frame,
            text="Download and Save Excel",
            command=self.start_download_thread,
            height=2,
            width=24,
        )
        self.run_button.pack(side="left")

        tk.Button(
            button_frame,
            text="Clear Log",
            command=self.clear_log,
            height=2,
            width=14,
        ).pack(side="left", padx=8)

        self.status_label = tk.Label(main_frame, text="Ready", anchor="w")
        self.status_label.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(5, 5))

        self.output_box = scrolledtext.ScrolledText(main_frame, wrap="word", height=16)
        self.output_box.grid(row=6, column=0, columnspan=3, sticky="nsew")
        main_frame.rowconfigure(6, weight=1)

    def choose_output_file(self):
        ticker = self.ticker_var.get().strip().upper() or "STOCK"
        default_name = f"{ticker}_historical_data_with_indicators_and_score.xlsx"

        filename = filedialog.asksaveasfilename(
            title="Save Excel file as",
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )

        if filename:
            self.output_file_var.set(filename)

    def append_log(self, message):
        self.output_box.insert(tk.END, message + "\n")
        self.output_box.see(tk.END)

    def set_status(self, message):
        self.status_label.config(text=message)

    def clear_log(self):
        self.output_box.delete("1.0", tk.END)
        self.set_status("Ready")

    def start_download_thread(self):
        ticker = self.ticker_var.get().strip().upper()

        if not ticker:
            messagebox.showerror("Missing ticker", "Please enter a ticker, for example AAPL.")
            return

        if not self.output_file_var.get().strip():
            default_file = Path.cwd() / f"{ticker}_historical_data_with_indicators_and_score.xlsx"
            self.output_file_var.set(str(default_file))

        self.run_button.config(state="disabled")
        self.set_status("Working...")
        self.append_log(f"Starting download for {ticker}...")

        worker = threading.Thread(target=self.run_download, daemon=True)
        worker.start()

    def gui_progress(self, message):
        self.root.after(0, lambda: self.append_log(message))

    def run_download(self):
        try:
            output_path = create_historical_stock_excel(
                ticker=self.ticker_var.get(),
                start_date=self.start_date_var.get(),
                end_date=self.end_date_var.get(),
                output_file=self.output_file_var.get(),
                progress_callback=self.gui_progress,
            )

            self.root.after(0, lambda: self.set_status("Done."))
            self.root.after(0, lambda: messagebox.showinfo("Done", f"Excel file saved to:\n{output_path}"))

        except (FileNotFoundError, PermissionError, ValueError, KeyError, TypeError, OSError) as error:
            self.root.after(0, lambda: self.set_status("Error"))
            self.root.after(0, lambda: self.append_log(f"ERROR: {error}"))
            self.root.after(0, lambda: messagebox.showerror("Error", str(error)))

        finally:
            self.root.after(0, lambda: self.run_button.config(state="normal"))


if __name__ == "__main__":
    root_window = tk.Tk()
    app = HistoricalStockGUI(root_window)
    root_window.mainloop()
