from enum import Enum
import time, os, sys
import math
import xlwings as xw
import yfinance as yf
from currency_converter import CurrencyConverter
import pandas as pd

print(
    """
==============================
Stock Portfolio Overview
==============================
"""
)

print(sys.path)

class Column(Enum):
    """ Column Name Translation from Excel, 1 = Column A, 2 = Column B, ... """

    longName = 1
    symbol = 2
    currentPrice = 5
    currency = 6
    conversionRate = 7
    open = 8
    dayLow = 9
    dayHigh = 10
    fiftyTwoWeekHigh = 12
    fiftyTwoWeekLow = 11
    fiftyDayAverage = 13
    twoHundredDayAverage = 14

def timestamp():
    t = time.localtime()
    timestamp = time.strftime("%b-%d-%Y_%H:%M:%S", t)
    return timestamp

def clearContentInExcel():
    """Clear the old contents in Excel"""
    if LAST_ROW > START_ROW:
        print(f"Clear Contents from row {START_ROW} to {LAST_ROW}")
        for data in Column:
            if not data.name == "symbol":
                sht.range((START_ROW, data.value), (LAST_ROW, data.value)).options(
                    expand="down"
                ).clear_contents()
        return None

def convert_to_target_currency(data, conversion_rate):
    return data * conversion_rate

def pullStockData():
    """
    Steps:
    1) Create an empty DataFrame
    2) Iterate over tickers, pull data from Yahoo Finance & add data to dictonary "new row"
    3) Append "new row" to DataFrame
    4) Return DataFrame
    """
    if tickers:
        print(f"Iterating over the following tickers: {tickers}")
        df = pd.DataFrame()
        for ticker in tickers:
            print(f"~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
            print(f"Pulling financial data for: {ticker} ...")
            data = yf.Ticker(ticker).get_info()
            try:
                open_price = data["open"]
            except KeyError:
                open_price = None

            # If no open price can be found, Yahoo Finance will return 'None'
            if open_price is None:
                # If opening price is None, append empty dataframe (row)
                print(f"Ticker: {ticker} not found on Yahoo Finance. Please check")
                df = df._append(pd.Series(dtype=str), ignore_index=True)
            else:
                try:
                    try:
                        long_name = data["longName"]
                    except (TypeError, KeyError):
                        long_name = None

                    try:
                        current_price = "%.2f" % convert_to_target_currency(
                            data["currentPrice"], conversion_rate
                        )
                    except:
                        current_price = None

                    yield_rel = None

                    ticker_currency = data["currency"]
                    conversion_rate = c.convert(1, ticker_currency, TARGET_CURRENCY)

                    new_row = {
                        "symbol": ticker,
                        "currency": ticker_currency,
                        "longName": long_name,
                        "conversionRate": "%.2f" % conversion_rate,
                        "open": "%.2f" % convert_to_target_currency(
                            open_price, conversion_rate
                        ),
                        "currentPrice": current_price,
                        "dayLow": "%.2f" % convert_to_target_currency(
                            data["dayLow"], conversion_rate
                        ),
                        "dayHigh": "%.2f" % convert_to_target_currency(
                            data["dayHigh"], conversion_rate
                        ),
                        "fiftyTwoWeekLow": "%.2f" % convert_to_target_currency(
                            data["fiftyTwoWeekLow"], conversion_rate
                        ),
                        "fiftyTwoWeekHigh": "%.2f" % convert_to_target_currency(
                            data["fiftyTwoWeekHigh"], conversion_rate
                        ),
                        "fiftyDayAverage": "%.2f" % convert_to_target_currency(
                            data["fiftyDayAverage"], conversion_rate
                        ),
                        "twoHundredDayAverage": "%.2f" % convert_to_target_currency(
                            data["twoHundredDayAverage"], conversion_rate
                        )
                    }
                    df = df._append(new_row, ignore_index=True)
                    print(f"Successfully pulled financial data for: {ticker}")

                except Exception as e:
                    # Error Handling
                    exc_type, exc_obj, exc_tb = sys.exc_info()
                    fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                    print(exc_type, fname, exc_tb.tb_lineno)
                    # Append Empty Row
                    df = df._append(pd.Series(dtype=str), ignore_index=True)
        return df
    return pd.DataFrame()

def writeValueToExcel(df):
    if not df.empty:
        print(f"~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        print(f"Writing data to Excel...")
        options = dict(index=False, header=False)
        for data in Column:
            if not data.name == "symbol":
                sht.range(START_ROW, data.value).options(**options).value = df[
                    data.name
                ]
        return None

def main():
    print(f"Please wait. The program is running ...")
    clearContentInExcel()
    df = pullStockData()
    writeValueToExcel(df)
    print(f"Program ran successfully!")

xw.Book("stock_portfolio.xlsm").set_mock_caller()
wb = xw.Book.caller()
sht = wb.sheets["Portfolio"]
show_msgbox = wb.macro("modMsgBox.ShowMsgBox")
TARGET_CURRENCY = sht.range("TARGET_CURRENCY").value
sht.range("TIMESTAMP").value = timestamp()
tickers = (
    # sht.range(START_ROW, Column.ticker.value).options(expand="down", numbers=str).value
    sht.range("B10").options(expand="down", numbers=str).value
)
START_ROW = sht.range("TICKER").row + 1  # Plus one row after the heading
LAST_ROW = sht.range(sht.cells.last_cell.row, Column.symbol.value).end("up").row
# LAST_ROW = sht.range("TICKER").row + len(tickers)
c = CurrencyConverter()

if __name__ == "__main__":
    main()
