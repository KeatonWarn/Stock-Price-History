import yfinance as yf
import pandas as pd
from openpyxl import Workbook

def get_stock_data(ticker):
    stock = yf.Ticker(ticker)
    market_cap = stock.info['marketCap']
    hist = stock.history(period="max")
    return market_cap, hist

def save_to_excel(ticker, market_cap, hist_data):
    file_name = f"{ticker}_stock_data.xlsx"
    hist_data = hist_data.reset_index()[['Date', 'Close']]
    hist_data.columns = ['Date', 'Price']

#   #Set the tzinfo attribute of the Date column to None
    hist_data['Date'] = hist_data['Date'].apply(lambda x: x.replace(tzinfo=None))

    wb = Workbook()
    ws = wb.active
    ws.title = f"{ticker} Stock Data"
    ws.append(["Ticker", ticker])
    ws.append(["Market Capitalization", market_cap])

    ws.append([])
    ws.append(['Date', 'Price'])

    for index, row in hist_data.iterrows():
        ws.append([row['Date'], row['Price']])

    wb.save(file_name)
    print(f"{ticker} stock data saved to {file_name}")

if __name__ == "__main__":
    ticker = input("Enter a stock ticker: ").strip()
    market_cap, hist_data = get_stock_data(ticker)
    save_to_excel(ticker, market_cap, hist_data)
