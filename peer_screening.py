import xlwings as xw
import yfinance as yf
import time
import pandas as pd
from sklearn.linear_model import LinearRegression
import numpy as np

wb = xw.Book(r"C:\Users\sjung\OneDrive\바탕 화면\새 폴더 (4)\trading_helper.xlsx")

def get_marketcap(page, ticker):
    ws = wb.sheets[page]
    stock_target = yf.Ticker(ticker)
    print (stock_target.info.get("marketCap"))
    data = yf.download(ticker, period="3mo", interval="1d")
    print(data)
    data2 = stock_target.quarterly_financials
    print(data2)


get_marketcap(4, "TSLA")
