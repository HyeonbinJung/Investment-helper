import xlwings as xw
import yfinance as yf
import time
import pandas as pd
from sklearn.linear_model import LinearRegression
import numpy as np
from helper_function import add_history
from helper_function import get_quarterly_revenue
from helper_function import get_quarterly_costofrev
from helper_function import get_quarterly_EBIT
from helper_function import get_quarterly_EBITDA
from helper_function import clear_sheet
from helper_function import clear_all
from helper_function import estimation_regression
from helper_function import get_marketcap


wb = xw.Book(r"C:\Users\sjung\Downloads\trading_helper.xlsx")
ws = wb.sheets[0]  # Short-term trading sheet
ws_peer = wb.sheets[3]

clear_all(ws)



# Key Financials (Income Statement)
ticker = ws["E2"].value
if ticker:
    stock = yf.Ticker(ticker)
    get_quarterly_revenue(ws, "E25", stock, ticker, "2024-06-30")
    get_quarterly_revenue(ws, "F25", stock, ticker, "2024-09-30")
    get_quarterly_revenue(ws, "G25", stock, ticker, "2024-12-31")
    get_quarterly_revenue(ws, "H25", stock, ticker, "2025-03-31")
    get_quarterly_revenue(ws, "I25", stock, ticker, "2025-06-30")
    get_quarterly_costofrev(ws, "E27", stock, ticker, "2024-06-30")
    get_quarterly_costofrev(ws, "F27", stock, ticker, "2024-09-30")
    get_quarterly_costofrev(ws, "G27", stock, ticker, "2024-12-31")
    get_quarterly_costofrev(ws, "H27", stock, ticker, "2025-03-31")
    get_quarterly_costofrev(ws, "I27", stock, ticker, "2025-06-30")
    get_quarterly_EBIT(ws, "E34", stock, ticker, "2024-06-30")
    get_quarterly_EBIT(ws, "F34", stock, ticker, "2024-09-30")
    get_quarterly_EBIT(ws, "G34", stock, ticker, "2024-12-31")
    get_quarterly_EBIT(ws, "H34", stock, ticker, "2025-03-31")
    get_quarterly_EBIT(ws, "I34", stock, ticker, "2025-06-30")
    get_quarterly_EBITDA(ws, "E37", stock, ticker, "2024-06-30")
    get_quarterly_EBITDA(ws, "F37", stock, ticker, "2024-09-30")
    get_quarterly_EBITDA(ws, "G37", stock, ticker, "2024-12-31")
    get_quarterly_EBITDA(ws, "H37", stock, ticker, "2025-03-31")
    get_quarterly_EBITDA(ws, "I37", stock, ticker, "2025-06-30")

# Peer valuation
get_marketcap("F18", "J18")
get_marketcap("F21", "J21")

# For Pest 5 years
ticker = ws["E2"].value
if ticker:
    stock = yf.Ticker(ticker)
    add_history(ws, "E11", stock, "2024-06-27", "2024-06-30")
    add_history(ws, "F11", stock, "2024-09-27", "2024-09-30")
    add_history(ws, "G11", stock, "2024-12-25", "2024-12-31")
    add_history(ws, "H11", stock, "2025-03-27", "2025-03-30")
    add_history(ws, "I11", stock, "2025-06-27", "2025-06-30")

# Stock price for 7 days
    date7 = ws["E16"].value
    date6 = ws["F16"].value
    date5 = ws["G16"].value
    date4 = ws["H16"].value
    date3 = ws["I16"].value
    date2 = ws["J16"].value
    date1 = ws["K16"].value
    date0 = ws["L16"].value
    today7 = stock.history(start=date7, end=date6)
    today6 = stock.history(start=date6, end=date5)
    today5 = stock.history(start=date5, end=date4)
    today4 = stock.history(start=date4, end=date3)
    today3 = stock.history(start=date3, end=date2)
    today2 = stock.history(start=date2, end=date1)
    today1 = stock.history(start=date1, end=date0)
    today0 = stock.history(start=date0, end=date0)

    if not today7.empty:
        price_today7 = today7["Close"].iloc[0]
        ws["E17"].value = price_today7
    if not today6.empty:
        price_today6 = today6["Close"].iloc[0]
        ws["F17"].value = price_today6
    if not today5.empty:
        price_today5 = today5["Close"].iloc[0]
        ws["G17"].value = price_today5
    if not today4.empty:
        price_today4 = today4["Close"].iloc[0]
        ws["H17"].value = price_today4
    if not today3.empty:
        price_today3 = today3["Close"].iloc[0]
        ws["I17"].value = price_today3
    if not today2.empty:
        price_today2 = today2["Close"].iloc[0]
        ws["J17"].value = price_today2
    if not today1.empty:
        price_today1 = today1["Close"].iloc[0]
        ws["K17"].value = price_today1
    if not today0.empty:
        price_today0 = today0["Close"].iloc[0]
        ws["L17"].value = price_today0


estimation_regression("O29", ws["M29"].value)
estimation_regression("O30", ws["M30"].value)
estimation_regression("O31", ws["M31"].value)
estimation_regression("O32", ws["M32"].value)
estimation_regression("O33", ws["M33"].value)

while True:
    ticker = ws["E2"].value 
    if ticker:
        stock = yf.Ticker(ticker)
        price = stock.history(period="1d")["Close"].iloc[-1]
        ws["E7"].value = price  
        ws["F2"].value = time.strftime("%H:%M:%S")  
        print(f"{ticker}: {price}")
    time.sleep(5)


