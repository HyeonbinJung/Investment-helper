import xlwings as xw
import yfinance as yf
import time
import pandas as pd

wb = xw.Book(r"C:\Users\sjung\Downloads\trading_helper.xlsx")
ws = wb.sheets[0]  # Short-term trading sheet
ws_peer = wb.sheets[2]

# Helper Function
def add_history(ws, cell, stock, start, end): 
    ws[cell].value = stock.history(start=start, end=end)["Close"].iloc[0]

def get_quarterly_revenue(ws, cell, stock, ticker, target_quarter):
    df = stock.quarterly_financials
    if "Total Revenue" not in df.index:
        print("No Data found")
    target = pd.to_datetime(target_quarter)
    if target in df.columns:
        ws[cell].value = df.loc["Total Revenue", target]
    else:
        print("No Data found")
        
def get_quarterly_costofrev(ws, cell, stock, ticker, target_quarter):
    df = stock.quarterly_financials
    if "Cost Of Revenue" not in df.index:
        print("No Data found")
    target = pd.to_datetime(target_quarter)
    if target in df.columns:
        ws[cell].value = df.loc["Cost Of Revenue", target]
    else:
        print("No Data found")

def get_quarterly_EBIT(ws, cell, stock, ticker, target_quarter):
    df = stock.quarterly_financials
    if "EBIT" not in df.index:
        print("No Data found")
    target = pd.to_datetime(target_quarter)
    if target in df.columns:
        ws[cell].value = df.loc["EBIT", target]
    else:
        print("No Data found")

def get_quarterly_EBITDA(ws, cell, stock, ticker, target_quarter):
    df = stock.quarterly_financials
    if "EBITDA" not in df.index:
        print("No Data found")
    target = pd.to_datetime(target_quarter)
    if target in df.columns:
        ws[cell].value = df.loc["EBITDA", target]
    else:
        print("No Data found")
        
# Key Financials (Income Statement)
ticker = ws["E2"].value
if ticker:
    stock = yf.Ticker(ticker)
    get_quarterly_revenue(ws, "E25", stock, ticker, "2024-06-30")
    get_quarterly_revenue(ws, "F25", stock, ticker, "2024-09-30")
    get_quarterly_revenue(ws, "G25", stock, ticker, "2024-12-31")
    get_quarterly_revenue(ws, "H25", stock, ticker, "2025-03-31")
    get_quarterly_revenue(ws, "I25", stock, ticker, "2025-06-30")
    get_quarterly_costofrev(ws, "E26", stock, ticker, "2024-06-30")
    get_quarterly_costofrev(ws, "F26", stock, ticker, "2024-09-30")
    get_quarterly_costofrev(ws, "G26", stock, ticker, "2024-12-31")
    get_quarterly_costofrev(ws, "H26", stock, ticker, "2025-03-31")
    get_quarterly_costofrev(ws, "I26", stock, ticker, "2025-06-30")
    get_quarterly_EBIT(ws, "E33", stock, ticker, "2024-06-30")
    get_quarterly_EBIT(ws, "F33", stock, ticker, "2024-09-30")
    get_quarterly_EBIT(ws, "G33", stock, ticker, "2024-12-31")
    get_quarterly_EBIT(ws, "H33", stock, ticker, "2025-03-31")
    get_quarterly_EBIT(ws, "I33", stock, ticker, "2025-06-30")
    get_quarterly_EBITDA(ws, "E36", stock, ticker, "2024-06-30")
    get_quarterly_EBITDA(ws, "F36", stock, ticker, "2024-09-30")
    get_quarterly_EBITDA(ws, "G36", stock, ticker, "2024-12-31")
    get_quarterly_EBITDA(ws, "H36", stock, ticker, "2025-03-31")
    get_quarterly_EBITDA(ws, "I36", stock, ticker, "2025-06-30")

# Peer valuation
tickertarget = ws_peer["F16"].value
stock_target = yf.Ticker(tickertarget)
market_cap_target = stock_target.info.get("marketCap")
ws_peer["J16"].value = market_cap_target 

# For Pest 5 years
ticker = ws["E2"].value
if ticker:
    stock = yf.Ticker(ticker)
    add_history(ws, "E11", stock, "2021-12-25", "2022-01-01")
    add_history(ws, "F11", stock, "2022-12-25", "2023-01-01")
    add_history(ws, "G11", stock, "2023-12-25", "2024-01-01")
    add_history(ws, "H11", stock, "2024-12-25", "2025-01-01")
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




while True:
    ticker = ws["E2"].value 
    if ticker:
        stock = yf.Ticker(ticker)
        price = stock.history(period="1d")["Close"].iloc[-1]
        ws["E7"].value = price  
        ws["F2"].value = time.strftime("%H:%M:%S")  
        print(f"{ticker}: {price}")
    time.sleep(5)

