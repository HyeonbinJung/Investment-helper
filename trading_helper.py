import xlwings as xw
import yfinance as yf
import time

wb = xw.Book(r"C:\Users\sjung\Downloads\trading_helper.xlsx")
ws = wb.sheets[0]  # Short-term trading sheet
ws_peer = wb.sheets[2]

#Peer valuation
tickertarget = ws_peer["F16"].value
stock_target = yf.Ticker(tickertarget)
market_cap_target = stock_target.info.get("marketCap")
ws_peer["J16"].value = market_cap_target 


ticker = ws["E2"].value
if ticker:
    stock = yf.Ticker(ticker)

    fy2021 = stock.history(start="2021-12-25", end="2022-01-01")
    pricefy2021 = fy2021["Close"].iloc[0]
    ws["E11"].value = pricefy2021

    fy2022 = stock.history(start="2022-12-25", end="2023-01-01")
    pricefy2022 = fy2022["Close"].iloc[0]
    ws["F11"].value = pricefy2022

    fy2023 = stock.history(start="2023-12-25", end="2024-01-01")
    pricefy2023 = fy2023["Close"].iloc[0]
    ws["G11"].value = pricefy2023

    fy2024 = stock.history(start="2024-12-25", end="2025-01-01")
    pricefy2024 = fy2024["Close"].iloc[0]
    ws["H11"].value = pricefy2024

    half_2025 = stock.history(start="2025-06-27", end="2025-06-30")
    price_half_2025 = half_2025["Close"].iloc[0]
    ws["I11"].value = price_half_2025

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
