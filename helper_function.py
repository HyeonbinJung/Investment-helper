import xlwings as xw
import yfinance as yf
import time
import pandas as pd
from sklearn.linear_model import LinearRegression
import numpy as np

wb = xw.Book(r"C:\Users\sjung\Downloads\trading_helper.xlsx")
ws = wb.sheets[0]  # Short-term trading sheet
ws_peer = wb.sheets[3]


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

def clear_sheet(ws, range):
    ws.range(range).value = None

def clear_all(ws):
    clear_sheet(ws, "E11:J11")
    clear_sheet(ws, "E17:L17")
    clear_sheet(ws, "E25:I25")
    clear_sheet(ws, "E27:I27")
    clear_sheet(ws, "E34:I34")
    clear_sheet(ws, "E37:I37")
    clear_sheet(ws, "O29:O33")

def estimation_regression(cell, Revenue):
    x = np.array([ws["E25"].value,
                  ws["F25"].value,
                  ws["G25"].value,
                  ws["H25"].value,
                  ws["I25"].value]).reshape(-1, 1)
    y = np.array([ws["E11"].value,
                  ws["F11"].value,
                  ws["G11"].value,
                  ws["H11"].value,
                  ws["I11"].value])


    model = LinearRegression()
    model.fit(x, y)
    print(model.predict([[Revenue]]))
    ws[cell].value = model.predict([[Revenue]])

def get_marketcap(target_cell, cell_to_write):
    tickertarget = ws_peer[target_cell].value
    stock_target = yf.Ticker(tickertarget)
    ws_peer[cell_to_write].value = stock_target.info.get("marketCap")
