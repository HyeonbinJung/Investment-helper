from datetime import date, timedelta
import xlwings as xw
import yfinance as yf
import time
from openpyxl.utils import get_column_letter, column_index_from_string
import pandas as pd

wb = xw.Book(r"C:\Users\sjung\OneDrive\바탕 화면\새 폴더 (4)\trading_helper.xlsx")

stock = yf.Ticker("TSLA")
hist = stock.history(start="2025-08-20", end="2025-09-08")
print(hist.head())

def get_chart(days, page, cell_Alpha, cell_num, ticker):
    ws = wb.sheets[page]
    col_num = column_index_from_string(cell_Alpha)
    stock = yf.Ticker(ticker)
    for i in range(days):
        reverse = days - i 
        alpha = get_column_letter(col_num + i)
        insert_cell = alpha + str(cell_num)
        stock_cell = alpha + str(cell_num + 1)
        ws[insert_cell].value = (date.today() - timedelta(days=reverse))
        target_date = (date.today() - timedelta(days=reverse))
        target_date_2 = (date.today() - timedelta(days=reverse - 1))
        try:
            stock_price = stock.history(start=target_date, end=target_date_2)["Close"].iloc[0]
            ws[stock_cell].value = stock.history(start=target_date, end=target_date_2)["Close"].iloc[0]
        except Exception as e:
            ws[stock_cell].value = False
            
        
get_chart(30, 4, "A", 1, "TSLA")



