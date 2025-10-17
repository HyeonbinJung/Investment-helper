import psycopg2
from datetime import date, timedelta
import yfinance as yf
import time
import pandas as pd
password_sql = 

def create_table(table_name, columns, dbname, user, password, host="localhost", port="5432"):
    column_defs = ", ".join([f"{col} {dtype}" for col, dtype in columns.items()])
    query = f"CREATE TABLE IF NOT EXISTS {table_name} ({column_defs});"
    try:
        conn = psycopg2.connect(
            dbname=dbname,
            user=user,
            password=password,
            host=host,
            port=port
        )
        cur = conn.cursor()
        cur.execute(query)
        conn.commit()
    except Exception as e:
        print("Connection error", e)
    finally:
        if conn:
            cur.close()
            conn.close()
            
columns = {
    "date": "VARCHAR(50)",        
    "ticker": "VARCHAR(50)",      
    "marketcap": "FLOAT",         
    "open_price": "FLOAT",        
    "close_price": "FLOAT",       
    "high_price": "FLOAT",        
    "low_price": "FLOAT"    
}

create_table("stock_data", columns, "testdb", "postgres", password_sql, host="localhost", port="5432")


def insert_stock_data(dbname, user, password, insert_date, insert_ticker, marketcap, open_price, close_price, high_price, low_price, host="localhost", port="5432"):
    query = """
    INSERT INTO stock_data (date, ticker, marketcap, open_price, close_price, high_price, low_price)
    VALUES (%s, %s, %s, %s, %s, %s, %s);
    """
    try:
        data = {
            "date": insert_date,
            "ticker": insert_ticker,
            "marketcap": marketcap,        
            "open_price": open_price,
            "close_price": close_price,
            "high_price": high_price,
            "low_price": low_price}

        stock_data = data
        
        conn = psycopg2.connect(
            dbname=dbname,
            user=user,
            password=password,
            host=host,
            port=port
        )
        cur = conn.cursor()
        cur.execute(query, (
            stock_data["date"],
            stock_data["ticker"],
            stock_data["marketcap"],
            stock_data["open_price"],
            stock_data["close_price"],
            stock_data["high_price"],
            stock_data["low_price"]
        ))
        conn.commit()
        print("Data inserted")
    except Exception as e:
        print("Error occur during inserting Data", e)
    finally:
        if conn:
            cur.close()
            conn.close()



def clear_table(dbname, user, password, table_name, host="localhost", port="5432"):
    query = f"DELETE FROM {table_name};"
    try:
        conn = psycopg2.connect(
            dbname=dbname,
            user=user,
            password=password,
            host=host,
            port=port
        )
        cur = conn.cursor()
        cur.execute(query)
        conn.commit()
        print(f"All data from '{table_name}' deleted.")
    except Exception as e:
        print("Error while clearing table:", e)
    finally:
        if conn:
            cur.close()
            conn.close()

def get_chart_sql(days, ticker):
    for i in range(days):
        reverse = days - i 
        date_to_insert = (date.today() - timedelta(days=reverse - 1))
        tick_to_insert = ticker

        stock = yf.Ticker(ticker)
        target_date = (date.today() - timedelta(days=reverse - 1))
        target_date_2 = (date.today() - timedelta(days=reverse - 2))
        
        try:
            marketcap_to_insert = stock.info.get("marketCap")
            open_to_insert = float(stock.history(start=target_date, end=target_date_2)["Open"].iloc[0])
            close_to_insert = float(stock.history(start=target_date, end=target_date_2)["Close"].iloc[0])
            high_to_insert = float(stock.history(start=target_date, end=target_date_2)["High"].iloc[0])
            low_to_insert = float(stock.history(start=target_date, end=target_date_2)["Low"].iloc[0])
            
            insert_stock_data("testdb", "postgres", password_sql, date_to_insert, tick_to_insert, marketcap_to_insert, open_to_insert, close_to_insert
                              , high_to_insert, low_to_insert, host="localhost", port="5432")
            
        except Exception as e:
            marketcap_to_insert = stock.info.get("marketCap")
            open_to_insert = 0
            close_to_insert = 0
            high_to_insert = 0
            low_to_insert = 0
            
            insert_stock_data("testdb", "postgres", password_sql, date_to_insert, tick_to_insert, marketcap_to_insert, open_to_insert, close_to_insert
                              , high_to_insert, low_to_insert, host="localhost", port="5432")

#insert_stock_data("testdb", "postgres", password_sql, "2024-10-15", "TSLA", 20, 20, 20, 20, 20, host="localhost", port="5432")
clear_table("testdb", "postgres", password_sql, "stock_data", host="localhost", port="5432")
#get_chart_sql(30, "TSLA")


