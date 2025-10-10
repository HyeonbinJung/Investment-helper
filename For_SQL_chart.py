import psycopg2

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


def insert_stock_data(dbname, user, password, stock_data, host="localhost", port="5432"):
    query = """
    INSERT INTO stock_data (date, ticker, marketcap, open_price, close_price, high_price, low_price)
    VALUES (%s, %s, %s, %s, %s, %s, %s);
    """
    try:
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

data = {
    "date": "2025-10-09",
    "ticker": "AAPL",
    "marketcap": 0.0,        
    "open_price": 230.15,
    "close_price": 233.80,
    "high_price": 234.20,
    "low_price": 229.50
}

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



