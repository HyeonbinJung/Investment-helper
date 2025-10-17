
import psycopg2
from For_SQL_chart import create_table

password_sql = 
columns = {
    "Date": "VARCHAR(50)",        
    "ticker": "VARCHAR(50)",                  
    "stock_Price": "FLOAT",
    "number_of_stock": "INTEGER"
}

# Create Database for Portfolio
create_table("portfolio", columns, "testdb", "postgres", password_sql, host="localhost", port="5432")

# Return Purchased price / total investment / number of stock
def get_portfolio(ticker, return_type):
    query = f"""
        SELECT {return_type}
        FROM portfolio
        WHERE ticker ILIKE %s;
    """
    try:
        conn = psycopg2.connect(
            dbname="testdb",
            user="postgres",
            password=password_sql,
            host="localhost",
            port="5432"
        )
        cur = conn.cursor()
        cur.execute(query, (ticker,))
        rows = cur.fetchall()
        for row in rows:
            return(row[0])
    except Exception as e:
        print("Error:", e)
    finally:
        if conn:
            cur.close()
            conn.close()

def get_total_investment(ticker):
    return get_portfolio(ticker, stock_price) * get_portfolio(ticker, number_of_stock)

get_portfolio("TSLA", "stock_Price")



    
