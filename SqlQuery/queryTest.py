import pyodbc

def connect_to_sql_server():
    dsn = "Orders"
    database = "FFOrders"
    
    try:
        sqlServer = pyodbc.connect(f"DSN={dsn};DATABASE={database};")
        sqlCursor = sqlServer.cursor()
        print(f"Connection to SQL Server: {dsn} and Database: {database} was successful!")
        return sqlServer, sqlCursor
    
    except pyodbc.Error as ex:
        print(f"SQL Server connection failed: {ex}")
        return None, None

def close_sql_server_connection(sqlServer, sqlCursor):
    try:
        if sqlCursor:
            sqlCursor.close()
        
        if sqlServer:
            sqlServer.close()
        print("SQL Connection Closed.")
    
    except pyodbc.Error as ex:
        print(f"Failed to close SQL connection: {ex}")

def get_column_names(table_name):
    sqlServer, sqlCursor = connect_to_sql_server()
    
    if not sqlCursor:
        print("SQL Server connection is not available.")
        return None

    try:
        sqlCursor.execute(f"SELECT * FROM {table_name} WHERE 1=0")  # Select no rows, just get metadata
        column_names = [column[0] for column in sqlCursor.description]
        print(f"Column names in '{table_name}': {column_names}")
        return column_names

    except pyodbc.Error as ex:
        print(f"SQL query failed: {ex}")
        return None

    finally:
        close_sql_server_connection(sqlServer, sqlCursor)

table_name = "dbo.Heading2"
get_column_names(table_name)
