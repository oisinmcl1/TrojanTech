import pyodbc
import datetime
import decimal

def connect_SQL():
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

def close_SQL(sqlServer, sqlCursor):
    try:
        if sqlCursor:
            sqlCursor.close()
        
        if sqlServer:
            sqlServer.close()
        print("SQL Connection Closed.")
    
    except pyodbc.Error as ex:
        print(f"Failed to close SQL connection: {ex}")

def fetch_JobKeyID(sqlCursor):
    try:
        fetchQuery_JobKeyID = "SELECT JobKeyID FROM dbo.Heading WHERE SatelliteJobNumber = 'RB262';"
        sqlCursor.execute(fetchQuery_JobKeyID)
        row = sqlCursor.fetchone()
        
        if row:
            return row[0]
        
        else:
            print("No JobKeyID found for SatelliteJobNumber = 'RB262'")
            return None
    
    except pyodbc.Error as ex:
        print(f"SQL query failed: {ex}")
        return None

def get_Columns(sqlCursor, table):
    try:
        sqlCursor.execute(f"SELECT * FROM {table} WHERE 1=0")
        columns = [column[0] for column in sqlCursor.description]
        print(f"Column names in '{table}': {columns}")
        return columns
    
    except pyodbc.Error as ex:
        print(f"SQL query failed: {ex}")
        return None

def clean_data(value):
    if isinstance(value, str):
        value = value.strip()
        if value == '00:00:00':
            return None
        if value == '':
            return None
    elif isinstance(value, datetime.datetime):
        if value == datetime.datetime(1899, 12, 30) or value == datetime.datetime(1900, 1, 1):
            return None
        return value.strftime('%Y-%m-%d %H:%M:%S')
    elif isinstance(value, decimal.Decimal):
        return float(value)
    elif isinstance(value, bool):
        return 1 if value else 0
    return value

def fetch_SQL(sqlCursor, JobKeyID, table):
    columns = get_Columns(sqlCursor, table)
    if not columns:
        return None
    
    try:
        query = f"SELECT {', '.join(columns)} FROM {table} WHERE JobKeyID = ?"
        sqlCursor.execute(query, JobKeyID)
        row = sqlCursor.fetchone()
        
        if row:
            # Apply clean_data to each value in the row
            data = {col: clean_data(val) for col, val in zip(columns, row)}
            print("\nFetched Data:")
            print(f"{'Column':<30} | Value")
            print("-" * 50)
            for col, val in data.items():
                print(f"{col:<30} | {val}")
            return data
        
        else:
            print(f"No data found for JobKeyID = {JobKeyID}")
            return None
    
    except pyodbc.Error as ex:
        print(f"SQL query failed: {ex}")
        return None

def update_Access(data, mdb):
    try:
        connStr = (r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};" f"DBQ={mdb};")
        with pyodbc.connect(connStr) as accessConn:
            with accessConn.cursor() as accessCursor:
                if data:
                    columns = list(data.keys())
                    values = list(data.values())
                    
                    set_clause = ', '.join(f"[{col}] = ?" for col in columns)
                    query = f"UPDATE Heading2 SET {set_clause} WHERE JobKeyID = ?"
                    
                    # Log the final query and the number of columns vs values
                    print(f"Executing query: {query}")
                    print(f"With values: {values + [data['JobKeyID']]}")
                    print(f"Number of columns: {len(columns)}")
                    print(f"Number of values: {len(values) + 1}")  # +1 for JobKeyID
                    
                    accessCursor.execute(query, *values, data['JobKeyID'])
                    accessConn.commit()
                    print(f"Access database updated successfully for JobKeyID {data['JobKeyID']}.")
                
                else:
                    print("No data to update in Access database.")
    
    except pyodbc.Error as ex:
        print(f"Access Database connection or update failed: {ex}")

def main():
    sqlServer, sqlCursor = connect_SQL()
    if not sqlCursor:
        return

    JobKeyID = fetch_JobKeyID(sqlCursor)
    if JobKeyID:
        data = fetch_SQL(sqlCursor, JobKeyID, "dbo.Heading2")
        if data:
            update_Access(data, r"E:\BM\FuturaFrames\Orders\2024\Jun\Test\RB262.mdb")
    close_SQL(sqlServer, sqlCursor)

if __name__ == "__main__":
    main()
