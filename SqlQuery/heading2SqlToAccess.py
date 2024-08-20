import pyodbc
import datetime

sqlServer = None
sqlCursor = None

def connect_SQL():
    global sqlServer, sqlCursor
    
    dsn = "Orders"
    database = "FFOrders"
    
    try:
        sqlServer = pyodbc.connect(f"DSN={dsn};DATABASE={database};")
        sqlCursor = sqlServer.cursor()
        print(f"Connection to SQL Server: {dsn} and Database: {database} was successful!")
    
    except pyodbc.Error as ex:
        print(f"SQL Server connection failed: {ex}")

def close_SQL():
    global sqlServer, sqlCursor
    
    try:
        if sqlCursor:
            sqlCursor.close()
        
        if sqlServer:
            sqlServer.close()
        print("SQL Connection Closed.")
    
    except pyodbc.Error as ex:
        print(f"Failed to close SQL connection: {ex}")


def fetch_JobKeyID():
    global sqlCursor
    
    if not sqlCursor:
        print("SQL Server connection is not available.")
        return None

    try:
        fetchQuery_JobKeyID = """
        SELECT JobKeyID
        FROM dbo.Heading
        WHERE SatelliteJobNumber = 'RB262';
        """
        
        sqlCursor.execute(fetchQuery_JobKeyID)
        row = sqlCursor.fetchone()

        if row:
            JobKeyID = row[0]
            print(f"Fetched JobKeyID: {JobKeyID}")
            return JobKeyID
        else:
            print("No JobKeyID found for SatelliteJobNumber = 'RB262'")
            return None

    except pyodbc.Error as ex:
        print(f"SQL query failed: {ex}")
        return None


def get_Columns(table):
    global sqlCursor
    
    if not sqlCursor:
        print("SQL Server connection is not available.")
        return None

    try:
        sqlCursor.execute(f"SELECT * FROM {table} WHERE 1=0")
        columns = [column[0] for column in sqlCursor.description]
        
        print(f"Column names in '{table}': {columns}")
        return columns

    except pyodbc.Error as ex:
        print(f"SQL query failed: {ex}")
        return None


def clean_data(value):
    if isinstance(value, datetime.datetime):
        if value == datetime.datetime(1899, 12, 30) or value == datetime.datetime(1900, 1, 1):
            return None
    return value


def fetch_SQL(JobKeyID, table):
    global sqlCursor
    
    if not sqlCursor:
        print("SQL Server connection is not available.")
        return None

    try:
        columns = get_Columns(table)
        if not columns:
            return None

        query = f"SELECT {', '.join(columns)} FROM {table} WHERE JobKeyID = ?"
        sqlCursor.execute(query, JobKeyID)
        row = sqlCursor.fetchone()

        if row:
            data = dict(zip(columns, [clean_data(val) for val in row]))
            print("\nFetched Data:")
            print(f"{'Column':<30} | Value")
            print("-" * 50)
            
            for column, value in data.items():
                print(f"{column:<30} | {value}")
            return data
        else:
            print(f"No data found for JobKeyID = {JobKeyID}")
            return None

    except pyodbc.Error as ex:
        print(f"SQL query failed: {ex}")
        return None

def update_Access(data, mdb):
    try:
        connStr = (
            r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
            f"DBQ={mdb};"
        )

        with pyodbc.connect(connStr) as accessConn:
            with accessConn.cursor() as accessCursor:
                print(f"Connection to Access Database: {mdb} was successful!")

                if data:
                    columns = list(data.keys())
                    values = list(data.values())
                    
                    # Prepare the placeholders for SQL update statement
                    set_clause = ', '.join(f"{col} = ?" for col in columns)
                    query = f"UPDATE Heading2 SET {set_clause} WHERE JobKeyID = ?"
                    
                    print(f"Executing query: {query}")
                    print(f"With values: {values + [data['JobKeyID']]}")
                    
                    accessCursor.execute(query, *values, data['JobKeyID'])
                    accessConn.commit()
                    print(f"Access database updated successfully for JobKeyID {data['JobKeyID']}.")
                else:
                    print("No data to update in Access database.")

    except pyodbc.Error as ex:
        print(f"Access Database connection or update failed: {ex}")

    finally:
        try:
            accessCursor.close()
            accessConn.close()
        except NameError:
            pass
        print("Access Connection Closed.")


connect_SQL()

JobKeyID = fetch_JobKeyID()

if JobKeyID is not None:
    print(f"JobKeyID: {JobKeyID}")
    data = fetch_SQL(JobKeyID, "dbo.Heading2")
    if data:
        update_Access(data, r"E:\BM\FuturaFrames\Orders\2024\Jun\Test\RB262.mdb")
else:
    print("No valid JobKeyID retrieved.")

close_SQL()
