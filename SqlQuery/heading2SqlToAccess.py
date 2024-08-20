import pyodbc

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
            data = dict(zip(columns, row))
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


connect_SQL()

JobKeyID = fetch_JobKeyID()

if JobKeyID is not None:
    print(f"JobKeyID: {JobKeyID}")
    fetch_SQL(JobKeyID, "dbo.Heading2")
else:
    print("No valid JobKeyID retrieved.")

close_SQL()
