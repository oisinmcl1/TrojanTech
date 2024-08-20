import pyodbc

sqlServer = None
sqlCursor = None

def connect_to_sql_server():
    global sqlServer, sqlCursor
    
    dsn = "Orders"
    database = "FFOrders"
    
    try:
        sqlServer = pyodbc.connect(f"DSN={dsn};DATABASE={database};")
        sqlCursor = sqlServer.cursor()
        print(f"Connection to SQL Server: {dsn} and Database: {database} was successful!")
    
    except pyodbc.Error as ex:
        print(f"SQL Server connection failed: {ex}")

def close_sql_server_connection():
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

connect_to_sql_server()

JobKeyID = fetch_JobKeyID()

if JobKeyID is not None:
    print(f"JobKeyID: {JobKeyID}")
else:
    print("No valid JobKeyID retrieved.")

close_sql_server_connection()
