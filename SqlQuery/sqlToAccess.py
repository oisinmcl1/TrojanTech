import pyodbc


"""
CONNECT TO SQL SERVER
FETCH DBO_HEADING, DBO_HEADING2 AND DBO_HEADINGEVONET
"""

dsn = "Orders"
database = "FFOrders"

# Try connect to database and create a cursor
try:
    sqlServer = pyodbc.connect(f"DSN={dsn};DATABASE={database};")

    sqlCursor = sqlServer.cursor()

    print(f"Connection to SQL Server: {dsn} and Database: {database} was successful!")

    # Fetch Query from SQL Server
    fetchQuery = ""

    # Execute fetch query
    ## sqlCursor.execute(fetchQuery)
    ## rows = sqlCursor.fetchall()

# If error, print message and ensure rows is none
except pyodbc.Error as ex:
    print(f"SQL Server connection or query failed: {ex}")
    rows = None

# Close connection
finally:
    sqlCursor.close()
    sqlServer.close()
    print("SQL Connection Closed.")





"""
CONNECT TO ACCESS DATABASE
UPDATE HEADING, HEADING2 AND HEADINGEVONET
"""

mdb = r"E:\BM\FuturaFrames\Orders\2024\Jun\RB262.mdb"

try:
    # Connection string to Access database
    connStr = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={mdb};"
    )

    # Try to connect to Access database and create cursor
    accessConn = pyodbc.connect(connStr)

    accessCursor = accessConn.cursor()

    print(f"Connection to Access Database: {mdb} was successful!")

# If error, print error message
except pyodbc.Error as ex:
    print(f"Access Database connection failed: {ex}")

# Close connection
finally:
    try:
        accessCursor.close()
        accessConn.close()
        print("Access Connection Closed.")
    except:
        print("No connection to close or error closing connection")
