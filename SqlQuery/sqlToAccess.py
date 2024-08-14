import pyodbc

dsn = "Orders"
database = "FFOrders"

# Try connect to database and create a cursor
try:
    sqlServer = pyodbc.connect(f"DSN={dsn};DATABASE={database};")

    sqlCursor = sqlServer.cursor()

    print(f"Connection to SQL Server: {dsn} and Database: {database} was successful!")

# If error from pyodbc, print the SQLSTATE code
except pyodbc.Error as ex:
    sqlstate = ex.args[1]
    print(f"An error occurred while trying to connect: {sqlstate}")

# Try close connection if it was established
finally:
    try:
        sqlCursor.close()
        sqlServer.close()
        print("Connection closed.")
    except:
        pass