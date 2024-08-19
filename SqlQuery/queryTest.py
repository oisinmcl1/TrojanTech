import pyodbc

def fetch_contact_name():
    dsn = "Orders"
    database = "FFOrders"
    satellite_job_number = 'RB262'

    try:
        # Establish connection to SQL Server
        sqlServer = pyodbc.connect(f"DSN={dsn};DATABASE={database};")
        sqlCursor = sqlServer.cursor()

        print(f"Connection to SQL Server: {dsn} and Database: {database} was successful!")

        # SQL query to fetch the Contact name based on SatelliteJobNumber
        fetchQuery_Contact = """
        SELECT Contact
        FROM dbo.Heading
        WHERE SatelliteJobNumber = ?;
        """

        # Execute the query with the satellite job number parameter
        sqlCursor.execute(fetchQuery_Contact, satellite_job_number)
        contact_name = sqlCursor.fetchone()

        if contact_name:
            print(f"Fetched Contact name for SatelliteJobNumber '{satellite_job_number}': {contact_name[0]}")
        else:
            print(f"No data found for SatelliteJobNumber = '{satellite_job_number}'.")

    except pyodbc.Error as ex:
        print(f"SQL Server connection or query failed: {ex}")

    finally:
        sqlCursor.close()
        sqlServer.close()
        print("SQL Connection Closed.")

# Run the script
fetch_contact_name()
