import pyodbc
import json

# Database connection parameters
dsn = "Orders"
database = "FFOrders"

# Log file paths
input_log_file = "C:\\temp\\attachments\\job_log.json"
existing_jobs_log_file = "C:\\temp\\attachments\\existing_jobs_log.json"
non_existing_jobs_log_file = "C:\\temp\\attachments\\non_existing_jobs_log.json"

def fetch_existing_jobs():
    """Fetch the SatelliteJobNumber from the dbo.heading table."""
    try:
        # Establish connection
        print(f"Connecting to SQL Server: DSN={dsn}, Database={database}")
        sqlServer = pyodbc.connect(f"DSN={dsn};DATABASE={database};")
        sqlCursor = sqlServer.cursor()

        print("Connection successful! Fetching SatelliteJobNumber data...")
        # SQL query to fetch the SatelliteJobNumber
        sqlCursor.execute("SELECT SatelliteJobNumber FROM dbo.heading")
        result = sqlCursor.fetchall()

        # Convert the result into a set for faster lookup
        existing_jobs = {row[0] for row in result}

        print(f"Fetched {len(existing_jobs)} SatelliteJobNumbers from the database.")
        return existing_jobs

    except pyodbc.Error as ex:
        print(f"SQL Server connection or query failed: {ex}")
        return set()

    finally:
        # Ensure that the connection is closed properly
        try:
            sqlCursor.close()
            sqlServer.close()
            print("SQL Connection Closed.")
        except Exception as e:
            print(f"Error closing the connection: {e}")

def separate_jobs_by_existence(existing_jobs):
    """Separate jobs into those that exist in the SQL table and those that don't."""
    try:
        with open(input_log_file, 'r') as log_file:
            job_log = json.load(log_file)

        existing_jobs_log = {}
        non_existing_jobs_log = {}

        for job_name, job_info in job_log.items():
            if job_name in existing_jobs:
                existing_jobs_log[job_name] = job_info
            else:
                non_existing_jobs_log[job_name] = job_info

        # Write existing jobs log
        with open(existing_jobs_log_file, 'w') as log_file:
            json.dump(existing_jobs_log, log_file, indent=4)

        # Write non-existing jobs log
        with open(non_existing_jobs_log_file, 'w') as log_file:
            json.dump(non_existing_jobs_log, log_file, indent=4)

        print(f"Existing jobs log saved to {existing_jobs_log_file}")
        print(f"Non-existing jobs log saved to {non_existing_jobs_log_file}")

    except Exception as e:
        print(f"Error processing logs: {e}")

if __name__ == "__main__":
    # Fetch existing jobs from the SQL table
    existing_jobs = fetch_existing_jobs()

    # Separate and save jobs by existence
    if existing_jobs:
        separate_jobs_by_existence(existing_jobs)
    else:
        print("No existing jobs found in the database.")
