import os
import json
import pyodbc  # Use pyodbc for MDB access
from datetime import datetime

# Directory where attachments are saved
attachments_dir = "C:\\temp\\attachments"
log_file_path = os.path.join(attachments_dir, "job_log.json")

def get_job_info_from_mdb(mdb_file):
    """Extract the DateCreated and JobKeyID fields from the Heading table in the MDB file."""
    try:
        conn_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={mdb_file};'
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()

        cursor.execute("SELECT TOP 1 DateCreated, JobKeyID FROM Heading")
        row = cursor.fetchone()
        if row:
            # Convert DateCreated to just the date part
            date_created = row[0].date() if isinstance(row[0], datetime) else None
            job_key_id = str(row[1])
            return date_created, job_key_id
        else:
            return None, None
    except Exception as e:
        print(f"Error reading MDB file {mdb_file}: {e}")
        return None, None
    finally:
        try:
            cursor.close()
            conn.close()
        except:
            pass

def rebuild_log(attachments_dir, log_file_path):
    """Rebuild the job log by scanning the attachments directory."""
    job_log = {}

    # Scan through the attachments directory
    for item in os.listdir(attachments_dir):
        item_path = os.path.join(attachments_dir, item)

        # If it's a folder, treat it as a job folder for SVGs
        if os.path.isdir(item_path):
            job_name = item
            mdb_file = os.path.join(attachments_dir, job_name + ".mdb")
            
            # Check for an MDB file with the same name as the folder
            if os.path.exists(mdb_file):
                date_created, job_key_id = get_job_info_from_mdb(mdb_file)
                job_log[job_name] = {
                    'mdb': True, 
                    'svg_folder': True, 
                    'Date': date_created.strftime('%Y-%m-%d') if date_created else None, 
                    'JobKeyID': job_key_id,
                    'transfer': False  # Set transfer to False by default
                }
            else:
                job_log[job_name] = {
                    'mdb': False, 
                    'svg_folder': True, 
                    'transfer': False  # Set transfer to False by default
                }

        # If it's an MDB file, ensure it's logged even if there’s no corresponding folder
        elif item.endswith('.mdb'):
            job_name = os.path.splitext(item)[0]
            date_created, job_key_id = get_job_info_from_mdb(item_path)

            if job_name not in job_log:
                job_log[job_name] = {
                    'mdb': True, 
                    'svg_folder': False, 
                    'Date': date_created.strftime('%Y-%m-%d') if date_created else None, 
                    'JobKeyID': job_key_id,
                    'transfer': False  # Set transfer to False by default
                }
            else:
                job_log[job_name]['mdb'] = True
                job_log[job_name]['Date'] = date_created.strftime('%Y-%m-%d') if date_created else None
                job_log[job_name]['JobKeyID'] = job_key_id
                job_log[job_name]['transfer'] = False  # Ensure transfer is set to False

    # Save the log to a JSON file
    with open(log_file_path, 'w') as log_file:
        json.dump(job_log, log_file, indent=4)

    print(f"Log file successfully rebuilt at {log_file_path}")

if __name__ == "__main__":
    rebuild_log(attachments_dir, log_file_path)
