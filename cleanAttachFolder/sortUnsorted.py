import os
import shutil
import json
import pyodbc  # Use pyodbc for MDB access

# Folders for attachments and unsorted files
attachments_dir = "C:\\temp\\attachments"
unsorted_dir = os.path.join(attachments_dir, "Unsorted")
log_file_path = os.path.join(attachments_dir, "job_log.json")

def extract_zip(zip_path, target_dir):
    """Extract all files from a ZIP file into the specified directory."""
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(target_dir)

def get_date_created_from_mdb(mdb_file):
    """Extract the DateCreated field from the Heading table in the MDB file."""
    try:
        conn_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={mdb_file};'
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()

        cursor.execute("SELECT TOP 1 DateCreated FROM Heading")
        row = cursor.fetchone()
        if row:
            return str(row[0])
        else:
            return None
    except Exception as e:
        print(f"Error reading MDB file {mdb_file}: {e}")
        return None
    finally:
        try:
            cursor.close()
            conn.close()
        except:
            pass

def move_svg_files_to_job_folders(directory, job_log):
    """Move SVG files into folders named after their respective jobs."""
    for item in os.listdir(directory):
        item_path = os.path.join(directory, item)

        if item.endswith('.svg'):
            # Extract job name (everything before the first hyphen)
            job_name = item.split('-')[0].strip()
            svg_folder_path = os.path.join(attachments_dir, job_name)

            # Ensure the folder exists
            os.makedirs(svg_folder_path, exist_ok=True)

            # Move the SVG file into the job folder
            shutil.move(item_path, svg_folder_path)
            print(f"Moved {item} to {svg_folder_path}")

            # Update the log
            if job_name in job_log:
                job_log[job_name]['svg_folder'] = True
            else:
                job_log[job_name] = {'mdb': False, 'svg_folder': True}

def cleanup_folder(attachments_dir, unsorted_dir, log_file_path):
    # Create the unsorted directory if it doesn't exist
    if not os.path.exists(unsorted_dir):
        os.makedirs(unsorted_dir)

    # Initialize log structure
    job_log = {}

    # First pass: Process ZIP files and extract contents
    for item in os.listdir(attachments_dir):
        item_path = os.path.join(attachments_dir, item)

        # Skip the Unsorted directory itself
        if item == 'Unsorted' or item == os.path.basename(log_file_path):
            continue

        # Identify and process ZIP files
        if item.endswith('.zip'):
            temp_extract_dir = os.path.join(attachments_dir, f"_extract_{os.path.splitext(item)[0]}")
            os.makedirs(temp_extract_dir, exist_ok=True)
            extract_zip(item_path, temp_extract_dir)
            os.remove(item_path)  # Remove the ZIP file after extraction

            # Process the extracted files
            move_svg_files_to_job_folders(temp_extract_dir, job_log)

        elif item.endswith('.mdb'):
            job_name = os.path.splitext(item)[0]
            date_created = get_date_created_from_mdb(item_path)
            if job_name not in job_log:
                job_log[job_name] = {'mdb': True, 'svg_folder': False, 'Date': date_created}
            else:
                job_log[job_name]['mdb'] = True
                job_log[job_name]['Date'] = date_created

        elif item.endswith('.svg'):
            move_svg_files_to_job_folders(attachments_dir, job_log)

        elif os.path.isdir(item_path) and not item.startswith('SVG'):
            job_name = item
            if job_name in job_log:
                job_log[job_name]['svg_folder'] = True
            else:
                job_log[job_name] = {'mdb': False, 'svg_folder': True}

        else:
            # Move anything else to Unsorted
            shutil.move(item_path, os.path.join(unsorted_dir, item))
            job_name = os.path.splitext(item)[0] if '.' in item else item
            if job_name not in job_log:
                job_log[job_name] = {'mdb': False, 'svg_folder': False, 'unsorted': True}

    # Process remaining SVG files in Unsorted folder
    move_svg_files_to_job_folders(unsorted_dir, job_log)

    # Double-check to remove any files left in Unsorted if successfully moved
    for item in os.listdir(unsorted_dir):
        item_path = os.path.join(unsorted_dir, item)
        job_name = item.split('-')[0].strip()
        svg_folder_path = os.path.join(attachments_dir, job_name)

        if os.path.exists(os.path.join(svg_folder_path, item)):
            os.remove(item_path)
            print(f"Removed {item} from Unsorted as it was successfully moved.")

    # Check if Unsorted is empty and remove it if so
    if not os.listdir(unsorted_dir):
        os.rmdir(unsorted_dir)

    # Save the log to a JSON file
    with open(log_file_path, 'w') as log_file:
        json.dump(job_log, log_file, indent=4)

if __name__ == "__main__":
    cleanup_folder(attachments_dir, unsorted_dir, log_file_path)
    print(f"Folder cleanup complete. SVG files have been moved to job folders, and job log has been saved to {log_file_path}.")
