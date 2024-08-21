import os
import zipfile
import shutil
import json
import pyodbc  # Use pyodbc for MDB access

# Folder where attachments are saved
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

def process_extracted_files(temp_extract_dir, job_log):
    """Process files extracted from a ZIP file."""
    for extracted_item in os.listdir(temp_extract_dir):
        extracted_item_path = os.path.join(temp_extract_dir, extracted_item)

        if extracted_item.endswith('.mdb'):
            # Handle existing MDB file case
            dest_path = os.path.join(attachments_dir, extracted_item)
            if os.path.exists(dest_path):
                # Rename the new MDB file to avoid overwriting
                base, ext = os.path.splitext(extracted_item)
                new_name = f"{base}_new{ext}"
                dest_path = os.path.join(attachments_dir, new_name)
            shutil.move(extracted_item_path, dest_path)

            job_name = os.path.splitext(extracted_item)[0]
            date_created = get_date_created_from_mdb(dest_path)
            if job_name not in job_log:
                job_log[job_name] = {'mdb': True, 'svg_folder': False, 'Date': date_created}
            else:
                job_log[job_name]['mdb'] = True
                job_log[job_name]['Date'] = date_created

        elif extracted_item.endswith('.svg'):
            job_name = extracted_item.split('-')[0]
            svg_folder_path = os.path.join(attachments_dir, job_name)
            os.makedirs(svg_folder_path, exist_ok=True)
            shutil.move(extracted_item_path, svg_folder_path)
            if job_name in job_log:
                job_log[job_name]['svg_folder'] = True
            else:
                job_log[job_name] = {'mdb': False, 'svg_folder': True}

        else:
            # Move anything else to Unsorted
            shutil.move(extracted_item_path, os.path.join(unsorted_dir, extracted_item))
            job_name = os.path.splitext(extracted_item)[0] if '.' in extracted_item else extracted_item
            if job_name not in job_log:
                job_log[job_name] = {'mdb': False, 'svg_folder': False, 'unsorted': True}

    # Remove the temporary extraction directory
    os.rmdir(temp_extract_dir)

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
            process_extracted_files(temp_extract_dir, job_log)

        elif item.endswith('.mdb'):
            job_name = os.path.splitext(item)[0]
            date_created = get_date_created_from_mdb(item_path)
            if job_name not in job_log:
                job_log[job_name] = {'mdb': True, 'svg_folder': False, 'Date': date_created}
            else:
                job_log[job_name]['mdb'] = True
                job_log[job_name]['Date'] = date_created

        elif item.endswith('.svg'):
            job_name = item.split('-')[0]
            svg_folder_path = os.path.join(attachments_dir, job_name)
            os.makedirs(svg_folder_path, exist_ok=True)
            shutil.move(item_path, svg_folder_path)
            if job_name in job_log:
                job_log[job_name]['svg_folder'] = True
            else:
                job_log[job_name] = {'mdb': False, 'svg_folder': True}

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

    # Second pass: Check the Unsorted folder and process any missed files
    for item in os.listdir(unsorted_dir):
        item_path = os.path.join(unsorted_dir, item)

        # Identify and process ZIP files that may still contain useful content
        if item.endswith('.zip'):
            temp_extract_dir = os.path.join(unsorted_dir, f"_extract_{os.path.splitext(item)[0]}")
            os.makedirs(temp_extract_dir, exist_ok=True)
            extract_zip(item_path, temp_extract_dir)
            os.remove(item_path)  # Remove the ZIP file after extraction

            # Process the extracted files
            process_extracted_files(temp_extract_dir, job_log)

        elif item.endswith('.mdb'):
            job_name = os.path.splitext(item)[0]
            date_created = get_date_created_from_mdb(item_path)
            dest_path = os.path.join(attachments_dir, item)
            if os.path.exists(dest_path):
                # Rename the new MDB file to avoid overwriting
                base, ext = os.path.splitext(item)
                new_name = f"{base}_new{ext}"
                dest_path = os.path.join(attachments_dir, new_name)
            shutil.move(item_path, dest_path)
            if job_name not in job_log:
                job_log[job_name] = {'mdb': True, 'svg_folder': False, 'Date': date_created}
            else:
                job_log[job_name]['mdb'] = True
                job_log[job_name]['Date'] = date_created

        elif item.endswith('.svg'):
            job_name = item.split('-')[0]
            svg_folder_path = os.path.join(attachments_dir, job_name)
            os.makedirs(svg_folder_path, exist_ok=True)
            shutil.move(item_path, svg_folder_path)
            if job_name in job_log:
                job_log[job_name]['svg_folder'] = True
            else:
                job_log[job_name] = {'mdb': False, 'svg_folder': True}

        else:
            # Keep anything that doesn't match the criteria in Unsorted
            job_name = os.path.splitext(item)[0] if '.' in item else item
            if job_name not in job_log:
                job_log[job_name] = {'mdb': False, 'svg_folder': False, 'unsorted': True}

    # Save the log to a JSON file
    with open(log_file_path, 'w') as log_file:
        json.dump(job_log, log_file, indent=4)

if __name__ == "__main__":
    cleanup_folder(attachments_dir, unsorted_dir, log_file_path)
    print(f"Folder cleanup complete. Unsorted files have been moved to the 'Unsorted' folder, and job log has been saved to {log_file_path}.")
