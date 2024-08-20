import os
import zipfile
import shutil
import json

# Folder where attachments are saved
attachments_dir = "C:\\temp\\attachments"
unsorted_dir = os.path.join(attachments_dir, "Unsorted")
log_file_path = os.path.join(attachments_dir, "job_log.json")

def extract_svg_zip(zip_path, target_dir):
    """Extract SVG files from a ZIP file into the specified directory."""
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(target_dir)

def cleanup_folder(attachments_dir, unsorted_dir, log_file_path):
    # Create the unsorted directory if it doesn't exist
    if not os.path.exists(unsorted_dir):
        os.makedirs(unsorted_dir)

    # Initialize log structure
    job_log = {}

    # Traverse the directory
    for item in os.listdir(attachments_dir):
        item_path = os.path.join(attachments_dir, item)

        # Skip the Unsorted directory itself
        if item == 'Unsorted' or item == os.path.basename(log_file_path):
            continue

        # Identify MDB files and categorize them by job name
        if item.endswith('.mdb'):
            job_name = os.path.splitext(item)[0]
            if job_name not in job_log:
                job_log[job_name] = {'mdb': True, 'svg_folder': False}
            else:
                job_log[job_name]['mdb'] = True

        # Identify ZIP files containing SVGs and process them
        elif item.endswith('.zip') and item.startswith('SVG'):
            job_name = item.replace('SVG', '').split('.')[0]
            svg_folder_path = os.path.join(attachments_dir, job_name)

            # Extract SVGs into a new folder named after the job
            if not os.path.exists(svg_folder_path):
                os.makedirs(svg_folder_path)

            extract_svg_zip(item_path, svg_folder_path)
            os.remove(item_path)  # Remove the ZIP file after extraction

            # Update log
            if job_name not in job_log:
                job_log[job_name] = {'mdb': False, 'svg_folder': True}
            else:
                job_log[job_name]['svg_folder'] = True

        # Identify folders that are already created for jobs (possibly SVG folders)
        elif os.path.isdir(item_path) and not item.startswith('SVG'):
            job_name = item
            if job_name in job_log:
                job_log[job_name]['svg_folder'] = True
            else:
                job_log[job_name] = {'mdb': False, 'svg_folder': True}

        # If item doesn't fit the expected format, move it to the unsorted folder
        else:
            shutil.move(item_path, os.path.join(unsorted_dir, item))

    # Final sweep for any remaining SVG files in the main folder and move them to their respective folders
    for item in os.listdir(attachments_dir):
        item_path = os.path.join(attachments_dir, item)
        if item.endswith('.svg'):
            job_name = item.split('.')[0]
            svg_folder_path = os.path.join(attachments_dir, job_name)

            if not os.path.exists(svg_folder_path):
                os.makedirs(svg_folder_path)

            shutil.move(item_path, svg_folder_path)

            if job_name in job_log:
                job_log[job_name]['svg_folder'] = True
            else:
                job_log[job_name] = {'mdb': False, 'svg_folder': True}

    # Move any remaining files that were not processed into the unsorted folder
    for item in os.listdir(attachments_dir):
        item_path = os.path.join(attachments_dir, item)
        if not os.path.isdir(item_path) and not item.endswith('.mdb') and item != 'Unsorted':
            shutil.move(item_path, os.path.join(unsorted_dir, item))
        elif os.path.isdir(item_path) and item not in job_log and item != 'Unsorted':
            shutil.move(item_path, os.path.join(unsorted_dir, item))

    # Save the log to a JSON file
    with open(log_file_path, 'w') as log_file:
        json.dump(job_log, log_file, indent=4)

if __name__ == "__main__":
    cleanup_folder(attachments_dir, unsorted_dir, log_file_path)
    print(f"Folder cleanup complete. Unsorted files have been moved to the 'Unsorted' folder, and job log has been saved to {log_file_path}.")
