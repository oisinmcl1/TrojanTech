import os
import zipfile
import shutil

# Folder where attachments are saved
attachments_dir = "C:\\temp\\attachments"
unsorted_dir = os.path.join(attachments_dir, "Unsorted")

def cleanup_folder(attachments_dir, unsorted_dir):
    # Create the unsorted directory if it doesn't exist
    if not os.path.exists(unsorted_dir):
        os.makedirs(unsorted_dir)

    # Create a dictionary to store job names and their corresponding files
    jobs = {}

    # Traverse the directory
    for item in os.listdir(attachments_dir):
        item_path = os.path.join(attachments_dir, item)

        # Skip the Unsorted directory itself
        if item == 'Unsorted':
            continue

        # Identify MDB files and categorize them by job name
        if item.endswith('.mdb'):
            job_name = os.path.splitext(item)[0]
            jobs[job_name] = {'mdb': item_path, 'svg_dir': None, 'zip_files': []}

        # Identify ZIP files
        elif item.endswith('.zip'):
            with zipfile.ZipFile(item_path, 'r') as zip_ref:
                # Check contents of the ZIP file
                for zip_item in zip_ref.namelist():
                    if zip_item.endswith('.mdb'):
                        job_name = os.path.splitext(zip_item)[0]
                        if job_name not in jobs:
                            jobs[job_name] = {'mdb': None, 'svg_dir': None, 'zip_files': []}
                        jobs[job_name]['zip_files'].append(item_path)
                    elif zip_item.endswith('.svg'):
                        job_name = zip_item.split('/')[-1].split('.')[0].replace('SVG', '')
                        if job_name not in jobs:
                            jobs[job_name] = {'mdb': None, 'svg_dir': None, 'zip_files': []}
                        jobs[job_name]['zip_files'].append(item_path)

        # Identify folders that are likely SVG folders
        elif os.path.isdir(item_path) and item.startswith('SVG'):
            job_name = item.replace('SVG', '')
            if job_name in jobs:
                jobs[job_name]['svg_dir'] = item_path

        # If item doesn't fit the expected format, move it to the unsorted folder
        else:
            shutil.move(item_path, os.path.join(unsorted_dir, item))

    # Now, clean up the directory based on the identified jobs
    for job_name, job_info in jobs.items():
        # If an MDB file is found, delete any ZIP files containing that MDB file
        if job_info['mdb']:
            for zip_file in job_info['zip_files']:
                if job_info['mdb'] in zipfile.ZipFile(zip_file, 'r').namelist():
                    os.remove(zip_file)

        # If SVG files have been extracted, delete the original ZIP files
        if job_info['svg_dir']:
            for zip_file in job_info['zip_files']:
                with zipfile.ZipFile(zip_file, 'r') as zip_ref:
                    if any(name.endswith('.svg') for name in zip_ref.namelist()):
                        os.remove(zip_file)

            # Rename the SVG folder to match the job name
            new_svg_dir = os.path.join(attachments_dir, job_name)
            if job_info['svg_dir'] != new_svg_dir:
                shutil.move(job_info['svg_dir'], new_svg_dir)
                job_info['svg_dir'] = new_svg_dir

            # Move any SVG files in the main folder to the corresponding job folder
            for item in os.listdir(attachments_dir):
                item_path = os.path.join(attachments_dir, item)
                if item.endswith('.svg') and item.startswith(job_name):
                    shutil.move(item_path, job_info['svg_dir'])

        # If no SVG folder exists but there are extracted SVG files in the main folder, create the job folder
        elif not job_info['svg_dir']:
            new_svg_dir = os.path.join(attachments_dir, job_name)
            os.makedirs(new_svg_dir, exist_ok=True)

            # Move any SVG files in the main folder to the corresponding job folder
            for item in os.listdir(attachments_dir):
                item_path = os.path.join(attachments_dir, item)
                if item.endswith('.svg') and item.startswith(job_name):
                    shutil.move(item_path, new_svg_dir)

    # Move any remaining files that were not processed into the unsorted folder
    for item in os.listdir(attachments_dir):
        item_path = os.path.join(attachments_dir, item)
        if not os.path.isdir(item_path) and not item.endswith('.mdb') and item != 'Unsorted':
            shutil.move(item_path, os.path.join(unsorted_dir, item))
        elif os.path.isdir(item_path) and item not in jobs and item != 'Unsorted':
            shutil.move(item_path, os.path.join(unsorted_dir, item))

if __name__ == "__main__":
    cleanup_folder(attachments_dir, unsorted_dir)
    print("Folder cleanup complete. Unsorted files have been moved to the 'Unsorted' folder.")
