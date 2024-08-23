import os
import shutil
import json
from datetime import datetime

import headingAccess
import heading2Access
import headingEvoAccess


attachments_dir = "C:\\temp\\attachments"
base_dir = "E:\\BM\\FuturaFrames\\Orders"
log_file_path = os.path.join(attachments_dir, "job_log.json")
error_log_file_path = os.path.join(attachments_dir, "error_log.json")
success_log_file_path = os.path.join(attachments_dir, "success_log.json")


def get_destination_path(job_date):
    """Returns the path to the folder based on the job's date."""
    try:
        job_date = datetime.strptime(job_date, '%Y-%m-%d')
        year = job_date.strftime('%Y')
        month = job_date.strftime('%b')
        return os.path.join(base_dir, year, month)
    except ValueError as e:
        print(f"Error parsing date: {job_date} - {e}")
        return None


def transfer_job(job_name, job_info):
    """Transfers the job files to the appropriate folder."""
    try:
        destination_path = get_destination_path(job_info['Date'])
        
        if not destination_path:
            raise ValueError("Invalid destination path")

        # Ensure the destination directory exists
        if not os.path.exists(destination_path):
            os.makedirs(destination_path)

        # Paths to source files
        mdb_file = os.path.join(attachments_dir, f"{job_name}.mdb")
        svg_folder = os.path.join(attachments_dir, job_name)

        # Transfer the MDB file
        if job_info['mdb'] and os.path.exists(mdb_file):
            shutil.move(mdb_file, os.path.join(destination_path, f"{job_name}.mdb"))

        # Transfer the SVG folder
        if job_info['svg_folder'] and os.path.exists(svg_folder):
            shutil.move(svg_folder, os.path.join(destination_path, job_name))

        # Set the transfer flag to True if successful
        return True

    except Exception as e:
        print(f"Failed to transfer job {job_name}: {e}")
        return False


def cleanup_attachments_dir(job_log):
    """Remove all successfully transferred jobs from the attachments directory."""
    for job_name, job_info in job_log.items():
        if job_info['transfer']:
            # Remove MDB file if it exists
            mdb_file = os.path.join(attachments_dir, f"{job_name}.mdb")
            if os.path.exists(mdb_file):
                os.remove(mdb_file)

            # Remove SVG folder if it exists
            svg_folder = os.path.join(attachments_dir, job_name)
            if os.path.exists(svg_folder):
                shutil.rmtree(svg_folder)


def run_jobs(satellite_job_number):
    """Runs the three main jobs."""
    try:
        headingAccess.main(satellite_job_number)
    except Exception as e:
        print(f"Error running headingAccess for {satellite_job_number}: {e}")

    try:
        heading2Access.main(satellite_job_number)
    except Exception as e:
        print(f"Error running heading2Access for {satellite_job_number}: {e}")

    try:
        headingEvoAccess.main(satellite_job_number)
    except Exception as e:
        print(f"Error running headingEvoAccess for {satellite_job_number}: {e}")


def main():
    with open(log_file_path, 'r') as log_file:
        job_log = json.load(log_file)

    failed_jobs = {}
    successful_jobs = {}

    for job_name, job_info in job_log.items():
        if not job_info['transfer']:  # Skip already transferred jobs
            print(f"Processing job: {job_name}")
            run_jobs(job_name)
            successful = transfer_job(job_name, job_info)
            if successful:
                job_log[job_name]['transfer'] = True
                successful_jobs[job_name] = job_info
                print(f"Job {job_name} processed and transferred successfully.")
            else:
                failed_jobs[job_name] = job_info
                print(f"Job {job_name} encountered an error during transfer.")

    # Update the job log
    with open(log_file_path, 'w') as log_file:
        json.dump(job_log, log_file, indent=4)

    # Log any jobs that failed to transfer
    if failed_jobs:
        with open(error_log_file_path, 'w') as error_log_file:
            json.dump(failed_jobs, error_log_file, indent=4)
        print(f"The following jobs failed to transfer: {', '.join(failed_jobs.keys())}")

    # Log all successfully transferred jobs
    if successful_jobs:
        with open(success_log_file_path, 'w') as success_log_file:
            json.dump(successful_jobs, success_log_file, indent=4)
        print(f"The following jobs were successfully transferred: {', '.join(successful_jobs.keys())}")

    # Cleanup attachments directory
    cleanup_attachments_dir(job_log)

if __name__ == "__main__":
    main()
