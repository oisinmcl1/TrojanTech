import json
from datetime import datetime

# Path to the job log file
log_file_path = "C:\\temp\\job_log.json"

def find_july_jobs(log_file_path):
    """Find all jobs from July in the log file."""
    try:
        # Load the job log
        with open(log_file_path, 'r') as log_file:
            job_log = json.load(log_file)

        # Filter out jobs that have a valid date in July
        july_jobs = [
            {"job_name": job_name, "job_info": job_info}
            for job_name, job_info in job_log.items()
            if job_info.get('Date') and datetime.strptime(job_info['Date'], '%Y-%m-%d').month == 7
        ]

        # Print the July jobs
        for job in july_jobs:
            job_name = job['job_name']
            job_date = job['job_info']['Date']
            job_transfer = job['job_info']['transfer']
            print(f"Job: {job_name}, Date: {job_date}, Transferred: {job_transfer}")

        return july_jobs

    except Exception as e:
        print(f"An error occurred: {e}")
        return []

if __name__ == "__main__":
    # Find and print all July jobs
    july_jobs = find_july_jobs(log_file_path)
