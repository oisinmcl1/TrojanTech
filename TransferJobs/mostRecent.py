import json
from datetime import datetime

# Path to the job log file
log_file_path = "C:\\temp\\job_log.json"

def find_most_recent_jobs(log_file_path, num_jobs=20):
    """Find the most recent jobs from the log file."""
    try:
        # Load the job log
        with open(log_file_path, 'r') as log_file:
            job_log = json.load(log_file)

        # Filter out jobs that have a valid date
        jobs_with_dates = [
            {"job_name": job_name, "job_info": job_info}
            for job_name, job_info in job_log.items()
            if job_info.get('Date')
        ]

        # Sort jobs by date in descending order (most recent first)
        sorted_jobs = sorted(
            jobs_with_dates,
            key=lambda x: datetime.strptime(x['job_info']['Date'], '%Y-%m-%d'),
            reverse=True
        )

        # Get the most recent `num_jobs` jobs
        most_recent_jobs = sorted_jobs[:num_jobs]

        # Print or return the most recent jobs
        for job in most_recent_jobs:
            job_name = job['job_name']
            job_date = job['job_info']['Date']
            job_transfer = job['job_info']['transfer']
            print(f"Job: {job_name}, Date: {job_date}, Transferred: {job_transfer}")

        return most_recent_jobs

    except Exception as e:
        print(f"An error occurred: {e}")
        return []

if __name__ == "__main__":
    # Find and print the 20 most recent jobs
    recent_jobs = find_most_recent_jobs(log_file_path, num_jobs=20)
