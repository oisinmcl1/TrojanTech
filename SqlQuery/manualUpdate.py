import headingAccess
import heading2Access
import headingEvoAccess

def main():
    satellite_job_number = 'RB515'

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

if __name__ == "__main__":
    main()
