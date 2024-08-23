import headingSqlToAccess
import heading2SqlToAccess
import headingEvoSqlToAccess

def main():
    satellite_job_number = 'RB431R'

    # Call the main function from each module
    headingSqlToAccess.main(satellite_job_number)
    heading2SqlToAccess.main(satellite_job_number)
    headingEvoSqlToAccess.main(satellite_job_number)

if __name__ == "__main__":
    main()
