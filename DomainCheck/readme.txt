Domain Expiration Date Checker

Description:
This is a PowerShell script that checks the expiration date of a domain. It works by querying the appropriate WHOIS server for the domain's Top-Level Domain (TLD), extracting the expiration date from the WHOIS data, and calculating the number of days until the domain expires. If the number of days until expiration is less than or equal to a specified threshold, it returns true; otherwise, it returns false. The script also writes the expiration date, the number of days until expiration, and whether the domain is within the threshold to separate text files.

Prerequisites:
- PowerShell

Instructions:

1. Save the provided script as 'domaincheck.ps1' in your desired directory.

2. Open PowerShell and navigate to the directory where you saved the script.

3. Source the script using the following command:
   PS C:\Users\YourUsername\YourDirectory> . .\domaincheck.ps1

4. Now, you can run the 'Get-DomainExpirationDate' function with the '-Domain' parameter followed by the domain name you want to check. For example:
   PS C:\Users\YourUsername\YourDirectory> Get-DomainExpirationDate -Domain "rte.ie"

This will return the expiration date of the domain and create three text files in the current directory: 'expiration_date.txt', 'days_until_expiration.txt', and 'is_within_threshold.txt', which contain the expiration date, the number of days until expiration, and whether the domain is within the threshold, respectively.


Oisín Mc Laughlin
19/06/2024