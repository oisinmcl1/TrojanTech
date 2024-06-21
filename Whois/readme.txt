Get-DomainExpiryInfo

This PowerShell script fetches the expiry information of a domain. It was inspired by a post on PowerShellIsFun by HARM VEENSTRA.
https://powershellisfun.com/2022/06/12/get-whois-information-using-powershell/

Functionality

The Get-DomainExpiryInfo function takes a domain name as input and optionally a days threshold (default is 30). It fetches the WHOIS information for the domain from the who.is website and parses the HTML content to extract the expiry date of the domain. It then calculates the number of days until expiry and checks if the domain is expiring soon (based on the days threshold).

The function returns an object with the following properties:
- ExpiryDate: The expiry date of the domain.
- DaysUntilExpiry: The number of days until the domain expires.
- ExpiresSoon: A boolean indicating whether the domain is expiring soon (based on the days threshold).

Usage

First, navigate to the directory containing the script and source it:

PS C:\Users\YourUsername\Path\To\Script> . .\whois.ps1

Then, you can call the function with the domain you want to check:

PS C:\Users\YourUsername\Path\To\Script> Get-DomainExpiryInfo -Domain "example.com"

This will output something like:

ExpiryDate      DaysUntilExpiry ExpiresSoon
----------      --------------- -----------
2025-03-31      282             False

Error Handling

If there's an error while fetching the WHOIS details, the function will output a warning message with the error details.

Note

This script is for educational purposes and should be used responsibly. Please respect the terms of service of the who.is website.


Oisín Mc Laughlin
20/06/2024