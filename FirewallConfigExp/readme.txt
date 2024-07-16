Overview

This PowerShell script is designed to retrieve the configuration of Palo Alto Firewalls using their API. 
It assumes the default gateway IP as the firewall IP for connection.


Get-APIKey
Uses the provided username and password to authenticate and retrieve an API key from the firewall.
Defaults to 'admin' as the username.
Encodes the password securely before sending it to the firewall's API endpoint.


Export-FirewallConfig
Once authenticated, exports the firewall configuration in XML format.
Saves the configuration file to a specified directory.


Password Handling

Security: The script ensures secure handling of passwords by using PowerShell's SecureString type for sensitive information.
Input: The password is expected as a base64-encoded string (-encodedPassword parameter) to protect it during input and transmission.
Decoding: Internally, the script decodes the base64-encoded password and converts it into a SecureString before interacting with the Palo Alto Networks firewall API.


Usage

Parameters

-username: Username for firewall authentication (default: admin).
-encodedPassword: Base64-encoded password for firewall authentication.
-outputDir: Directory path where the exported configuration file will be saved.

.\getConfig.ps1 -username "admin" -encodedPassword "base64EncodedPassword" -outputDir "C:\temp"


Addtional Resources

https://docs.paloaltonetworks.com/develop/api#sort=relevancy&layout=card&numberOfResults=25
https://docs.paloaltonetworks.com/pan-os/11-1/pan-os-panorama-api

Oisín Mc Laughlin
16/07/2024