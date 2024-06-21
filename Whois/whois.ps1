function Get-DomainExpiryInfo {
    param(
        [parameter(Mandatory = $true)][string]$Domain,
        [parameter(Mandatory = $false)][int]$DaysThreshold = 30
    )

    try {
        # Fetch the HTML content from who.is website
        $response = Invoke-WebRequest -Uri "https://who.is/whois/$Domain" -TimeoutSec 15 -ErrorAction Stop

        # Output the status code and headers for debugging
        # Write-Output ("Status Code: " + $response.StatusCode)
        # Write-Output ("Headers: " + $response.Headers)

        # Output the HTML to a file for inspection
        # $response.Content | Out-File -FilePath "whoisWebResult.html" -Force

        # Parse the HTML content
        $parsedHtml = $response.ParsedHtml
        $dateElements = $parsedHtml.getElementsByTagName('div') | Where-Object { $_.className -eq 'col-md-8 queryResponseBodyValue' }

        # Debug: Output all matched elements
        # Write-Output "Matched Elements:"
        # $dateElements | ForEach-Object { Write-Output $_.innerText }

        # Extract all date elements
        $dates = @()
        foreach ($element in $dateElements) {
            if ($element.innerText -match '^\d{4}-\d{2}-\d{2}$') {
                $dates += [datetime]::ParseExact($element.innerText.Trim(), 'yyyy-MM-dd', $null)
            }
        }

        if ($dates.Count -eq 0) {
            throw "No valid date elements found in WHOIS data"
        }

        # Assume the latest date is the expiry date
        $expiryDate = $dates | Sort-Object -Descending | Select-Object -First 1

        # Debug: Output the detected expiry date
        # Write-Output ("Detected Expiry Date: " + $expiryDate)

        $currentDate = Get-Date
        $daysUntilExpiry = ($expiryDate - $currentDate).Days
        $expiresSoon = $daysUntilExpiry -le $DaysThreshold

        # Format the expiry date for simplified output
        $expiryDateString = $expiryDate.ToString('yyyy-MM-dd')

        # Write results to text files
        $expiryDateString | Out-File -FilePath "expiryDate.txt" -Force
        $daysUntilExpiry | Out-File -FilePath "daysUntilExpiry.txt" -Force
        $expiresSoon | Out-File -FilePath "expiresSoon.txt" -Force

        # Return results
        return [PSCustomObject]@{
            ExpiryDate = $expiryDateString
            DaysUntilExpiry = $daysUntilExpiry
            ExpiresSoon = $expiresSoon
        }
    }
    catch {
        Write-Warning ("Error getting WHOIS details for $Domain. $_")
    }
}
