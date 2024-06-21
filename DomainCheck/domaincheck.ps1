function Get-DomainExpirationDate {
    Param(
        [Parameter(Mandatory=$true)]
        [string]$domain,

        [int]$threshold = 30
    )

    try {
        # Determine the appropriate WHOIS server
        $tld = $domain.Split('.')[-1]
        $whoisServer = switch ($tld) {
            "com" { "whois.verisign-grs.com" }
            "net" { "whois.verisign-grs.com" }
            "org" { "whois.pir.org" }
            "info" { "whois.afilias.net" }
            "co" { "whois.nic.co" }
            "us" { "whois.nic.us" }
            "ie" { "whois.iedr.ie" }
            "uk" { "whois.nic.uk" }
            "co.uk" { "whois.nic.uk" }
            default { "whois.iana.org" }
        }

        # Query the WHOIS server
        $tcpClient = New-Object System.Net.Sockets.TcpClient($whoisServer, 43)
        $stream = $tcpClient.GetStream()
        $buffer = [Text.Encoding]::ASCII.GetBytes("$domain`r`n")
        $stream.Write($buffer, 0, $buffer.Length)

        $reader = New-Object System.IO.StreamReader($stream)
        $whoisData = $reader.ReadToEnd()

        $tcpClient.Close()

        # Extract expiration date
        if ($whoisData -match 'Expiry Date:\s+(\S+)' -or $whoisData -match 'Registry Expiry Date:\s+(\S+)' -or $whoisData -match 'Expiration Date:\s+(\S+)' -or $whoisData -match 'Renewal Date:\s+(\S+)') {
            $expirationDate = [datetime]::Parse($matches[1])
            
            # Calculate days until expiration
            $daysUntilExpiration = ($expirationDate - (Get-Date)).Days

            # Check if within threshold
            $isWithinThreshold = $daysUntilExpiration -le $threshold

            # Write outputs to files
            $expirationDate.ToString("dd/MM/yyyy") | Out-File -FilePath ".\expiration_date.txt" -NoNewline
            $daysUntilExpiration | Out-File -FilePath ".\days_until_expiration.txt" -NoNewline
            $isWithinThreshold | Out-File -FilePath ".\is_within_threshold.txt" -NoNewline

            return $expirationDate.ToString("dd/MM/yyyy")
        } else {
            "Expiration date not found" | Out-File -FilePath ".\expiration_date.txt" -NoNewline
            "N/A" | Out-File -FilePath ".\days_until_expiration.txt" -NoNewline
            $false | Out-File -FilePath ".\is_within_threshold.txt" -NoNewline

            return "Expiration date not found"
        }
    } catch {
        $_.Exception.Message
    }
}
