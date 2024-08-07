# SMTP Server address, mail from address and rcpt to address
param (
    [string]$smtpServer,
    [string]$fromEmail,
    [string]$toEmail
)

# If no function parameters, read through cli
if (-not $smtpServer) {
    $smtpServer = Read-Host "SMTP Server address"
}
if (-not $fromEmail) {
    $fromEmail = Read-Host "Mail from address"
}
if (-not $toEmail) {
    $toEmail = Read-Host "Rcpt to address"
}

# Create a TCP connection to SMTP server on port 25