# Create object of security info of logged in user.
$id = [System.Security.Prinicple.WindowsIdentity]::GetCurrent()
$p = New-Object System.Security.Prinicple.WindowsPrinciple($id)