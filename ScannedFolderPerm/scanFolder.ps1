$folder = Read-Host "User Folder:" 

$folderPath = "D:\Documents\$folder"
$userName = "Scan_To_Folder"

# Get current acl for folder
$acl = Get-Acl $folderPath
# Define the permissions to grant (excluding Full Control and Special Permissions)
$rights = [System.Security.AccessControl.FileSystemRights]::ReadAndExecute -bor 
          [System.Security.AccessControl.FileSystemRights]::ListDirectory -bor
          [System.Security.AccessControl.FileSystemRights]::Read -bor
          [System.Security.AccessControl.FileSystemRights]::Write

# Create a new access rule for the specified user
$rule = New-Object System.Security.AccessControl.FileSystemAccessRule($userName, $rights, "ContainerInherit, ObjectInherit", "None", "Allow")

## $rule = New-Object System.Security.AccessControl.FileSystemAccessRule($userName, "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")

# Add and apply rule
$acl.AddAccessRule($rule)
Set-Acl $folderPath $acl

# Share the folder with full control permissions for the Everyone group
$shareName = Split-Path -Leaf $folderPath
net share $shareName=$folderPath /grant:Everyone,Full

Write-Host "Folder $folderPath correct sharing permissions and $userName added"