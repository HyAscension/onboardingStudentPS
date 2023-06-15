$ErrorActionPreference = "Stop"

Write-Host ""
$AccountName = "acc name to automate the process"
$Password = Get-Content ".\mfaGReaderCert.txt" | ConvertTo-SecureString
$Credential = New-Object System.Management.Automation.PSCredential($AccountName, $Password)
Connect-MsolService -Credential $Credential