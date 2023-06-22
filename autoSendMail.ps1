$ErrorActionPreference = "Stop"

Write-Host ""
$AccountName = "acc name to automate the process"
$Password = Get-Content ".\mfaGReaderCert.txt" | ConvertTo-SecureString
$Credential = New-Object System.Management.Automation.PSCredential($AccountName, $Password)
Connect-MsolService -Credential $Credential

# Wrap in function cause redundancy
function Listing-Departments() {
    # Get all accounts excluding guest
    $Users = Get-MsolUser -All | Where-Object { $_.UserType -ne "Guest" }
    $Report = [System.Collections.Generic.List[Object]]::new()
    Write-Host "Read" $Users.Count "accounts..."

    # Loop through each users to create a custom data line for csv extract
    foreach ($User in $Users) {
        $ReportLine = [PSCustomObject] @{
            User        = $User.UserPrincipalName
            DisplayName = $user.DisplayName
            Department  = $User.Department
        }
        $Report.Add($ReportLine)
    }
    # $Report | Select User, Department | Sort User | Out-GridView
    $Report | Sort-Object Name | Export-CSV -NoTypeInformation -Encoding UTF8 # Path\File.csv
}

# Import pfx certificate
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
$secureString = ConvertTo-SecureString "Password" -AsPlainText -Force
$cert = Import-PfxCertificate -FilePath "Path\File.pfx" -CertStoreLocation 'Cert:\CurrentUser\My' -Password $secureString

# Credential hash table
$appRegistration = @{
    TenantId          = ""
    ClientId          = ""
    ClientCertificate = $cert
}

# Get token
$msalToken = Get-msaltoken @appRegistration -ForceRefresh

$studentList = Import-Csv -Path "Path\File.csv"

# Email content for each mail
foreach ($userName in $studentList){
    $requestBody = @{
        "message"         = [PSCustomObject]@{
            "subject"      = "Subject Email"
            "body"         = [PSCustomObject]@{
                "contentType" = "Text"
                "content"     = 

# No indentation to avoid indents in email

"Greetings Person,

Here's the bulk of content

Warm regards,

Name"

# End of content area

            }
            "toRecipients" = @(
                # Hash table for recipients
                [PSCustomObject]@{
                    "emailAddress" = [PSCustomObject]@{
                        "address" = $userName.User
                    }
                }
            )

            # Code for attachments
            "attachments"  = @(
                @{
                    "@odata.type"  = "#microsoft.graph.fileAttachment"
                    "name"         = "TextFile"
                    "contentType"  = "text/plain"
                    "contentBytes" = $attachementBytes
                }
            )
        }
        "saveToSentItems" = "true"
    }

    # Make the graph request
    $request = @{
        "Headers"     = @{Authorization = $msalToken.CreateAuthorizationHeader() }
        "Method"      = "Post"
        "Uri"         = "Link"
        "Body"        = $requestBody | ConvertTo-Json -Depth 5
        "ContentType" = "application/json"
    }

    # Execute
    Invoke-RestMethod @request
}