# To run, you will need to set a couple environment variables
# $env:username = EXO username
# $env:passwd = EXO password
# $env:AzureWebJobsStorage = This is the connection string for the Azure Storage Account where the emails will be stored and where the Azure Storage Table will be created
# $env:ContainerName = This is the Azure Blob container where the .eml files will be stored


# Connect to Exchange Online
# Set the credentials for connecting to Exchange Online
$Password = ConvertTo-SecureString $env:passwd -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential($env:username, $Password)

# Connect to Exchange Online
Connect-ExchangeOnline -Credential $Credential -UseRPSSession

# Get the Azure Storage Account Context
$ctx = New-AzStorageContext -ConnectionString $env:AzureWebJobsStorage
$tableName = "PhishingEmails"
$tableRow = @{}

# Get the Azure Storage Table
$cloudTable = (Get-AzStorageTable -Name $tableName -Context $ctx).CloudTable

# Create a temporary folder to store the exported emails
$TempFolder = "$env:Temp\PhishingEmails"
New-Item -ItemType Directory -Path $TempFolder | Out-Null

# Search for the phishing emails in the quarantine and export them to Azure Storage
$emailTypes = @("HighConfPhish","Phish")

# Get some stats
$TotalEmails = 0
$TotalLength = 0

foreach($emailType in $emailTypes)
{
    $SearchResults = Get-QuarantineMessage -Type $emailType -PageSize 2
    foreach ($SearchResult in $SearchResults) {
        # see if email already exists in the table
        $partitionKey1 = $emailType
        $cleanId = $SearchResult.Identity.Replace("\","_")
        $tableRow = Get-AzTableRow -table $cloudTable -partitionKey $partitionKey1 -rowKey $cleanId
        if($null -ne $tableRow)
        {
            "Email already exists in the table: {0}" -f $SearchResult.Identity | Write-Host -ForegroundColor Yellow
            continue
        } else {
            "Email does not exist in the table: {0}" -f $SearchResult.Identity | Write-Host -ForegroundColor Green

            try{
                $e = Export-QuarantineMessage -Identity $SearchResult.Identity -ErrorAction Continue
                "Successfully exported email {0}" -f $SearchResult.Identity  | Write-Host -ForegroundColor Blue
            } catch {
                "Error exporting email {0}" -f $SearchResult.Identity | Write-Host -ForegroundColor Red
                break
            }
                
        # Write file to disk
        $filepath = "{0}\{1}.eml" -f $TempFolder,$cleanId 
        [System.Text.Encoding]::Ascii.GetString([System.Convert]::FromBase64String($e.eml)) | Out-File $filepath -Encoding ascii

        # Send file to Azure Storage Blob
        $w = Set-AzStorageBlobContent -Blob $cleanId -File $filepath -Context $ctx -Container $env:ContainerName -Force

        # Add the email to the Azure Storage Table Index
        $tableAdd = Add-AzTableRow `
            -table $cloudTable `
            -partitionKey $partitionKey1 `
            -rowKey ($cleanId) -property @{"type"=$emailType;"length"=$w.Length}
        }

        $attachements += @(
            [pscustomobject]@{"@odata.type"="#microsoft.graph.fileAttachment";
            "name"= "$filepath";
            "contentType"= "text/plain";
            "contentBytes"= "$e.eml"}
        )
        
        $TotalEmails++
        $TotalLength += $w.Length

    }
}



$TotalEmails | Write-Host -ForegroundColor Green
$TotalLength | Write-Host -ForegroundColor Green


#############################################################################
## Logon to API to grap token
#############################################################################

$resourceAppIdUri = 'https://graph.microsoft.com/.default'
$oAuthUri = "https://login.microsoftonline.com/$env:tenantId/oauth2/v2.0/token"

$authBody = [Ordered] @{
  scope = $resourceAppIdUri
  client_id = $env:clientId
  client_secret = $env:clientSecret
  grant_type = 'client_credentials'
}
$authResponse = Invoke-RestMethod -Method Post -Uri $oAuthUri -Body $authBody -ErrorAction Stop
$token = $authResponse | Select-Object -ExpandProperty access_token

#############################################################################
## Send Mail
#############################################################################
# Set the email address to which the forwarded emails will be sent
$To = "admin@m365x59205144.onmicrosoft.com"
$From = "btirch@tirchfamily.com"

# Set the subject and body of the forwarded email
$Subject = "Phishing Emails"
$Body = "The attached file contains the phishing emails that were found in the quarantine."

$url = "https://graph.microsoft.com/v1.0/users/$From/sendMail"

$Body = @{
    "message" = @{
        "subject" = $Subject
        "body" = @{
            "contentType" = "Text"
            "content" = $Body
        }
        "toRecipients" = @(
            @{
                "emailAddress" = @{
                    "address" = $To
                }
            }
        )
        "attachments" = "" #$attachments
    }
} | ConvertTo-Json -Depth 99

$BodyJsonsend = @"
                    {
                        "message": {
                          "subject": "Hello World from Microsoft Graph API",
                          "body": {
                            "contentType": "HTML",
                            "content": "This Mail is sent via Microsoft <br>
                            GRAPH <br>
                            API<br>
                            
                            "
                          },
                          "toRecipients": [
                            {
                              "emailAddress": {
                                "address": "$To"
                              }
                            }
                          ]
                        },
                        "saveToSentItems": "false"
                      }
"@

$headers = @{
    'Content-Type' = 'application/json'
    'Accept' = 'application/json'
    'Authorization' = "Bearer $token"
}

try{
    $response = Invoke-WebRequest -Method Post -Body $BodyJsonsend -Uri $url -Headers $headers -ErrorAction Stop
    $data =  ($response | ConvertFrom-Json).results | ConvertTo-Json -Depth 99
    return 1
} catch {
    "Error sending email: {0}" -f $data.statuscode | Write-Host 
    return $null
}


# Delete the temporary folder
Remove-Item -Recurse -Force $TempFolder
# Clean up the Azure Storage Account and Table
# Get-AzStorageBlob -Container $env:ContainerName -Context $ctx | Remove-AzStorageBlob
# Get-AzTableRow -table $cloudTable | Remove-AzTableRow -table $cloudTable
