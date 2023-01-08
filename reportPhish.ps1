$tenantId = $env:tenantId
$clientId = $env:clientId
$appSecret = $env:clientSecret

# Connect to Exchange Online
# Set the credentials for connecting to Exchange Online
$User = $env:username
$Password = ConvertTo-SecureString $env:passwd -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential($User, $Password)
# Connect to Exchange Online

# Connect-ExchangeOnline -Credential $Credential

# Set the email address to which the forwarded emails will be sent
$To = "kenrward@outlook.com"
$From = "admin@m365x59205144.onmicrosoft.com"

# Set the subject and body of the forwarded email
$Subject = "Phishing Emails"
$Body = "The attached file contains the phishing emails that were found in the quarantine."

# Create a temporary folder to store the exported emails
# $TempFolder = "$env:Temp\PhishingEmails"
# New-Item -ItemType Directory -Path $TempFolder | Out-Null
$attachements = $null


# Search for the phishing emails in the quarantine and export them to the temporary folder
# https://learn.microsoft.com/en-us/powershell/module/exchange/export-quarantinemessage?view=exchange-ps
$SearchResults = Get-QuarantineMessage -Type "HighConfPhish" -PageSize 1
foreach ($SearchResult in $SearchResults) {
    try{
        $e = Export-QuarantineMessage -Identity $SearchResult.Identity -ErrorAction Stop
        $cleanId = $SearchResult.Identity.Replace("\","_")
        #$filepath = "{0}\{1}.eml" -f $TempFolder,$i
        $attachements += @(
            [pscustomobject]@{"@odata.type"="#microsoft.graph.fileAttachment";
            "name"= "$cleanId";
            "contentType"= "text/plain";
            "contentBytes"= $e.eml}
        )
        # Write file to disk
        #[System.Text.Encoding]::Ascii.GetString([System.Convert]::FromBase64String($e.eml)) | Out-File $filepath -Encoding ascii
        "Successfully exported email: {0}" -f $SearchResult.Identity | Write-Host -ForegroundColor Blue
    } catch {
        "Error exporting email: {0}" -f $SearchResult.Identity | Write-Host -ForegroundColor Red
    }
   
}



# Create the forwarded email with the exported emails as attachments
# Must use the graph API to send the email
# https://practical365.com/upgrade-powershell-scripts-sendmailmessage/


#############################################################################
## Logon to API to grap token
#############################################################################

$resourceAppIdUri = 'https://graph.microsoft.com/.default'
$oAuthUri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

$authBody = [Ordered] @{
  scope = $resourceAppIdUri
  client_id = $clientId
  client_secret = $appSecret
  grant_type = 'client_credentials'
}
$authResponse = Invoke-RestMethod -Method Post -Uri $oAuthUri -Body $authBody -ErrorAction Stop
$token = $authResponse | Select-Object -ExpandProperty access_token

#############################################################################
## Send Mail
#############################################################################

$url = "https://graph.microsoft.com/v1.0/users/$From/sendMail"
$attachements = $attachements | ConvertTo-Json -Depth 10

$BodyJsonsend = @"
                    {
                        "message": {
                          "subject": "Hello World from Microsoft Graph API",
                          "body": {
                            "contentType": "HTML",
                            "content": "This Mail is sent via Microsoft <br>GRAPH <br>API<br>"
                          },
                          "toRecipients": [
                            {
                              "emailAddress": {
                                "address": "$To"
                              }
                            }
                          ],
                          "attachments": [
                              $attachements
                          ]
                        },
                        "saveToSentItems": "true"
                      }
"@ 

<# $headers = @{
    'Content-Type' = 'application/json'
    'Accept' = 'application/json'
    'Authorization' = "Bearer $token"
} #>

# $env:requestURL
$url = $env:requestURL

try{
    $response = Invoke-WebRequest -Method Post -Body $attachements -Uri $url -ErrorAction Stop
    "Successfully sent email: {0}" -f $response.StatusDescription | Write-Host -ForegroundColor Blue
} catch {
    "Error sending email: {0}" -f $response.StatusDescription | Write-Host -ForegroundColor Red
}

# Delete the temporary folder
# Remove-Item -Recurse -Force $TempFolder
