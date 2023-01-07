# Connect to Exchange Online
# Connect-ExchangeOnline 

# Set the email address to which the forwarded emails will be sent
$To = "kenrward@gmail.com"
$From = "test.eop@tirchfamily.com"

# Set the subject and body of the forwarded email
$Subject = "Phishing Emails"
$Body = "The attached file contains the phishing emails that were found in the quarantine."

# Create a temporary folder to store the exported emails
$TempFolder = "$env:Temp\PhishingEmails"
New-Item -ItemType Directory -Path $TempFolder | Out-Null
$i = 0

# Search for the phishing emails in the quarantine and export them to the temporary folder
# https://learn.microsoft.com/en-us/powershell/module/exchange/export-quarantinemessage?view=exchange-ps
$SearchResults = Get-QuarantineMessage -QuarantineTypes HighConfPhish
foreach ($SearchResult in $SearchResults) {
    $e = Export-QuarantineMessage -Identity $SearchResult.Identity
    $e.BodyEncoding
    $filepath = "{0}\{1}.eml" -f $TempFolder,$i
    $e | Select-Object -ExpandProperty Eml | Out-File $filepath -Encoding ascii
    $i++
}

# Create the forwarded email with the exported emails as attachments
# Must use the graph API to send the email
# https://practical365.com/upgrade-powershell-scripts-sendmailmessage/
$tenantId = $env:tenantId
$clientId = $env:clientId
$appSecret = $env:clientSecret

#############################################################################
## Logon to API to grap token
#############################################################################
function Get-AuthToken{
    [cmdletbinding()]
        Param(
            [Parameter(Mandatory = $true, Position = 0)]
            [string]$clientId,
            [parameter(Mandatory = $true, Position = 1)]
            [string]$appSecret,
            [Parameter(Mandatory = $true, Position = 2)]
            [string]$tenantId
        )

$resourceAppIdUri = 'https://graph.microsoft.com/.default'
$oAuthUri = "https://login.microsoftonline.com/$tenantId/oauth2/token"

$authBody = [Ordered] @{
  scope = $resourceAppIdUri
  client_id = $clientId
  client_secret = $appSecret
  grant_type = 'client_credentials'
}
$authResponse = Invoke-RestMethod -Method Post -Uri $oAuthUri -Body $authBody -ErrorAction Stop
$token = $authResponse | Select-Object -ExpandProperty access_token
return $token
}

# Delete the temporary folder
Remove-Item -Recurse -Force $TempFolder