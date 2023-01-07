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

# Delete the temporary folder
Remove-Item -Recurse -Force $TempFolder