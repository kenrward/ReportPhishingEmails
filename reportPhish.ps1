#$tenantId = $env:tenantId
#$clientId = $env:clientId
#$appSecret = $env:clientSecret
#$strRG = $env:strRG
#$strAccountName = $env:strAccountName

# Connect to Exchange Online
# Set the credentials for connecting to Exchange Online
$Password = ConvertTo-SecureString $env:passwd -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential($env:username, $Password)

# Connect to Exchange Online
Connect-ExchangeOnline -Credential $Credential -UseRPSSession


# Get the Azure Storage Account Context
$ctx = New-AzStorageContext -ConnectionString $env:AzureWebJobsStorage

# Create a temporary folder to store the exported emails
$TempFolder = "$env:Temp\PhishingEmails"
New-Item -ItemType Directory -Path $TempFolder | Out-Null

# Search for the phishing emails in the quarantine and export them to the temporary folder
# https://learn.microsoft.com/en-us/powershell/module/exchange/export-quarantinemessage?view=exchange-ps
# $SearchResults = Get-QuarantineMessage -Type "HighConfPhish" -PageSize 1

$SearchResults = Get-QuarantineMessage -Type "Phish" 
foreach ($SearchResult in $SearchResults) {
    try{
        $e = Export-QuarantineMessage -Identity $SearchResult.Identity -ErrorAction Continue
        $cleanId = $SearchResult.Identity.Replace("\","_")
        
        # Write file to disk
        $filepath = "{0}\{1}.eml" -f $TempFolder,$cleanId 
        [System.Text.Encoding]::Ascii.GetString([System.Convert]::FromBase64String($e.eml)) | Out-File $filepath -Encoding ascii
        
        # Send file to Azure Storage Blob
        Set-AzStorageBlobContent -Blob $cleanId -File $filepath -Context $ctx -Container $env:ContainerName 

        "Successfully exported email: {0}" -f $SearchResult.Identity | Write-Host -ForegroundColor Blue
    } catch {
        "Error exporting email: {0}" -f $SearchResult.Identity | Write-Host -ForegroundColor Red
    }
   
}

# Delete the temporary folder
Remove-Item -Recurse -Force $TempFolder
