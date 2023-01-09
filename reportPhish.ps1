$tenantId = $env:tenantId
$clientId = $env:clientId
$appSecret = $env:clientSecret
$strRG = $env:strRG
$strAccountName = $env:strAccountName

# Connect to Exchange Online
# Set the credentials for connecting to Exchange Online
$User = $env:username
$Password = ConvertTo-SecureString $env:passwd -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential($User, $Password)
# Connect to Exchange Online

# Connect-ExchangeOnline -Credential $Credential

# Connect to Azure
# Login-AzAccount 

# Get the Azure Storage Account Context
$ctx = New-AzStorageContext -ConnectionString $env:AzureWebJobsStorage

# Search for the phishing emails in the quarantine and export them to the temporary folder
# https://learn.microsoft.com/en-us/powershell/module/exchange/export-quarantinemessage?view=exchange-ps
# $SearchResults = Get-QuarantineMessage -Type "HighConfPhish" -PageSize 1

$SearchResults = Get-QuarantineMessage -Type "HighConfPhish" -PageSize 1
foreach ($SearchResult in $SearchResults) {
    try{
        $e = Export-QuarantineMessage -Identity $SearchResult.Identity -ErrorAction Stop
        $cleanId = $SearchResult.Identity.Replace("\","_")
        
        # Write file to disk
        #$filepath = "{0}\{1}.eml" -f $TempFolder,$cleanId 
        #[System.Text.Encoding]::Ascii.GetString([System.Convert]::FromBase64String($e.eml)) | Out-File $filepath -Encoding ascii
        
        # Send file to Azure Storage Blob
        #[System.Text.Encoding]::Ascii.GetString([System.Convert]::FromBase64String($e.eml)) | Set-AzStorageBlobContent -Blob $cleanId -Container $container -Context $ctx

        "Successfully exported email: {0}" -f $SearchResult.Identity | Write-Host -ForegroundColor Blue
    } catch {
        "Error exporting email: {0}" -f $SearchResult.Identity | Write-Host -ForegroundColor Red
    }
   
}


