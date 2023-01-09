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

# Get Adv Hunting Table Names from Azure Storage Table Service
$cloudTable = (Get-AzStorageTable -Name $tableName -Context $ctx).CloudTable

# Create a temporary folder to store the exported emails
$TempFolder = "$env:Temp\PhishingEmails"
New-Item -ItemType Directory -Path $TempFolder | Out-Null

# Search for the phishing emails in the quarantine and export them to the temporary folder
# https://learn.microsoft.com/en-us/powershell/module/exchange/export-quarantinemessage?view=exchange-ps

$emailTypes = @("HighConfPhish","Phish")

foreach($emailType in $emailTypes)
{
    $SearchResults = Get-QuarantineMessage -Type $emailType
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

    }
}

# Delete the temporary folder
Remove-Item -Recurse -Force $TempFolder
