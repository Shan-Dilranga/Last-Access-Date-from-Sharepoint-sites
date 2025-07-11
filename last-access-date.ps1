# Connect to SharePoint Online  
$Tenant = "#your tenant id"
$ClientId = "#Your client id"
$SiteURL = "https://your-web-site.sharepoint.com/sites/sitename"

# Connect to SharePoint
Connect-PnPOnline -Url $SiteURL -ClientId $ClientId -Interactive

# Connect to Exchange Online for Audit Logs
Connect-ExchangeOnline -UserPrincipalName "ashan.dilhara@bcstechnology.com.au"

# Specify the output CSV file path
$OutputCSV = "C:\output-path\SharePoint_FileSizes.csv"

# Initialize an array to store file details
$FileDetails = @()

# Get all document libraries
$Libraries = Get-PnPList | Where-Object { $_.BaseType -eq "DocumentLibrary" }

# Fetch audit logs once for better performance (increased result size)
$AuditLogs = Search-UnifiedAuditLog -StartDate (Get-Date).AddMonths(-9) -EndDate (Get-Date) -Operations FileAccessed -ResultSize 5000

# Debugging: Check if audit logs are retrieved
Write-Host "Retrieved $($AuditLogs.Count) audit log entries"

# Loop through each library and get file sizes
foreach ($Library in $Libraries) {
    $startTime = Get-Date  # Start tracking time
    Write-Host "Processing library: $($Library.Title)"

    # Get list items (Limit processing to 100 files per library to optimize speed)
    $Files = Get-PnPListItem -List $Library.Title -PageSize 500 -Fields "FileLeafRef", "FileRef", "File_x0020_Size", "Created", "Modified" | Select-Object -First 200

    foreach ($File in $Files) {
        $FileSize = $File["File_x0020_Size"]
        $FilePath = $File["FileRef"]
        $CreatedDate = $File["Created"]
        $ModifiedDate = $File["Modified"]
        $LastAccessDate = "No Access Logs"

        # Debugging: Print file paths to verify matching
        Write-Host "Checking logs for: $FilePath"

        # Find access logs in pre-fetched data (case-insensitive match)
        $AuditLog = $AuditLogs | ForEach-Object {
            $AuditEntry = $_.AuditData | ConvertFrom-Json
            if ($AuditEntry.PSObject.Properties['ObjectId']) {
                $FullFileUrl = $AuditEntry.ObjectId  # Full URL from logs
                $SiteUrl = $AuditEntry.SiteUrl       # Site URL from logs
                $RelativePath = $FullFileUrl -replace [regex]::Escape($SiteUrl), ''  # Extract relative path

                if ($RelativePath -eq $FilePath) { 
                    $AuditEntry
                }
            }
        } | Sort-Object CreationTime -Descending | Select-Object -First 1

        if ($AuditLog) {
            $LastAccessDate = $AuditLog.CreationTime  # Extract CreationTime correctly
        }

        # Add file details to the array
        $FileDetails += [PSCustomObject]@{
            Library        = $Library.Title
            FilePath       = $FilePath
            FileSizeKB     = [math]::Round($FileSize / 1024, 2)  # Convert bytes to KB
            FileSizeMB     = [math]::Round($FileSize / 1MB, 2)   # Convert bytes to MB
            CreatedDate    = $CreatedDate
            ModifiedDate   = $ModifiedDate
            LastAccessDate = $LastAccessDate
        }
    }

    $endTime = Get-Date  # Stop tracking time
    $timeTaken = ($endTime - $startTime).TotalSeconds
    Write-Host "✅ Finished processing library: $($Library.Title) in $timeTaken seconds"
}

# Export the data to CSV
$FileDetails | Export-Csv -Path $OutputCSV -NoTypeInformation

Write-Host "File details exported to: $OutputCSV"
