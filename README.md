# Last-Access-Date-from-Sharepoint-sites
This shell script provides the last access date of share point sites

# SharePoint File Access Audit Script

This PowerShell script connects to a SharePoint Online site, retrieves file information from all document libraries, checks when each file was last accessed using Exchange Online audit logs, and exports the data to a CSV file.

## 📌 Features

- Connects securely to SharePoint Online using PnP PowerShell with App Registration and interactive login.
- Connects to Exchange Online using `Connect-ExchangeOnline`.
- Retrieves all document libraries from a specified SharePoint site.
- Lists up to 200 files per library with their metadata:
  - File name
  - File size (KB and MB)
  - Created date
  - Modified date
- Searches audit logs from the past **9 months** for `FileAccessed` events.
- Matches file access logs to each file using relative path matching.
- Exports results to a CSV file.

## ⚙️ Prerequisites

- **PnP.PowerShell** module installed (`Install-Module PnP.PowerShell`)
- **ExchangeOnlineManagement** module installed (`Install-Module ExchangeOnlineManagement`)
- App registration with the necessary permissions to connect to SharePoint Online
- Exchange Online auditing must be enabled
- PowerShell 7+ is recommended for performance

## 🔐 Authentication

- SharePoint: Uses `Connect-PnPOnline -ClientId -Interactive` (requires Azure AD App Registration)
- Exchange Online: Uses `Connect-ExchangeOnline` with your UPN

## 🧾 Audit Log Settings

- Retrieves up to **5000** audit log entries related to `FileAccessed` operations in the last 9 months.
- Searches for access records that match each file based on URL path comparison.

## 📤 Output

The script exports the collected data to a CSV file with the following columns:

| Column         | Description                           |
|----------------|---------------------------------------|
| Library        | Name of the SharePoint library        |
| FilePath       | Server-relative file path             |
| FileSizeKB     | File size in kilobytes (KB)           |
| FileSizeMB     | File size in megabytes (MB)           |
| CreatedDate    | Date when the file was created        |
| ModifiedDate   | Date when the file was last modified  |
| LastAccessDate | Date when the file was last accessed (based on audit logs) |

## 📁 Example Output Location


