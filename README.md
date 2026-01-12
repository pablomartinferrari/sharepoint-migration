# SharePoint Migration Comparison Tool

This PowerShell script compares files between a file server and SharePoint to identify missing or outdated files after migration. It respects file modification dates to avoid overwriting newer versions in SharePoint.

## Features

- ✅ Compares files between file server shares and SharePoint
- ✅ Identifies missing files that need migration
- ✅ Detects files newer on server (can be migrated)
- ✅ Skips files newer in SharePoint (protects user edits)
- ✅ Handles locked/inaccessible files gracefully (reports them separately)
- ✅ Uses certificate-based authentication (no passwords)
- ✅ Generates detailed CSV report and summary
- ✅ Handles large datasets efficiently

## Prerequisites

1. **PowerShell 5.1 or later** (included with Windows Server)
2. **PnP.PowerShell module** (recommended for best compatibility):

   ```powershell
   Install-Module -Name PnP.PowerShell -Scope CurrentUser
   ```

   Alternatively, you can use the Azure PowerShell module:

   ```powershell
   Install-Module -Name Az.Accounts -Scope CurrentUser
   ```

3. **Certificate installed** in the certificate store (CurrentUser\My or LocalMachine\My)
4. **App Registration** with:
   - Client ID
   - Tenant ID
   - Certificate thumbprint
   - `Sites.FullControl.All` or `Sites.Read.All` permission

## Setup

1. **Copy the configuration template:**

   ```powershell
   Copy-Item config.json.example config.json
   ```

2. **Edit `config.json`** with your details:

   - `TenantId`: Your Azure AD tenant ID
   - `ClientId`: Your app registration client ID
   - `Thumbprint`: Thumbprint of the certificate installed on the server
   - `SharePointSiteUrl`: Full URL to your SharePoint site
   - `FileServerPath`: Root path on file server (e.g., `G:\shared\clients`). The last folder name will be preserved in SharePoint paths.
   - `LibraryName`: SharePoint document library name (usually "Documents")
   - `StartDate`: (Optional) Filter files created/modified after this date. Format: `"2024-01-01"` or `"2024-01-01 00:00:00"`
   - `EndDate`: (Optional) Filter files created/modified before this date. If `null` and `StartDate` is set, defaults to current date/time. Format: `"2024-12-31"` or `"2024-12-31 23:59:59"`

   **Path Mapping Example:**

   - File Server: `G:\shared\clients\client 1\my doc.pdf`
   - Root Path: `G:\shared\clients`
   - SharePoint Path: `clients\client 1\my doc.pdf` (root folder "clients" is included)

   **Date Filtering:**

   - Files are included if their **Created** OR **Modified** date falls within the specified range
   - If only `StartDate` is provided, `EndDate` defaults to the current date/time
   - Example: `"StartDate": "2024-01-01"` will include all files modified/created since January 1, 2024

3. **Ensure the certificate is installed:**
   - The certificate must be in either:
     - Current User certificate store: `Cert:\CurrentUser\My`
     - Local Machine certificate store: `Cert:\LocalMachine\My`
   - You can verify with:
     ```powershell
     Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object { $_.Thumbprint -eq "YOUR-THUMBPRINT" }
     ```

## Usage

Run the script from PowerShell:

```powershell
.\Compare-Migration.ps1 -ConfigPath "config.json" -ReportPath "migration-report.csv"
```

### Parameters

- `-ConfigPath` (Required): Path to your configuration JSON file
- `-ReportPath` (Optional): Path for the output CSV report. If not specified, auto-generates a unique name based on folder name and timestamp (e.g., `migration-report-clients-20240115-143022.csv`)

### Running Multiple Instances

You can run multiple instances of the script simultaneously to process different folders in parallel:

```powershell
# Terminal 1: Process clients folder
.\Compare-Migration.ps1 -ConfigPath "config-clients.json" -ReportPath "report-clients.csv"

# Terminal 2: Process projects folder (simultaneously)
.\Compare-Migration.ps1 -ConfigPath "config-projects.json" -ReportPath "report-projects.csv"
```

**Requirements for parallel execution:**

- Each instance must use a **different config file** (or at least different `FileServerPath`)
- Each instance should use a **different report path** (or omit `-ReportPath` to auto-generate unique names)
- Each instance connects independently to SharePoint (no conflicts)
- File server access is read-only, so multiple instances can scan different folders simultaneously

### Example

```powershell
.\Compare-Migration.ps1 -ConfigPath "config.json" -ReportPath "report-$(Get-Date -Format 'yyyyMMdd').csv"
```

## Output

The script generates two files:

1. **CSV Report** (`migration-report.csv`): Detailed comparison of all files with columns:

   - `Status`: Missing, NewerOnServer, NewerInSharePoint, SizeMismatch, Locked
   - `FilePath`: Relative path of the file
   - `ServerSize`: File size on file server (bytes)
   - `ServerModified`: Last modified date on file server
   - `SharePointSize`: File size in SharePoint (bytes)
   - `SharePointModified`: Last modified date in SharePoint
   - `Action`: Recommended action (Migrate, CanMigrate, Skip, Review, ReviewLocked)
   - `ErrorMessage`: Error message for locked files (if applicable)

2. **Summary Report** (`migration-report-summary.txt`): High-level statistics and next steps

## Understanding the Results

### Status Types

- **Missing**: File exists on file server but not in SharePoint → **Action: Migrate**
- **NewerOnServer**: File is newer on file server → **Action: CanMigrate** (safe to migrate)
- **NewerInSharePoint**: File is newer in SharePoint → **Action: Skip** (protects user edits)
- **SizeMismatch**: Same modification time but different sizes → **Action: Review** (manual inspection needed)
- **Locked**: File is locked or inaccessible on file server (e.g., open in another application) → **Action: ReviewLocked** (close the file and re-run the script)

### Next Steps

1. **Review the CSV report** to see which files need attention
2. **Files marked "Migrate" or "CanMigrate"** can be safely migrated
3. **Files marked "Skip"** should NOT be migrated (they have newer versions in SharePoint)
4. **Files marked "Review"** need manual inspection to determine the correct version

## Troubleshooting

### Certificate Not Found

```
Error: Certificate with thumbprint XXX not found
```

**Solution**: Verify the certificate is installed and the thumbprint in config.json matches exactly.

### Cannot Connect to SharePoint

```
Error: Failed to authenticate
```

**Solutions**:

- Verify your app registration has `Sites.FullControl.All` or `Sites.Read.All` permission (with admin consent granted)
- Ensure the certificate is valid and not expired
- Check that the SharePoint site URL is correct
- The script will auto-install PnP.PowerShell if missing

### File Server Path Not Accessible

```
Error: File server path not accessible
```

**Solution**:

- Ensure the script is running with appropriate permissions
- Verify the UNC path is correct and accessible from the server
- Check network connectivity to the file server

### Locked Files

If you see warnings about locked files:

- Files that are open in applications (Excel, Word, etc.) cannot be read
- The script will still report these files with status "Locked"
- Close the files and re-run the script to get complete information
- Locked files are included in the report but with limited metadata

### Large Dataset Performance

For very large datasets (1.5TB+), the script may take some time. The script shows progress every 1000 files scanned. Consider:

- Running during off-hours
- Using PowerShell jobs for parallel processing (future enhancement)

## Security Notes

- The script uses certificate-based authentication (no passwords stored)
- Configuration file contains sensitive information - protect it appropriately
- Consider using Group Managed Service Accounts (gMSA) for service accounts
- Review and limit the permissions of your app registration to minimum required

## License

This script is provided as-is for migration validation purposes.
