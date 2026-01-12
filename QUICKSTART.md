# Quick Start Guide

## Prerequisites Check

1. **Install PnP.PowerShell module:**
   ```powershell
   Install-Module -Name PnP.PowerShell -Scope CurrentUser
   ```

2. **Verify certificate is installed:**
   ```powershell
   # Check Current User store
   Get-ChildItem -Path Cert:\CurrentUser\My | Select-Object Thumbprint, Subject, NotAfter
   
   # Check Local Machine store
   Get-ChildItem -Path Cert:\LocalMachine\My | Select-Object Thumbprint, Subject, NotAfter
   ```

## Setup (5 minutes)

1. **Copy the config template:**
   ```powershell
   Copy-Item config.json.example config.json
   ```

2. **Edit `config.json` with your details:**
   - `TenantId`: Found in Azure Portal → Azure Active Directory → Overview
   - `ClientId`: Found in Azure Portal → App registrations → Your app → Overview
   - `Thumbprint`: From the certificate you installed (no spaces)
   - `SharePointSiteUrl`: Full URL like `https://contoso.sharepoint.com/sites/mysite`
   - `FileServerPath`: Root path like `G:\shared\clients` (the last folder name "clients" will be preserved in SharePoint)
   - `LibraryName`: Usually "Documents" (optional, defaults to "Documents")
   - `StartDate`: (Optional) Filter by date - include files created/modified after this date, e.g., `"2024-01-01"`
   - `EndDate`: (Optional) Filter by date - include files created/modified before this date. If `null` and `StartDate` is set, defaults to now

   **Important:** The last folder name in `FileServerPath` will be included in SharePoint paths. For example:
   - File: `G:\shared\clients\client 1\file.pdf`
   - Root: `G:\shared\clients`
   - SharePoint: `clients\client 1\file.pdf`

   **Date Filtering:** Files are included if their Created OR Modified date is within the range. Useful for incremental migrations!

## Run the Comparison

```powershell
.\Compare-Migration.ps1 -ConfigPath "config.json"
```

If you don't specify `-ReportPath`, it will auto-generate a unique name like `migration-report-clients-20240115-143022.csv`.

This will create:
- `migration-report-*.csv` - Detailed file-by-file comparison
- `migration-report-*-summary.txt` - High-level summary

### Running Multiple Instances in Parallel

You can process multiple folders simultaneously by running multiple instances:

```powershell
# PowerShell window 1
.\Compare-Migration.ps1 -ConfigPath "config-clients.json"

# PowerShell window 2 (run simultaneously)
.\Compare-Migration.ps1 -ConfigPath "config-projects.json"
```

Each instance will generate its own unique report file automatically.

## Understanding the Results

### Status Column Meanings:
- **Missing** → File exists on server but not in SharePoint → **Migrate this file**
- **NewerOnServer** → File is newer on server → **Safe to migrate**
- **NewerInSharePoint** → File is newer in SharePoint → **DO NOT migrate** (protects user edits)
- **SizeMismatch** → Same date but different size → **Review manually**
- **Locked** → File is locked/inaccessible (e.g., open in Excel) → **Close file and re-run script**

### Next Steps:

1. **Filter the CSV for files to migrate:**
   ```powershell
   Import-Csv migration-report.csv | Where-Object { $_.Action -eq "Migrate" -or $_.Action -eq "CanMigrate" } | Export-Csv files-to-migrate.csv
   ```

2. **Review files marked "Review"** - these need manual inspection

3. **Use SharePoint Migration Tool** or your preferred method to migrate files from the filtered list

## Troubleshooting

### "Certificate not found"
- Verify thumbprint matches exactly (no spaces, uppercase)
- Check both CurrentUser and LocalMachine certificate stores
- Ensure certificate is not expired

### "Failed to connect with PnP.PowerShell"
- Verify app registration has `Sites.FullControl.All` or `Sites.Read.All` permission (with admin consent granted)
- Check that certificate thumbprint in Azure AD matches the installed certificate
- Ensure SharePoint site URL is correct
- The script will auto-install PnP.PowerShell if missing

### "File server path not accessible"
- Run PowerShell as administrator or with appropriate permissions
- Verify UNC path is accessible: `Test-Path "\\server\share"`
- Check network connectivity

## Performance Tips

For large datasets (1.5TB+):
- Script shows progress every 1000 files
- First run may take 30-60 minutes depending on file count
- Consider running during off-hours
- Results are saved incrementally (CSV is written at the end)
