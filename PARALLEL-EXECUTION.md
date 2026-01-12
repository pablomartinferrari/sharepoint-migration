# Running Multiple Instances in Parallel

The script is designed to support running multiple instances simultaneously, allowing you to process different folders in parallel for faster migration validation.

## Quick Start

### Option 1: Auto-Generated Report Names (Recommended)

Simply run multiple instances without specifying `-ReportPath`. Each will generate a unique report name:

```powershell
# PowerShell Window 1
.\Compare-Migration.ps1 -ConfigPath "config-clients.json"

# PowerShell Window 2 (run simultaneously)
.\Compare-Migration.ps1 -ConfigPath "config-projects.json"

# PowerShell Window 3 (run simultaneously)
.\Compare-Migration.ps1 -ConfigPath "config-archive.json"
```

Each instance will create reports like:
- `migration-report-clients-20240115-143022.csv`
- `migration-report-projects-20240115-143025.csv`
- `migration-report-archive-20240115-143028.csv`

### Option 2: Explicit Report Names

You can also specify unique report paths manually:

```powershell
# PowerShell Window 1
.\Compare-Migration.ps1 -ConfigPath "config-clients.json" -ReportPath "report-clients.csv"

# PowerShell Window 2
.\Compare-Migration.ps1 -ConfigPath "config-projects.json" -ReportPath "report-projects.csv"
```

## Configuration Files

Create separate config files for each folder you want to process:

### config-clients.json
```json
{
  "TenantId": "your-tenant-id",
  "ClientId": "your-client-id",
  "Thumbprint": "your-thumbprint",
  "SharePointSiteUrl": "https://yourtenant.sharepoint.com/sites/yoursite",
  "FileServerPath": "G:\\shared\\clients",
  "LibraryName": "Documents"
}
```

### config-projects.json
```json
{
  "TenantId": "your-tenant-id",
  "ClientId": "your-client-id",
  "Thumbprint": "your-thumbprint",
  "SharePointSiteUrl": "https://yourtenant.sharepoint.com/sites/yoursite",
  "FileServerPath": "G:\\shared\\projects",
  "LibraryName": "Documents"
}
```

## Best Practices

1. **Different Config Files**: Each instance should use a different config file (or at least different `FileServerPath`)

2. **Unique Report Paths**: Either omit `-ReportPath` (auto-generate) or specify unique names to avoid overwriting

3. **Same SharePoint Site**: All instances can connect to the same SharePoint site simultaneously (no conflicts)

4. **Different Folders**: Process different folders to maximize parallelization benefits

5. **Resource Considerations**: 
   - Each instance uses memory and CPU
   - Network bandwidth is shared
   - SharePoint API has rate limits (but multiple read operations are usually fine)

## Example: Processing Multiple Shares

If you have multiple file server shares to migrate:

```powershell
# Share 1: Clients
Start-Process powershell -ArgumentList "-NoExit", "-Command", ".\Compare-Migration.ps1 -ConfigPath 'config-clients.json'"

# Share 2: Projects  
Start-Process powershell -ArgumentList "-NoExit", "-Command", ".\Compare-Migration.ps1 -ConfigPath 'config-projects.json'"

# Share 3: Archive
Start-Process powershell -ArgumentList "-NoExit", "-Command", ".\Compare-Migration.ps1 -ConfigPath 'config-archive.json'"
```

This will open three separate PowerShell windows, each processing a different share.

## Limitations

- **Same SharePoint Site**: All instances connect to the same SharePoint site (fine for read operations)
- **File Server Access**: Read-only access, so multiple instances can scan different folders simultaneously
- **No Conflicts**: Each instance is independent - no shared state or locks

## Monitoring Progress

Each instance will show its own progress:
- File scanning progress
- Comparison results
- Locked file warnings
- Final summary

Check each PowerShell window to monitor progress, or review the generated report files.
