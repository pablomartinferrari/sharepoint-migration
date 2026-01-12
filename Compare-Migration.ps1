<#
.SYNOPSIS
    Compares files between a file server and SharePoint to identify missing or outdated files.

.DESCRIPTION
    This script compares files from a file server share with files in SharePoint, identifying:
    - Files that exist on the file server but not in SharePoint
    - Files that exist in both but are newer on the file server (and can be migrated)
    - Files that are newer in SharePoint (will be skipped to avoid overwriting user edits)

.PARAMETER ConfigPath
    Path to the JSON configuration file containing SharePoint credentials and paths.

.PARAMETER ReportPath
    Path where the comparison report will be saved (default: migration-report.csv)

.EXAMPLE
    .\Compare-Migration.ps1 -ConfigPath "config.json" -ReportPath "report.csv"
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$ConfigPath,
    
    [Parameter(Mandatory = $false)]
    [string]$ReportPath = $null
)

# Error handling
$ErrorActionPreference = "Stop"

# Load configuration
Write-Host "Loading configuration from $ConfigPath..." -ForegroundColor Cyan
if (-not (Test-Path $ConfigPath)) {
    throw "Configuration file not found: $ConfigPath"
}

$config = Get-Content $ConfigPath | ConvertFrom-Json

# Auto-generate report path if not specified (useful for multiple instances)
if (-not $ReportPath) {
    # Generate unique report name based on folder name and timestamp
    $rootFolderName = Split-Path -Leaf $config.FileServerPath
    # Sanitize folder name for use in filename (remove invalid characters)
    $rootFolderName = $rootFolderName -replace '[<>:"/\\|?*]', '-'
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $ReportPath = "migration-report-$rootFolderName-$timestamp.csv"
    Write-Host "Auto-generated report path: $ReportPath" -ForegroundColor Gray
}

# Validate configuration
$requiredFields = @('TenantId', 'ClientId', 'Thumbprint', 'SharePointSiteUrl', 'FileServerPath')
foreach ($field in $requiredFields) {
    if (-not $config.$field) {
        throw "Missing required configuration field: $field"
    }
}

# Parse date filters (optional)
$startDate = $null
$endDate = $null
if ($config.StartDate) {
    try {
        $startDate = [DateTime]::Parse($config.StartDate)
        Write-Host "Date filter: Including files created/modified after $($startDate.ToString('yyyy-MM-dd'))" -ForegroundColor Cyan
    }
    catch {
        throw "Invalid StartDate format in config. Use format like '2024-01-01' or '2024-01-01 00:00:00'"
    }
}
if ($config.EndDate) {
    try {
        $endDate = [DateTime]::Parse($config.EndDate)
        Write-Host "Date filter: Including files created/modified before $($endDate.ToString('yyyy-MM-dd'))" -ForegroundColor Cyan
    }
    catch {
        throw "Invalid EndDate format in config. Use format like '2024-12-31' or '2024-12-31 23:59:59'"
    }
}
else {
    # Default end date to now if start date is specified
    if ($startDate) {
        $endDate = Get-Date
        Write-Host "Date filter: End date not specified, using current date/time" -ForegroundColor Cyan
    }
}

# Function to connect to SharePoint using certificate authentication
function Connect-SharePoint {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$Thumbprint,
        [string]$SiteUrl
    )
    
    Write-Host "Connecting to SharePoint..." -ForegroundColor Cyan
    
    # Check if PnP.PowerShell is available
    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
        Write-Host "PnP.PowerShell module not found. Installing..." -ForegroundColor Yellow
        Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
    }
    
    Import-Module PnP.PowerShell -ErrorAction Stop
    
    try {
        # Connect using certificate authentication (matching existing pattern)
        Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $TenantId -WarningAction SilentlyContinue -ErrorAction Stop
        
        # Verify connection
        $web = Get-PnPWeb
        Write-Host "Connected to: $($web.Title)" -ForegroundColor Green
        return $true
    }
    catch {
        throw "Failed to connect with PnP.PowerShell: $_"
    }
}

# Helper function to normalize file paths for comparison
function Normalize-Path {
    param([string]$Path)
    
    # Convert to lowercase and normalize separators
    $normalized = $Path.ToLower()
    $normalized = $normalized -replace '/', '\'
    $normalized = $normalized.TrimStart('\')
    
    return $normalized
}

# Function to get all files from SharePoint site using PnP PowerShell
function Get-SharePointFiles {
    param(
        [string]$SiteUrl,
        [string]$LibraryName = "Documents"
    )
    
    Write-Host "Scanning SharePoint files..." -ForegroundColor Cyan
    
    if (-not (Get-Module -Name PnP.PowerShell)) {
        throw "PnP.PowerShell module is required"
    }
    
    $files = @{}
    $fileCount = 0
    
    # Try to get the specified library, or default to "Documents"
    $list = Get-PnPList -Identity $LibraryName -ErrorAction SilentlyContinue
    if (-not $list) {
        Write-Warning "Library '$LibraryName' not found. Trying 'Documents'..."
        $list = Get-PnPList -Identity "Documents" -ErrorAction SilentlyContinue
    }
    
    if (-not $list) {
        Write-Warning "Default document library not found. Scanning all document libraries..."
        $lists = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and $_.Hidden -eq $false }  # Document library template
        foreach ($l in $lists) {
            Write-Host "  Scanning library: $($l.Title)" -ForegroundColor Gray
            Get-FilesFromList -List $l -SiteUrl $SiteUrl -Files $files -FileCount ([ref]$fileCount)
        }
    }
    else {
        Write-Host "  Scanning library: $($list.Title)" -ForegroundColor Gray
        Get-FilesFromList -List $list -SiteUrl $SiteUrl -Files $files -FileCount ([ref]$fileCount)
    }
    
    Write-Host "Found $($files.Count) files in SharePoint" -ForegroundColor Green
    return $files
}

# Helper function to get files from a SharePoint list
function Get-FilesFromList {
    param(
        [Microsoft.SharePoint.Client.List]$List,
        [string]$SiteUrl,
        [hashtable]$Files,
        [ref]$FileCount
    )
    
    $items = Get-PnPListItem -List $List -PageSize 5000 -Fields "FileRef", "FileLeafRef", "File_x0020_Size", "Modified"
    
    # Get library name/URL for path normalization
    $libraryServerRelativeUrl = $List.RootFolder.ServerRelativeUrl
    $libraryName = $List.Title
    
    foreach ($item in $items) {
        if ($item.FileSystemObjectType -eq "File") {
            # Get relative path from site URL
            $fileRef = $item.FieldValues.FileRef
            $filePath = $fileRef -replace [regex]::Escape($SiteUrl), ""
            $filePath = $filePath.TrimStart('/')
            
            # Remove library name prefix if present (e.g., "Shared Documents/" or "Documents/")
            # SharePoint paths might be like "Shared Documents/clients/file.pdf" or "Documents/clients/file.pdf"
            $filePath = $filePath -replace "^$([regex]::Escape($libraryName))/", ""
            $filePath = $filePath -replace "^Shared Documents/", ""
            $filePath = $filePath -replace "^Documents/", ""
            
            # Normalize path for comparison (lowercase, backslashes)
            $normalizedPath = Normalize-Path -Path $filePath
            
            $files[$normalizedPath] = @{
                Name = $item.FieldValues.FileLeafRef
                Size = [long]$item.FieldValues.File_x0020_Size
                Modified = [DateTime]$item.FieldValues.Modified
                Path = $filePath  # Keep original for display
                NormalizedPath = $normalizedPath
            }
            
            $FileCount.Value++
            if ($FileCount.Value % 1000 -eq 0) {
                Write-Host "    Scanned $($FileCount.Value) files..." -ForegroundColor Gray
            }
        }
    }
}

# Function to get all files from file server
function Get-FileServerFiles {
    param(
        [string]$RootPath,
        [DateTime]$StartDate = $null,
        [DateTime]$EndDate = $null
    )
    
    Write-Host "Scanning file server: $RootPath..." -ForegroundColor Cyan
    
    if (-not (Test-Path $RootPath)) {
        throw "File server path not accessible: $RootPath"
    }
    
    # Get the root folder name (last folder in the path)
    $rootFolderName = Split-Path -Leaf $RootPath
    
    Write-Host "Root folder name: $rootFolderName" -ForegroundColor Gray
    
    $files = @{}
    $lockedFiles = @{}
    $fileCount = 0
    $filteredCount = 0
    $lockedCount = 0
    
    Get-ChildItem -Path $RootPath -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
        $fileCount++
        $fileInfo = $null
        $isLocked = $false
        $errorMessage = $null
        
        # Try to access file properties - catch locked file errors
        try {
            # Force property access to detect locked files
            $null = $_.Length
            $null = $_.LastWriteTime
            $null = $_.CreationTime
            $fileInfo = $_
        }
        catch {
            # File is likely locked or inaccessible
            $isLocked = $true
            $errorMessage = $_.Exception.Message
            # Still try to get the path if possible
            try {
                $fileInfo = $_
            }
            catch {
                # Can't even get the file object - skip it
                $lockedCount++
                return
            }
        }
        
        if ($isLocked) {
            # File is locked - still report it but with limited info
            $relativePath = $fileInfo.FullName -replace [regex]::Escape($RootPath), ""
            $relativePath = $relativePath.TrimStart('\', '/')
            $sharePointRelativePath = "$rootFolderName\$relativePath"
            $normalizedPath = Normalize-Path -Path $sharePointRelativePath
            
            $lockedFiles[$normalizedPath] = @{
                Name = $fileInfo.Name
                Size = $null
                Modified = $null
                Created = $null
                Path = $relativePath
                SharePointPath = $sharePointRelativePath
                NormalizedPath = $normalizedPath
                FullPath = $fileInfo.FullName
                IsLocked = $true
                ErrorMessage = $errorMessage
            }
            $lockedCount++
            return
        }
        
        # File is accessible - proceed with normal processing
        # Apply date filter if specified
        if ($StartDate -or $EndDate) {
            # Check both Created and Modified dates - include if either falls within range
            $createdDate = $fileInfo.CreationTime
            $modifiedDate = $fileInfo.LastWriteTime
            
            $includeFile = $false
            
            # Check if created date is in range
            $createdInRange = $true
            if ($StartDate -and $createdDate -lt $StartDate) {
                $createdInRange = $false
            }
            if ($EndDate -and $createdDate -gt $EndDate) {
                $createdInRange = $false
            }
            
            # Check if modified date is in range
            $modifiedInRange = $true
            if ($StartDate -and $modifiedDate -lt $StartDate) {
                $modifiedInRange = $false
            }
            if ($EndDate -and $modifiedDate -gt $EndDate) {
                $modifiedInRange = $false
            }
            
            # Include file if either created or modified date is in range
            $includeFile = $createdInRange -or $modifiedInRange
            
            if (-not $includeFile) {
                $filteredCount++
                return
            }
        }
        
        # Get relative path from root
        $relativePath = $fileInfo.FullName -replace [regex]::Escape($RootPath), ""
        $relativePath = $relativePath.TrimStart('\', '/')
        
        # Create SharePoint path: root folder name + relative path
        $sharePointRelativePath = "$rootFolderName\$relativePath"
        
        # Normalize path for comparison (lowercase, backslashes)
        $normalizedPath = Normalize-Path -Path $sharePointRelativePath
        
        $files[$normalizedPath] = @{
            Name = $fileInfo.Name
            Size = $fileInfo.Length
            Modified = $fileInfo.LastWriteTime
            Created = $fileInfo.CreationTime
            Path = $relativePath  # Original relative path (without root folder)
            SharePointPath = $sharePointRelativePath  # Path as it should appear in SharePoint
            NormalizedPath = $normalizedPath
            FullPath = $fileInfo.FullName
            IsLocked = $false
        }
        
        if ($files.Count % 1000 -eq 0) {
            Write-Host "  Processed $($files.Count) files (scanned $fileCount total)..." -ForegroundColor Gray
        }
    }
    
    # Build return object with both accessible and locked files
    $result = @{
        Files = $files
        LockedFiles = $lockedFiles
        LockedCount = $lockedCount
    }
    
    $statusMsg = "Found $($files.Count) files on file server"
    if ($lockedCount -gt 0) {
        $statusMsg += ", $lockedCount locked/inaccessible files"
    }
    if ($StartDate -or $EndDate) {
        $statusMsg += " (filtered $filteredCount files by date)"
    }
    Write-Host $statusMsg -ForegroundColor Green
    
    if ($lockedCount -gt 0) {
        Write-Host "Warning: $lockedCount files are locked or inaccessible (will be reported separately)" -ForegroundColor Yellow
    }
    
    return $result
}

# Main comparison logic
Write-Host "`n=== Migration Comparison Tool ===" -ForegroundColor Yellow
Write-Host ""

# Connect to SharePoint
Connect-SharePoint -TenantId $config.TenantId -ClientId $config.ClientId -Thumbprint $config.Thumbprint -SiteUrl $config.SharePointSiteUrl

# Get files from both sources
$libraryName = if ($config.LibraryName) { $config.LibraryName } else { "Documents" }
$sharePointFiles = Get-SharePointFiles -SiteUrl $config.SharePointSiteUrl -LibraryName $libraryName
$fileServerResult = Get-FileServerFiles -RootPath $config.FileServerPath -StartDate $startDate -EndDate $endDate
$fileServerFiles = $fileServerResult.Files
$lockedFiles = $fileServerResult.LockedFiles

# Compare files
Write-Host "`nComparing files..." -ForegroundColor Cyan

$results = @()
$missingCount = 0
$newerOnServerCount = 0
$newerInSharePointCount = 0
$identicalCount = 0
$lockedCount = $fileServerResult.LockedCount

foreach ($filePath in $fileServerFiles.Keys) {
    $serverFile = $fileServerFiles[$filePath]
    
    if ($sharePointFiles.ContainsKey($filePath)) {
        $spFile = $sharePointFiles[$filePath]
        
        # Compare modification dates
        if ($serverFile.Modified -gt $spFile.Modified) {
            # File is newer on server - can be migrated
            $results += [PSCustomObject]@{
                Status = "NewerOnServer"
                FilePath = $filePath
                SharePointPath = $serverFile.SharePointPath
                ServerSize = $serverFile.Size
                ServerModified = $serverFile.Modified
                SharePointSize = $spFile.Size
                SharePointModified = $spFile.Modified
                Action = "CanMigrate"
            }
            $newerOnServerCount++
        }
        elseif ($spFile.Modified -gt $serverFile.Modified) {
            # File is newer in SharePoint - skip to avoid overwriting
            $results += [PSCustomObject]@{
                Status = "NewerInSharePoint"
                FilePath = $filePath
                SharePointPath = $serverFile.SharePointPath
                ServerSize = $serverFile.Size
                ServerModified = $serverFile.Modified
                SharePointSize = $spFile.Size
                SharePointModified = $spFile.Modified
                Action = "Skip"
            }
            $newerInSharePointCount++
        }
        else {
            # Same modification time (or very close)
            if ($serverFile.Size -ne $spFile.Size) {
                # Same time but different size - might need migration
                $results += [PSCustomObject]@{
                    Status = "SizeMismatch"
                    FilePath = $filePath
                    SharePointPath = $serverFile.SharePointPath
                    ServerSize = $serverFile.Size
                    ServerModified = $serverFile.Modified
                    SharePointSize = $spFile.Size
                    SharePointModified = $spFile.Modified
                    Action = "Review"
                }
            }
            else {
                $identicalCount++
            }
        }
    }
    else {
        # File doesn't exist in SharePoint - needs migration
        $results += [PSCustomObject]@{
            Status = "Missing"
            FilePath = $filePath
            SharePointPath = $serverFile.SharePointPath
            ServerSize = $serverFile.Size
            ServerModified = $serverFile.Modified
            SharePointSize = $null
            SharePointModified = $null
            Action = "Migrate"
        }
        $missingCount++
    }
}

# Add locked files to results
foreach ($filePath in $lockedFiles.Keys) {
    $lockedFile = $lockedFiles[$filePath]
    
    # Check if file exists in SharePoint (even though we can't read it from server)
    $existsInSP = $sharePointFiles.ContainsKey($filePath)
    
    $results += [PSCustomObject]@{
        Status = "Locked"
        FilePath = $filePath
        SharePointPath = $lockedFile.SharePointPath
        ServerSize = $null
        ServerModified = $null
        SharePointSize = if ($existsInSP) { $sharePointFiles[$filePath].Size } else { $null }
        SharePointModified = if ($existsInSP) { $sharePointFiles[$filePath].Modified } else { $null }
        Action = if ($existsInSP) { "Review" } else { "ReviewLocked" }
        ErrorMessage = $lockedFile.ErrorMessage
    }
}

# Generate report
Write-Host "`n=== Comparison Results ===" -ForegroundColor Yellow
Write-Host "Total files on server: $($fileServerFiles.Count)" -ForegroundColor White
if ($lockedCount -gt 0) {
    Write-Host "Locked/inaccessible files: $lockedCount" -ForegroundColor Yellow
}
Write-Host "Total files in SharePoint: $($sharePointFiles.Count)" -ForegroundColor White
Write-Host "Missing files (need migration): $missingCount" -ForegroundColor Red
Write-Host "Files newer on server (can migrate): $newerOnServerCount" -ForegroundColor Yellow
Write-Host "Files newer in SharePoint (will skip): $newerInSharePointCount" -ForegroundColor Green
Write-Host "Identical files: $identicalCount" -ForegroundColor Gray

# Export report
$results | Export-Csv -Path $ReportPath -NoTypeInformation
Write-Host "`nReport saved to: $ReportPath" -ForegroundColor Green

# Generate SPMT-compatible migration manifest (for files that need migration)
$spmtManifestPath = $ReportPath -replace '\.csv$', '-spmt-manifest.csv'
$spmtManifest = @()

foreach ($result in $results) {
    if ($result.Action -eq "Migrate" -or $result.Action -eq "CanMigrate") {
        # Find the full path for this file
        $fullPath = $null
        foreach ($filePath in $fileServerFiles.Keys) {
            if ($fileServerFiles[$filePath].NormalizedPath -eq $result.FilePath) {
                $fullPath = $fileServerFiles[$filePath].FullPath
                break
            }
        }
        
        if ($fullPath) {
            # SPMT format: SourcePath, TargetUrl, TargetDocumentLibrary
            # Note: SPMT typically needs full paths and handles folder structure differently
            # This is a simplified version - you may need to adjust based on your SPMT version
            $spmtManifest += [PSCustomObject]@{
                SourcePath = $fullPath
                TargetUrl = $config.SharePointSiteUrl
                TargetDocumentLibrary = $libraryName
                TargetPath = $result.SharePointPath
                FileName = Split-Path -Leaf $fullPath
            }
        }
    }
}

if ($spmtManifest.Count -gt 0) {
    $spmtManifest | Export-Csv -Path $spmtManifestPath -NoTypeInformation
    Write-Host "SPMT migration manifest saved to: $spmtManifestPath ($($spmtManifest.Count) files)" -ForegroundColor Green
    Write-Host "Note: SPMT CSV format may vary by version. Review and adjust if needed." -ForegroundColor Yellow
}
else {
    Write-Host "No files to migrate - skipping SPMT manifest generation" -ForegroundColor Gray
}

# Create summary report
$summaryPath = $ReportPath -replace '\.csv$', '-summary.txt'
$summary = @"
Migration Comparison Summary
Generated: $(Get-Date)

File Server Path: $($config.FileServerPath)
SharePoint Site: $($config.SharePointSiteUrl)

Statistics:
- Total files on server: $($fileServerFiles.Count)
- Locked/inaccessible files: $lockedCount
- Total files in SharePoint: $($sharePointFiles.Count)
- Missing files (need migration): $missingCount
- Files newer on server (can migrate): $newerOnServerCount
- Files newer in SharePoint (will skip): $newerInSharePointCount
- Identical files: $identicalCount

Next Steps:
1. Review the CSV report for detailed file information
2. Files marked as "Migrate" or "CanMigrate" can be safely migrated
3. Files marked as "Skip" should not be migrated (newer version in SharePoint)
4. Files marked as "Review" need manual inspection
5. Files marked as "Locked" or "ReviewLocked" are locked/inaccessible - close the files and re-run the script
"@

$summary | Out-File -FilePath $summaryPath
Write-Host "Summary saved to: $summaryPath" -ForegroundColor Green

# Disconnect from SharePoint
if (Get-Module -Name PnP.PowerShell) {
    Disconnect-PnPOnline
}

Write-Host "`nComparison complete!" -ForegroundColor Green
