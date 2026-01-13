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

# Function to check if a specific file exists in SharePoint and return its metadata
function Get-SharePointFile {
    param(
        [string]$SiteUrl,
        [Microsoft.SharePoint.Client.List]$List,  # Pass the list object to avoid repeated lookups
        [string]$SharePointPath  # The path as it should appear in SharePoint (e.g., "clients\client 1\my doc.pdf")
    )
    
    if (-not $List) {
        return $null
    }
    
    # Convert backslashes to forward slashes for SharePoint
    $spPath = $SharePointPath -replace '\\', '/'
    
    # Get the library root folder URL and name
    $libraryRootUrl = $List.RootFolder.ServerRelativeUrl
    $libraryRootUrl = $libraryRootUrl.TrimEnd('/')
    $libraryName = $List.Title
    
    # Construct the full server-relative URL for the file
    # SharePoint paths are typically: /sites/sitename/LibraryName/path/to/file.pdf
    $fullServerRelativeUrl = "$libraryRootUrl/$spPath"
    
    # Try to get the file directly using the constructed URL
    # Direct file access bypasses the 5000 item list view threshold
    try {
        $file = Get-PnPFile -Url $fullServerRelativeUrl -ErrorAction Stop
        if ($file) {
            # Get list item metadata using the file's server-relative URL
            # This is a targeted query by specific file URL, not a folder query, so it's safe
            $fileItem = Get-PnPListItem -List $List -Query "<View><Query><Where><Eq><FieldRef Name='FileRef'/><Value Type='Text'>$($file.ServerRelativeUrl)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($fileItem) {
                return @{
                    Name = $fileItem.FieldValues.FileLeafRef
                    Size = [long]$fileItem.FieldValues.File_x0020_Size
                    Modified = [DateTime]$fileItem.FieldValues.Modified
                    Path = $SharePointPath
                    NormalizedPath = (Normalize-Path -Path $SharePointPath)
                }
            }
        }
    }
    catch {
        # File not found at that exact path, try alternative approaches
    }
    
    # Alternative: Try with different library name prefixes
    $possiblePaths = @(
        "$libraryName/$spPath",
        "Shared Documents/$spPath",
        "Documents/$spPath",
        $spPath
    )
    
    foreach ($testPath in $possiblePaths) {
        try {
            $testUrl = "$libraryRootUrl/$testPath"
            $file = Get-PnPFile -Url $testUrl -ErrorAction Stop
            if ($file) {
                # Get list item metadata - targeted query by specific file URL (safe from 5000 limit)
                $fileItem = Get-PnPListItem -List $List -Query "<View><Query><Where><Eq><FieldRef Name='FileRef'/><Value Type='Text'>$($file.ServerRelativeUrl)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($fileItem) {
                    return @{
                        Name = $fileItem.FieldValues.FileLeafRef
                        Size = [long]$fileItem.FieldValues.File_x0020_Size
                        Modified = [DateTime]$fileItem.FieldValues.Modified
                        Path = $SharePointPath
                        NormalizedPath = (Normalize-Path -Path $SharePointPath)
                    }
                }
            }
        }
        catch {
            continue
        }
    }
    
    # Note: We don't use list queries as a fallback because they can hit the 5000 item limit
    # Direct file access (Get-PnPFile -Url) is the only reliable method for large folders
    # If the file wasn't found via direct URL, it likely doesn't exist at that path
    return $null
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
    # Example: If RootPath is "G:\Scanned Documents", rootFolderName will be "Scanned Documents"
    # This folder name will be prepended to all file paths in SharePoint
    # So "G:\Scanned Documents\folder1\file.pdf" becomes "Scanned Documents\folder1\file.pdf" in SharePoint
    $rootFolderName = Split-Path -Leaf $RootPath
    
    Write-Host "Root folder name: $rootFolderName" -ForegroundColor Gray
    Write-Host "  Files will be mapped to SharePoint path: $rootFolderName\<relative path>" -ForegroundColor Gray
    
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

# Step 1: Scan file server first (more efficient for large SharePoint sites)
Write-Host "Step 1: Scanning file server..." -ForegroundColor Cyan
$fileServerResult = Get-FileServerFiles -RootPath $config.FileServerPath -StartDate $startDate -EndDate $endDate
$fileServerFiles = $fileServerResult.Files
$lockedFiles = $fileServerResult.LockedFiles

# Generate file server scan report (before SharePoint comparison)
# This report lists all files found during the scan, useful for tracking what will be compared
Write-Host "`nGenerating file server scan report..." -ForegroundColor Cyan

# Generate report path for file server scan
$rootFolderName = Split-Path -Leaf $config.FileServerPath
$rootFolderName = $rootFolderName -replace '[<>:"/\\|?*]', '-'
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$fileServerReportPath = "fileserver-scan-$rootFolderName-$timestamp.csv"

# Create report with all scanned files
$fileServerReport = @()
foreach ($filePath in $fileServerFiles.Keys) {
    $serverFile = $fileServerFiles[$filePath]
    $fileServerReport += [PSCustomObject]@{
        FilePath = $serverFile.Path
        SharePointPath = $serverFile.SharePointPath
        FullPath = $serverFile.FullPath
        FileName = $serverFile.Name
        Size = $serverFile.Size
        Created = $serverFile.Created
        Modified = $serverFile.Modified
        IsLocked = $serverFile.IsLocked
    }
}

# Add locked files to report
foreach ($filePath in $lockedFiles.Keys) {
    $lockedFile = $lockedFiles[$filePath]
    $fileServerReport += [PSCustomObject]@{
        FilePath = $lockedFile.Path
        SharePointPath = $lockedFile.SharePointPath
        FullPath = $lockedFile.FullPath
        FileName = $lockedFile.Name
        Size = $null
        Created = $null
        Modified = $null
        IsLocked = $true
        ErrorMessage = $lockedFile.ErrorMessage
    }
}

# Export report
$fileServerReport | Export-Csv -Path $fileServerReportPath -NoTypeInformation
Write-Host "File server scan report saved to: $fileServerReportPath" -ForegroundColor Green
Write-Host "  Total files found: $($fileServerFiles.Count)" -ForegroundColor Gray
if ($lockedFiles.Count -gt 0) {
    Write-Host "  Locked/inaccessible files: $($lockedFiles.Count)" -ForegroundColor Yellow
}
if ($startDate) {
    Write-Host "  Date filter applied: Files created/modified after $($startDate.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Gray
}
if ($endDate) {
    Write-Host "  Date filter applied: Files created/modified before $($endDate.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Gray
}
if (-not $startDate -and -not $endDate) {
    Write-Host "  No date filter applied: All files included" -ForegroundColor Gray
}

# Step 2: Connect to SharePoint (only after file server scan is complete)
Write-Host "`nStep 2: Connecting to SharePoint..." -ForegroundColor Cyan
Connect-SharePoint -TenantId $config.TenantId -ClientId $config.ClientId -Thumbprint $config.Thumbprint -SiteUrl $config.SharePointSiteUrl

# Get the SharePoint library/list (cache it for efficiency)
$libraryName = if ($config.LibraryName) { $config.LibraryName } else { "Documents" }
$spList = Get-PnPList -Identity $libraryName -ErrorAction SilentlyContinue
if (-not $spList) {
    $spList = Get-PnPList -Identity "Documents" -ErrorAction SilentlyContinue
}
if (-not $spList) {
    throw "Could not find SharePoint library '$libraryName' or 'Documents'"
}

# Step 3: Compare files - check each file server file individually in SharePoint
# Note: Using direct file access (Get-PnPFile -Url) which bypasses SharePoint's 5000 item limit
# This approach queries files one at a time by direct URL, avoiding list view threshold issues
Write-Host "`nStep 3: Comparing files (checking each file in SharePoint)..." -ForegroundColor Cyan
Write-Host "Note: Using direct file access - bypasses SharePoint 5000 item limit per folder" -ForegroundColor Gray

$results = @()
$missingCount = 0
$newerOnServerCount = 0
$newerInSharePointCount = 0
$identicalCount = 0
$lockedCount = $fileServerResult.LockedCount
$checkedCount = 0
$totalFiles = $fileServerFiles.Count

foreach ($filePath in $fileServerFiles.Keys) {
    $serverFile = $fileServerFiles[$filePath]
    $checkedCount++
    
    # Check if this specific file exists in SharePoint
    # This uses Get-PnPFile -Url which directly accesses the file by URL,
    # bypassing the 5000 item list view threshold
    $spFile = Get-SharePointFile -SiteUrl $config.SharePointSiteUrl -List $spList -SharePointPath $serverFile.SharePointPath
    
    if ($spFile) {
        # File exists in SharePoint - compare modification dates
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
    
    # Progress update every 100 files
    if ($checkedCount % 100 -eq 0) {
        $percentComplete = [Math]::Round(($checkedCount / $totalFiles) * 100, 1)
        Write-Host "  Checked $checkedCount of $totalFiles files ($percentComplete%)..." -ForegroundColor Gray
    }
}

# Add locked files to results
foreach ($filePath in $lockedFiles.Keys) {
    $lockedFile = $lockedFiles[$filePath]
    
    # Check if file exists in SharePoint (even though we can't read it from server)
    $spFile = Get-SharePointFile -SiteUrl $config.SharePointSiteUrl -List $spList -SharePointPath $lockedFile.SharePointPath
    
    $results += [PSCustomObject]@{
        Status = "Locked"
        FilePath = $filePath
        SharePointPath = $lockedFile.SharePointPath
        ServerSize = $null
        ServerModified = $null
        SharePointSize = if ($spFile) { $spFile.Size } else { $null }
        SharePointModified = if ($spFile) { $spFile.Modified } else { $null }
        Action = if ($spFile) { "Review" } else { "ReviewLocked" }
        ErrorMessage = $lockedFile.ErrorMessage
    }
}

# Calculate SharePoint file count (files that exist in both)
$sharePointFileCount = $newerOnServerCount + $newerInSharePointCount + $identicalCount + ($results | Where-Object { $_.Status -eq "SizeMismatch" }).Count

# Generate report
Write-Host "`n=== Comparison Results ===" -ForegroundColor Yellow
Write-Host "Total files on server: $($fileServerFiles.Count)" -ForegroundColor White
if ($lockedCount -gt 0) {
    Write-Host "Locked/inaccessible files: $lockedCount" -ForegroundColor Yellow
}
Write-Host "Files found in SharePoint: $sharePointFileCount" -ForegroundColor White
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
- Files found in SharePoint: $sharePointFileCount
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
