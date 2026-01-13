<#
.SYNOPSIS
    One-time migration script for ETL Clients folder to SharePoint
    Hardcoded for: \\etlrom-dc1\lab\shared\clients (etc - wilco) -> SharePoint ETL/Clients
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$ConfigPath,
    
    [Parameter(Mandatory = $false)]
    [switch]$Migrate,

    # Override the file server root. By default we DO NOT take this from config to avoid surprises.
    [Parameter(Mandatory = $false)]
    [string]$SourcePath = "\\etlrom-dc1\lab\shared\clients (etc - wilco)",

    # Override the library-relative base path. By default we DO NOT take this from config.
    [Parameter(Mandatory = $false)]
    [string]$SharePointBasePath = "ETL",

    # How many planned mappings to show + log early (helps validate destination quickly)
    [Parameter(Mandatory = $false)]
    [int]$PreviewCount = 25,

    # Optional date filter: only consider files created/modified on/after StartDate (local time)
    [Parameter(Mandatory = $false)]
    [DateTime]$StartDate,

    # Optional date filter: only consider files created/modified on/before EndDate (local time)
    [Parameter(Mandatory = $false)]
    [DateTime]$EndDate,

    # Safety: stop after N upload attempts (success or fail). 0 means no limit.
    [Parameter(Mandatory = $false)]
    [int]$StopAfter = 0,

    # Safety: stop immediately on first upload failure
    [Parameter(Mandatory = $false)]
    [switch]$FailFast,

    # Include files that already exist in SharePoint but are newer on the file server (overwrite SharePoint copy)
    [Parameter(Mandatory = $false)]
    [switch]$IncludeCanMigrate,

    # Log file (JSONL). Defaults to a timestamped file under .\logs
    [Parameter(Mandatory = $false)]
    [string]$LogPath
)

# Error handling
$ErrorActionPreference = "Stop"

# Function to sanitize a library-relative path (do not allow "Documents"/"Shared Documents" prefixes)
function Sanitize-LibraryRelativePath {
    param([string]$Path)

    if (-not $Path) {
        return $Path
    }

    $p = $Path.Trim()
    $p = $p -replace '\\', '/'
    $p = $p.Trim('/')

    if (-not $p) {
        return $p
    }

    $parts = $p -split '/'
    $filtered = @()
    foreach ($part in $parts) {
        if (-not $part) { continue }
        $filtered += $part
    }

    # If someone accidentally included the library name in the "base path", remove it.
    while ($filtered.Count -gt 0 -and ($filtered[0] -ieq 'documents' -or $filtered[0] -ieq 'shared documents')) {
        $filtered = $filtered[1..($filtered.Count - 1)]
    }

    return ($filtered -join '/')
}

# Ensure we have a log path early
if (-not $LogPath) {
    $LogPath = Join-Path -Path "." -ChildPath ("logs\etl-clients-migration-{0}.jsonl" -f (Get-Date -Format "yyyyMMdd-HHmmss"))
}
$logDir = Split-Path -Parent $LogPath
if ($logDir -and -not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

# Load configuration
Write-Host "Loading configuration from $ConfigPath..." -ForegroundColor Cyan
if (-not (Test-Path $ConfigPath)) {
    throw "Configuration file not found: $ConfigPath"
}

$config = Get-Content $ConfigPath | ConvertFrom-Json

# Paths: intentionally ignore config-provided paths to avoid accidentally nesting under "Documents"/"Shared Documents"
$sourcePath = $SourcePath
$targetSharePointBasePath = Sanitize-LibraryRelativePath -Path $SharePointBasePath
$targetRootFolderName = "Clients"  # Always transform root folder to "Clients"

# Validate configuration
$requiredFields = @('TenantId', 'ClientId', 'Thumbprint', 'SharePointSiteUrl')
foreach ($field in $requiredFields) {
    if (-not $config.$field) {
        throw "Missing required configuration field: $field"
    }
}

# Early check: Verify source path is accessible
Write-Host "Verifying source path is accessible..." -ForegroundColor Cyan
if (-not (Test-Path $sourcePath)) {
    throw "Source path is not accessible: $sourcePath. Please verify the path exists and you have access permissions."
}
Write-Host "Source path is accessible: $sourcePath" -ForegroundColor Green

# Early check: Verify SharePoint connection and access
Write-Host "Verifying SharePoint connection..." -ForegroundColor Cyan
try {
    # Check if PnP.PowerShell is available
    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
        Write-Host "PnP.PowerShell module not found. Installing..." -ForegroundColor Yellow
        Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module PnP.PowerShell -ErrorAction Stop
    
    # Verify certificate exists
    Write-Host "Checking for certificate..." -ForegroundColor Gray
    $cert = Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object { $_.Thumbprint -eq $config.Thumbprint }
    if (-not $cert) {
        $cert = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Thumbprint -eq $config.Thumbprint }
    }
    if (-not $cert) {
        throw "Certificate with thumbprint '$($config.Thumbprint)' not found in certificate store."
    }
    
    # Test connection to SharePoint
    Connect-PnPOnline -Url $config.SharePointSiteUrl -ClientId $config.ClientId -Thumbprint $config.Thumbprint -Tenant $config.TenantId -WarningAction SilentlyContinue -ErrorAction Stop
    
    # Verify we can access the Documents library
    $testList = Get-PnPList -Identity "Documents" -ErrorAction Stop
    if (-not $testList) {
        throw "Could not access 'Documents' library in SharePoint. Please verify permissions."
    }
    
    # Test reading a file/folder to ensure we have proper access
    $testFolder = Get-PnPFolder -Url $testList.RootFolder.ServerRelativeUrl -ErrorAction Stop
    if (-not $testFolder) {
        throw "Could not access SharePoint root folder. Please verify permissions."
    }
    
    Write-Host "SharePoint connection verified: $($config.SharePointSiteUrl)" -ForegroundColor Green
    Write-Host "  Documents library accessible" -ForegroundColor Gray
    
    # Disconnect for now (will reconnect later)
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}
catch {
    throw "Failed to verify SharePoint access: $_`nPlease check your SharePoint site URL, certificate, and permissions."
}

# Function to connect to SharePoint
function Connect-SharePoint {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$Thumbprint,
        [string]$SiteUrl
    )
    
    Write-Host "Connecting to SharePoint..." -ForegroundColor Cyan
    
    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
        Write-Host "PnP.PowerShell module not found. Installing..." -ForegroundColor Yellow
        Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
    }
    
    Import-Module PnP.PowerShell -ErrorAction Stop
    
    # Verify certificate exists
    Write-Host "Checking for certificate with thumbprint: $Thumbprint..." -ForegroundColor Gray
    $cert = Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object { $_.Thumbprint -eq $Thumbprint }
    if (-not $cert) {
        $cert = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Thumbprint -eq $Thumbprint }
    }
    
    if (-not $cert) {
        throw "Certificate with thumbprint '$Thumbprint' not found in certificate store."
    }
    
    Write-Host "Certificate found: $($cert.Subject) (Valid until: $($cert.NotAfter))" -ForegroundColor Gray
    
    try {
        Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $TenantId -WarningAction SilentlyContinue -ErrorAction Stop
        $web = Get-PnPWeb
        Write-Host "Connected to: $($web.Title)" -ForegroundColor Green
        return $true
    }
    catch {
        throw "Failed to connect with PnP.PowerShell: $_"
    }
}

# Function to normalize paths
function Normalize-Path {
    param([string]$Path)
    $normalized = $Path.ToLower()
    $normalized = $normalized -replace '/', '\'
    $normalized = $normalized.TrimStart('\')
    return $normalized
}

# Structured logger (JSONL). Safe for long runs and easy to tail/grep.
function Write-MigrationLog {
    param(
        [string]$Level,
        [string]$Message,
        [hashtable]$Data
    )

    $evt = [ordered]@{
        ts      = (Get-Date).ToString("o")
        level   = $Level
        message = $Message
    }
    if ($Data) {
        foreach ($k in $Data.Keys) {
            $evt[$k] = $Data[$k]
        }
    }

    ($evt | ConvertTo-Json -Compress) | Add-Content -Path $LogPath -Encoding UTF8
}

# Function to ensure folder exists in SharePoint
function Ensure-SharePointFolder {
    param(
        [Microsoft.SharePoint.Client.List]$List,
        [string]$FolderPath
    )
    
    if (-not $List -or -not $FolderPath) {
        return $false
    }
    
    $spFolderPath = $FolderPath -replace '\\', '/'
    $libraryRootUrl = $List.RootFolder.ServerRelativeUrl.TrimEnd('/')
    $fullFolderPath = "$libraryRootUrl/$spFolderPath"
    
    try {
        $folder = Get-PnPFolder -Url $fullFolderPath -ErrorAction Stop
        if ($folder) {
            return $true
        }
    }
    catch {
        # Folder doesn't exist, create it
        $pathParts = $spFolderPath -split '/'
        $currentPath = $libraryRootUrl
        
        foreach ($part in $pathParts) {
            if ($part) {
                $currentPath = "$currentPath/$part"
                try {
                    $folder = Get-PnPFolder -Url $currentPath -ErrorAction Stop
                }
                catch {
                    try {
                        $parentPath = $currentPath.Substring(0, $currentPath.LastIndexOf('/'))
                        Add-PnPFolder -Name $part -Folder $parentPath -ErrorAction Stop
                        Write-Host "      Created folder: $currentPath" -ForegroundColor Gray
                    }
                    catch {
                        Write-Host "      Warning: Failed to create folder $currentPath - $_" -ForegroundColor Yellow
                        return $false
                    }
                }
            }
        }
        return $true
    }
    
    return $false
}

# Function to upload file to SharePoint
function Copy-FileToSharePoint {
    param(
        [string]$SourcePath,
        [string]$SharePointPath,
        [Microsoft.SharePoint.Client.List]$List,
        [switch]$Overwrite
    )
    
    if (-not (Test-Path $SourcePath)) {
        return @{ Success = $false; Error = "Source file not found: $SourcePath" }
    }
    
    try {
        $spPath = $SharePointPath -replace '\\', '/'
        $fileName = Split-Path -Leaf $spPath
        $folderPath = $spPath.Substring(0, $spPath.LastIndexOf('/'))
        
        if ($folderPath) {
            $folderCreated = Ensure-SharePointFolder -List $List -FolderPath ($folderPath -replace '/', '\')
            if (-not $folderCreated) {
                return @{ Success = $false; Error = "Failed to create folder structure: $folderPath" }
            }
        }
        
        $libraryRootUrl = $List.RootFolder.ServerRelativeUrl.TrimEnd('/')
        
        if ($folderPath) {
            $targetFolderUrl = "$libraryRootUrl/$folderPath"
        }
        else {
            $targetFolderUrl = $libraryRootUrl
        }
        
        if ($Overwrite) {
            $file = Add-PnPFile -Path $SourcePath -Folder $targetFolderUrl -Overwrite -ErrorAction Stop
        }
        else {
            $file = Add-PnPFile -Path $SourcePath -Folder $targetFolderUrl -ErrorAction Stop
        }
        
        if ($file) {
            return @{ Success = $true; FileUrl = $file.ServerRelativeUrl; Error = $null }
        }
        else {
            return @{ Success = $false; Error = "Upload completed but file object not returned" }
        }
    }
    catch {
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

# Main script
Write-Host "`n=== ETL Clients Migration Script ===" -ForegroundColor Yellow
Write-Host "Source: $sourcePath" -ForegroundColor Cyan
Write-Host "Target: $($config.SharePointSiteUrl) (Documents library) -> $targetSharePointBasePath/$targetRootFolderName" -ForegroundColor Cyan
Write-Host "Log: $LogPath" -ForegroundColor Gray
Write-Host ""

Write-MigrationLog -Level "INFO" -Message "Starting migration run" -Data @{
    SourcePath            = $sourcePath
    SharePointSiteUrl     = $config.SharePointSiteUrl
    TargetLibraryIdentity = "Documents"
    TargetBasePath        = "$targetSharePointBasePath/$targetRootFolderName"
    Migrate               = [bool]$Migrate
    PreviewCount          = $PreviewCount
    StartDate             = if ($StartDate) { $StartDate.ToString("o") } else { $null }
    EndDate               = if ($EndDate) { $EndDate.ToString("o") } else { $null }
    StopAfter             = $StopAfter
    FailFast              = [bool]$FailFast
    IncludeCanMigrate     = [bool]$IncludeCanMigrate
}

# Step 1: Scan source folder
Write-Host "Step 1: Scanning source folder..." -ForegroundColor Cyan
$files = @()
$fileCount = 0
$skippedByDateCount = 0

Get-ChildItem -Path $sourcePath -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
    $fileCount++

    # Optional date filter: include a file if its Created OR Modified falls within the range.
    if ($StartDate -or $EndDate) {
        $created = $_.CreationTime
        $modified = $_.LastWriteTime

        if ($StartDate -and ($created -lt $StartDate) -and ($modified -lt $StartDate)) {
            $skippedByDateCount++
            if ($fileCount % 1000 -eq 0) {
                Write-Host "  Enumerated $fileCount files... kept $($files.Count), skipped by date $skippedByDateCount" -ForegroundColor Gray
            }
            return
        }
        if ($EndDate -and ($created -gt $EndDate) -and ($modified -gt $EndDate)) {
            $skippedByDateCount++
            if ($fileCount % 1000 -eq 0) {
                Write-Host "  Enumerated $fileCount files... kept $($files.Count), skipped by date $skippedByDateCount" -ForegroundColor Gray
            }
            return
        }
    }

    $relativePath = $_.FullName -replace [regex]::Escape($sourcePath), ""
    $relativePath = $relativePath.TrimStart('\', '/')
    
    # Build SharePoint path: ETL/Clients/[relative path]
    # Only transform the root folder name, keep subfolders as-is
    $sharePointPath = "$targetSharePointBasePath\$targetRootFolderName\$relativePath"
    $sharePointPath = Sanitize-LibraryRelativePath -Path $sharePointPath
    
    $files += @{
        FullPath = $_.FullName
        RelativePath = $relativePath
        SharePointPath = $sharePointPath
        Name = $_.Name
        Size = $_.Length
        Modified = $_.LastWriteTime
    }
    
    if ($fileCount % 1000 -eq 0) {
        if ($StartDate -or $EndDate) {
            Write-Host "  Enumerated $fileCount files... kept $($files.Count), skipped by date $skippedByDateCount" -ForegroundColor Gray
        }
        else {
            Write-Host "  Scanned $fileCount files..." -ForegroundColor Gray
        }
    }
}

Write-Host "Found $($files.Count) files to process" -ForegroundColor Green
if ($StartDate -or $EndDate) {
    Write-Host "Date filter active. Skipped by date: $skippedByDateCount" -ForegroundColor Gray
}
Write-MigrationLog -Level "INFO" -Message "Source scan complete" -Data @{
    TotalEnumerated = $fileCount
    FileCount = $files.Count
    SkippedByDate = $skippedByDateCount
    StartDate = if ($StartDate) { $StartDate.ToString("o") } else { $null }
    EndDate = if ($EndDate) { $EndDate.ToString("o") } else { $null }
}

# Step 2: Connect to SharePoint
Write-Host "`nStep 2: Connecting to SharePoint..." -ForegroundColor Cyan
Connect-SharePoint -TenantId $config.TenantId -ClientId $config.ClientId -Thumbprint $config.Thumbprint -SiteUrl $config.SharePointSiteUrl

# Get Documents library
$spList = Get-PnPList -Identity "Documents" -ErrorAction SilentlyContinue
if (-not $spList) {
    throw "Could not find 'Documents' library in SharePoint"
}

# Print where we're actually writing (helps diagnose accidental "Shared Documents" folder creation)
$libraryRootUrl = $spList.RootFolder.ServerRelativeUrl.TrimEnd('/')
Write-Host "Documents library detected:" -ForegroundColor Gray
Write-Host "  Title: $($spList.Title)" -ForegroundColor Gray
Write-Host "  RootFolder.ServerRelativeUrl: $libraryRootUrl" -ForegroundColor Gray
Write-Host "  Upload base (library-relative): $targetSharePointBasePath/$targetRootFolderName" -ForegroundColor Gray
Write-MigrationLog -Level "INFO" -Message "Connected to Documents library" -Data @{
    LibraryTitle  = $spList.Title
    LibraryRootUrl = $libraryRootUrl
    UploadBase     = "$targetSharePointBasePath/$targetRootFolderName"
}

# Step 3: Compare files (optional - check what exists)
Write-Host "`nStep 3: Checking existing files in SharePoint..." -ForegroundColor Cyan
$filesToMigrate = @()
$existingCount = 0
$canMigrateCount = 0
$skippedNewerInSharePointCount = 0
$checkCount = 0

foreach ($file in $files) {
    $checkCount++
    if ($checkCount % 100 -eq 0) {
        Write-Host "  Checking files: $checkCount of $($files.Count)..." -ForegroundColor Gray
        [Console]::Out.Flush()
    }
    
    $spPath = $file.SharePointPath -replace '\\', '/'
    $libraryRootUrl = $spList.RootFolder.ServerRelativeUrl.TrimEnd('/')
    $checkUrl = "$libraryRootUrl/$spPath"
    
    try {
        # Use -AsListItem so we can compare Modified timestamps (supports "CanMigrate" logic)
        $existingItem = Get-PnPFile -Url $checkUrl -AsListItem -ErrorAction SilentlyContinue

        if ($existingItem) {
            $existingCount++

            $spModified = $null
            try { $spModified = [DateTime]$existingItem.FieldValues.Modified } catch { $spModified = $null }

            if ($spModified -and $file.Modified -gt $spModified) {
                # Newer on server => CanMigrate (overwrite) IF user opts in
                $canMigrateCount++
                if ($IncludeCanMigrate) {
                    $fileToAdd = $file.Clone()
                    $fileToAdd["Status"] = "CanMigrate"
                    $fileToAdd["SharePointModified"] = $spModified
                    $filesToMigrate += $fileToAdd
                    if ($PreviewCount -gt 0 -and $filesToMigrate.Count -le $PreviewCount) {
                        Write-Host ("  PLAN [{0}] (CanMigrate) {1} -> {2}" -f $filesToMigrate.Count, $file.RelativePath, $checkUrl) -ForegroundColor DarkCyan
                        Write-MigrationLog -Level "PLAN" -Message "File can be migrated (newer on server)" -Data @{
                            Status            = "CanMigrate"
                            SourceFile        = $file.FullPath
                            RelativePath      = $file.RelativePath
                            TargetUrl         = $checkUrl
                            ServerModified    = $file.Modified
                            SharePointModified = $spModified
                        }
                    }
                }
            }
            elseif ($spModified -and $spModified -gt $file.Modified) {
                # Newer in SharePoint => skip to avoid overwriting
                $skippedNewerInSharePointCount++
            }
            else {
                # Same modified time (or couldn't read it) => treat as existing
            }
        }
        else {
            # Missing in SharePoint => migrate
            $fileToAdd = $file.Clone()
            $fileToAdd["Status"] = "Migrate"
            $filesToMigrate += $fileToAdd
            if ($PreviewCount -gt 0 -and $filesToMigrate.Count -le $PreviewCount) {
                Write-Host ("  PLAN [{0}] (Migrate) {1} -> {2}" -f $filesToMigrate.Count, $file.RelativePath, $checkUrl) -ForegroundColor DarkCyan
                Write-MigrationLog -Level "PLAN" -Message "File needs migration (missing)" -Data @{
                    Status       = "Migrate"
                    SourceFile   = $file.FullPath
                    RelativePath = $file.RelativePath
                    TargetUrl    = $checkUrl
                }
            }
        }
    }
    catch {
        # File doesn't exist, add to migration list
        $fileToAdd = $file.Clone()
        $fileToAdd["Status"] = "Migrate"
        $filesToMigrate += $fileToAdd
        if ($PreviewCount -gt 0 -and $filesToMigrate.Count -le $PreviewCount) {
            Write-Host ("  PLAN [{0}] (Migrate) {1} -> {2}" -f $filesToMigrate.Count, $file.RelativePath, $checkUrl) -ForegroundColor DarkCyan
            Write-MigrationLog -Level "PLAN" -Message "File needs migration (missing)" -Data @{
                Status       = "Migrate"
                SourceFile   = $file.FullPath
                RelativePath = $file.RelativePath
                TargetUrl    = $checkUrl
            }
        }
    }
}

Write-Host "Files already in SharePoint: $existingCount" -ForegroundColor Gray
Write-Host "Files newer on server (CanMigrate): $canMigrateCount" -ForegroundColor Gray
Write-Host "Skipped (newer in SharePoint): $skippedNewerInSharePointCount" -ForegroundColor Gray
Write-Host "Files to migrate: $($filesToMigrate.Count)" -ForegroundColor Cyan
Write-MigrationLog -Level "INFO" -Message "Plan computed" -Data @{
    ExistingCount    = $existingCount
    ToMigrateCount   = $filesToMigrate.Count
    PreviewLogged    = [Math]::Min($PreviewCount, $filesToMigrate.Count)
    CanMigrateCount  = $canMigrateCount
    SkippedNewerInSharePointCount = $skippedNewerInSharePointCount
    IncludeCanMigrate = [bool]$IncludeCanMigrate
}

# Step 4: Migrate files (if -Migrate is specified)
if ($Migrate) {
    Write-Host "`nStep 4: Migrating files to SharePoint..." -ForegroundColor Yellow
    
    if ($filesToMigrate.Count -eq 0) {
        Write-Host "No files to migrate." -ForegroundColor Gray
    }
    else {
        $migratedCount = 0
        $failedCount = 0
        $currentFile = 0
        $attemptedCount = 0
        
        foreach ($file in $filesToMigrate) {
            $currentFile++
            $percent = [Math]::Round(($currentFile / $filesToMigrate.Count) * 100, 1)
            
            Write-Host "  [$currentFile/$($filesToMigrate.Count)] ($percent%) Uploading: $($file.RelativePath)" -ForegroundColor Cyan
            Write-Host "    Starting upload..." -ForegroundColor Gray
            
            # Flush output buffer to ensure message appears immediately
            [Console]::Out.Flush()
            
            $attemptedCount++
            $shouldOverwrite = ($file.Status -eq "CanMigrate")
            Write-MigrationLog -Level "UPLOAD_START" -Message "Uploading file" -Data @{
                Index        = $currentFile
                Total        = $filesToMigrate.Count
                SourceFile   = $file.FullPath
                RelativePath = $file.RelativePath
                SharePointPath = $file.SharePointPath
                Status       = $file.Status
                Overwrite    = [bool]$shouldOverwrite
            }
            $result = Copy-FileToSharePoint -SourcePath $file.FullPath -SharePointPath $file.SharePointPath -List $spList -Overwrite:$shouldOverwrite
            
            if ($result.Success) {
                Write-Host "    ✓ Successfully uploaded" -ForegroundColor Green
                $migratedCount++
                Write-MigrationLog -Level "UPLOAD_OK" -Message "Upload succeeded" -Data @{
                    Index        = $currentFile
                    SourceFile   = $file.FullPath
                    TargetUrl    = $result.FileUrl
                }
            }
            else {
                Write-Host "    ✗ Failed: $($result.Error)" -ForegroundColor Red
                $failedCount++
                Write-MigrationLog -Level "UPLOAD_FAIL" -Message "Upload failed" -Data @{
                    Index      = $currentFile
                    SourceFile = $file.FullPath
                    Error      = $result.Error
                }
                if ($FailFast) {
                    throw "FailFast: stopping on first upload failure: $($result.Error)"
                }
            }
            
            # Flush after each file to ensure progress is visible
            [Console]::Out.Flush()

            if ($StopAfter -gt 0 -and $attemptedCount -ge $StopAfter) {
                Write-Host "StopAfter reached ($StopAfter upload attempts). Stopping early." -ForegroundColor Yellow
                Write-MigrationLog -Level "WARN" -Message "StopAfter reached; stopping early" -Data @{
                    StopAfter      = $StopAfter
                    AttemptedCount = $attemptedCount
                    MigratedCount  = $migratedCount
                    FailedCount    = $failedCount
                }
                break
            }
        }
        
        Write-Host "`n=== Migration Results ===" -ForegroundColor Yellow
        Write-Host "Successfully migrated: $migratedCount" -ForegroundColor Green
        Write-Host "Failed: $failedCount" -ForegroundColor Red
        Write-MigrationLog -Level "INFO" -Message "Migration results" -Data @{
            MigratedCount = $migratedCount
            FailedCount   = $failedCount
        }
    }
}
else {
    Write-Host "`nNote: Use -Migrate parameter to actually upload files" -ForegroundColor Yellow
    Write-Host "Example: .\Migrate-ETL-Clients.ps1 -ConfigPath config.json -Migrate" -ForegroundColor Gray
    Write-Host ("Tip: run a tiny validation first: -Migrate -StopAfter 1 (or 5)") -ForegroundColor Gray
}

# Disconnect
if (Get-Module -Name PnP.PowerShell) {
    Disconnect-PnPOnline
}

Write-Host "`nScript complete!" -ForegroundColor Green
