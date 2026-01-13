<#
.SYNOPSIS
    Streaming migration script for ETL Clients folder to SharePoint (Documents library).

.DESCRIPTION
    This script is intentionally modeled after Compare-Migration.ps1 streaming behavior:
    it enumerates the file server recursively, and for each file it immediately checks SharePoint
    and (if -Migrate is set) uploads right away.

    Hardcoded destination mapping:
      <file-server-root>\<relative path>  ->  SharePoint Documents library: ETL/Clients/<relative path>

    Notes:
    - This still enumerates the whole tree; date filters reduce "kept" files but enumeration is required.
    - Uses timestamp tolerance to avoid false "CanMigrate" classifications due to rounding/milliseconds.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$ConfigPath,

    [Parameter(Mandatory = $false)]
    [switch]$Migrate,

    # Hardcoded ETL defaults (override only if you explicitly want to)
    [Parameter(Mandatory = $false)]
    [string]$SourcePath = "\\etlrom-dc1\lab\shared\clients (etc - wilco)",

    [Parameter(Mandatory = $false)]
    [string]$TargetBasePath = "ETL",

    # Always uses Clients as the root folder under TargetBasePath
    [Parameter(Mandatory = $false)]
    [string]$TargetRootFolderName = "Clients",

    # Include files that already exist in SharePoint but are newer on the file server (overwrite SharePoint copy)
    [Parameter(Mandatory = $false)]
    [switch]$IncludeCanMigrate,

    # Optional date filter: include if Created OR Modified is within range (local time)
    [Parameter(Mandatory = $false)]
    [Nullable[DateTime]]$StartDate = $null,

    [Parameter(Mandatory = $false)]
    [Nullable[DateTime]]$EndDate = $null,

    # SharePoint + Windows timestamps can differ by milliseconds / timezone kind while printing the same.
    [Parameter(Mandatory = $false)]
    [double]$TimestampToleranceSeconds = 2,

    # Show/log the first N planned actions (helps validate destination quickly)
    [Parameter(Mandatory = $false)]
    [int]$PreviewCount = 25,

    # Safety: stop after N upload attempts (success or fail). 0 means no limit.
    [Parameter(Mandatory = $false)]
    [int]$StopAfter = 0,

    # Safety: stop immediately on first upload failure
    [Parameter(Mandatory = $false)]
    [switch]$FailFast,

    # Log file (JSONL). Defaults to a timestamped file under .\logs
    [Parameter(Mandatory = $false)]
    [string]$LogPath
)

# Error handling
$ErrorActionPreference = "Stop"

function Sanitize-LibraryRelativePath {
    param([string]$Path)

    if (-not $Path) { return $Path }

    $p = $Path.Trim()
    $p = $p -replace '\\', '/'
    $p = $p.Trim('/')
    if (-not $p) { return $p }

    $parts = $p -split '/'
    $filtered = @()
    foreach ($part in $parts) {
        if (-not $part) { continue }
        $filtered += $part
    }

    # Remove accidental library names if present
    while ($filtered.Count -gt 0 -and ($filtered[0] -ieq 'documents' -or $filtered[0] -ieq 'shared documents')) {
        $filtered = $filtered[1..($filtered.Count - 1)]
    }

    return ($filtered -join '/')
}

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
        foreach ($k in $Data.Keys) { $evt[$k] = $Data[$k] }
    }

    ($evt | ConvertTo-Json -Compress) | Add-Content -Path $LogPath -Encoding UTF8
}

function Connect-SharePoint {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$Thumbprint,
        [string]$SiteUrl
    )

    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
        Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module PnP.PowerShell -ErrorAction Stop

    # Verify certificate exists
    $cert = Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object { $_.Thumbprint -eq $Thumbprint }
    if (-not $cert) {
        $cert = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Thumbprint -eq $Thumbprint }
    }
    if (-not $cert) {
        throw "Certificate with thumbprint '$Thumbprint' not found in certificate store."
    }

    Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $TenantId -WarningAction SilentlyContinue -ErrorAction Stop
}

function Ensure-SharePointFolder {
    param(
        [Microsoft.SharePoint.Client.List]$List,
        [string]$FolderPath
    )

    if (-not $List -or -not $FolderPath) { return $false }

    $spFolderPath = $FolderPath -replace '\\', '/'
    $libraryRootUrl = $List.RootFolder.ServerRelativeUrl.TrimEnd('/')
    $fullFolderPath = "$libraryRootUrl/$spFolderPath"

    try {
        $folder = Get-PnPFolder -Url $fullFolderPath -ErrorAction Stop
        return [bool]$folder
    }
    catch {
        $pathParts = $spFolderPath -split '/'
        $currentPath = $libraryRootUrl

        foreach ($part in $pathParts) {
            if (-not $part) { continue }
            $currentPath = "$currentPath/$part"
            try {
                $null = Get-PnPFolder -Url $currentPath -ErrorAction Stop
            }
            catch {
                $parentPath = $currentPath.Substring(0, $currentPath.LastIndexOf('/'))
                Add-PnPFolder -Name $part -Folder $parentPath -ErrorAction Stop | Out-Null
            }
        }
        return $true
    }
}

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
        $folderPath = $spPath.Substring(0, $spPath.LastIndexOf('/'))

        if ($folderPath) {
            $ok = Ensure-SharePointFolder -List $List -FolderPath ($folderPath -replace '/', '\')
            if (-not $ok) {
                return @{ Success = $false; Error = "Failed to create folder structure: $folderPath" }
            }
        }

        $libraryRootUrl = $List.RootFolder.ServerRelativeUrl.TrimEnd('/')
        $targetFolderUrl = if ($folderPath) { "$libraryRootUrl/$folderPath" } else { $libraryRootUrl }

        if ($Overwrite) {
            $file = Add-PnPFile -Path $SourcePath -Folder $targetFolderUrl -Overwrite -ErrorAction Stop
        }
        else {
            $file = Add-PnPFile -Path $SourcePath -Folder $targetFolderUrl -ErrorAction Stop
        }

        if ($file) {
            return @{ Success = $true; FileUrl = $file.ServerRelativeUrl; Error = $null }
        }
        return @{ Success = $false; Error = "Upload completed but file object not returned" }
    }
    catch {
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

Write-Host "`n=== Stream ETL Clients Migration ===" -ForegroundColor Yellow

if (-not $LogPath) {
    $LogPath = Join-Path -Path "." -ChildPath ("logs\etl-clients-stream-{0}.jsonl" -f (Get-Date -Format "yyyyMMdd-HHmmss"))
}
$logDir = Split-Path -Parent $LogPath
if ($logDir -and -not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

if (-not (Test-Path $ConfigPath)) {
    throw "Configuration file not found: $ConfigPath"
}
$config = Get-Content $ConfigPath | ConvertFrom-Json

# Validate configuration
$requiredFields = @('TenantId', 'ClientId', 'Thumbprint', 'SharePointSiteUrl')
foreach ($field in $requiredFields) {
    if (-not $config.$field) { throw "Missing required configuration field: $field" }
}

# Date filters: if not explicitly provided, fall back to config values
if (-not $PSBoundParameters.ContainsKey('StartDate') -and $config.StartDate) {
    try { $StartDate = [DateTime]::Parse($config.StartDate) } catch { throw "Invalid StartDate format in config. Use format like '2024-01-01' or '2024-01-01 00:00:00'" }
}
if (-not $PSBoundParameters.ContainsKey('EndDate') -and $config.EndDate) {
    try { $EndDate = [DateTime]::Parse($config.EndDate) } catch { throw "Invalid EndDate format in config. Use format like '2024-12-31' or '2024-12-31 23:59:59'" }
}

$targetBase = Sanitize-LibraryRelativePath -Path $TargetBasePath
$targetRoot = $TargetRootFolderName

Write-Host "Source: $SourcePath" -ForegroundColor Cyan
Write-Host "Target: $($config.SharePointSiteUrl) (Documents library) -> $targetBase/$targetRoot" -ForegroundColor Cyan
Write-Host "Log: $LogPath" -ForegroundColor Gray
if ($StartDate -or $EndDate) {
    Write-Host "Date filter active: StartDate=$StartDate EndDate=$EndDate" -ForegroundColor Gray
}
Write-Host ""

Write-MigrationLog -Level "INFO" -Message "Starting streaming migration run" -Data @{
    SourcePath = $SourcePath
    SharePointSiteUrl = $config.SharePointSiteUrl
    TargetLibraryIdentity = "Documents"
    TargetBasePath = "$targetBase/$targetRoot"
    Migrate = [bool]$Migrate
    IncludeCanMigrate = [bool]$IncludeCanMigrate
    StartDate = if ($StartDate) { ([DateTime]$StartDate).ToString("o") } else { $null }
    EndDate = if ($EndDate) { ([DateTime]$EndDate).ToString("o") } else { $null }
    TimestampToleranceSeconds = $TimestampToleranceSeconds
    PreviewCount = $PreviewCount
    StopAfter = $StopAfter
    FailFast = [bool]$FailFast
}

Write-Host "Connecting to SharePoint..." -ForegroundColor Cyan
Connect-SharePoint -TenantId $config.TenantId -ClientId $config.ClientId -Thumbprint $config.Thumbprint -SiteUrl $config.SharePointSiteUrl
$spList = Get-PnPList -Identity "Documents" -ErrorAction Stop
$libraryRootUrl = $spList.RootFolder.ServerRelativeUrl.TrimEnd('/')

Write-Host "Documents library detected:" -ForegroundColor Gray
Write-Host "  Title: $($spList.Title)" -ForegroundColor Gray
Write-Host "  RootFolder.ServerRelativeUrl: $libraryRootUrl" -ForegroundColor Gray
Write-Host ""

Write-MigrationLog -Level "INFO" -Message "Connected to Documents library" -Data @{
    LibraryTitle = $spList.Title
    LibraryRootUrl = $libraryRootUrl
}

if (-not (Test-Path $SourcePath)) {
    throw "Source path is not accessible: $SourcePath"
}

$enumerated = 0
$kept = 0
$filteredByDate = 0
$plannedLogged = 0

$missingCount = 0
$canMigrateCount = 0
$skippedNewerInSharePointCount = 0
$skippedExistingCount = 0

$attemptedUploads = 0
$uploadedOk = 0
$uploadedFail = 0

$start = if ($StartDate) { [DateTime]$StartDate } else { $null }
$end = if ($EndDate) { [DateTime]$EndDate } else { $null }

Write-Host "Scanning file server (streaming)..." -ForegroundColor Cyan

Get-ChildItem -Path $SourcePath -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
    $enumerated++

    if ($start -or $end) {
        $created = $_.CreationTime
        $modified = $_.LastWriteTime

        if ($start -and ($created -lt $start) -and ($modified -lt $start)) {
            $filteredByDate++
            if ($enumerated % 1000 -eq 0) {
                Write-Host "  Enumerated $enumerated... kept $kept, filtered by date $filteredByDate" -ForegroundColor Gray
            }
            return
        }
        if ($end -and ($created -gt $end) -and ($modified -gt $end)) {
            $filteredByDate++
            if ($enumerated % 1000 -eq 0) {
                Write-Host "  Enumerated $enumerated... kept $kept, filtered by date $filteredByDate" -ForegroundColor Gray
            }
            return
        }
    }

    $kept++

    # Map to SharePoint path: ETL/Clients/<relative path>
    $relativePath = $_.FullName -replace [regex]::Escape($SourcePath), ""
    $relativePath = $relativePath.TrimStart('\', '/')
    $sharePointPath = "$targetBase\$targetRoot\$relativePath"
    $sharePointPath = Sanitize-LibraryRelativePath -Path $sharePointPath

    $spPath = $sharePointPath -replace '\\', '/'
    $checkUrl = "$libraryRootUrl/$spPath"

    # Check SharePoint (direct URL) for metadata
    $existingItem = $null
    try { $existingItem = Get-PnPFile -Url $checkUrl -AsListItem -ErrorAction SilentlyContinue } catch { $existingItem = $null }

    $status = $null
    $action = "Skip"
    $overwrite = $false
    $spModified = $null
    $deltaSeconds = $null

    if (-not $existingItem) {
        $status = "Missing"
        $action = "Migrate"
        $missingCount++
    }
    else {
        try { $spModified = [DateTime]$existingItem.FieldValues.Modified } catch { $spModified = $null }
        if ($spModified) {
            $deltaSeconds = ($_.LastWriteTime - $spModified).TotalSeconds
            if ($deltaSeconds -gt $TimestampToleranceSeconds) {
                $status = "NewerOnServer"
                $canMigrateCount++
                if ($IncludeCanMigrate) {
                    $action = "CanMigrate"
                    $overwrite = $true
                }
                else {
                    $action = "Skip"
                }
            }
            elseif ($deltaSeconds -lt (-1 * $TimestampToleranceSeconds)) {
                $status = "NewerInSharePoint"
                $skippedNewerInSharePointCount++
                $action = "Skip"
            }
            else {
                $status = "SameTimestamp"
                $skippedExistingCount++
                $action = "Skip"
            }
        }
        else {
            $status = "Exists"
            $skippedExistingCount++
            $action = "Skip"
        }
    }

    if ($plannedLogged -lt $PreviewCount -and ($action -eq "Migrate" -or $action -eq "CanMigrate")) {
        $plannedLogged++
        Write-Host ("  PLAN [{0}] ({1}) {2} -> {3}" -f $plannedLogged, $action, $relativePath, $checkUrl) -ForegroundColor DarkCyan
        Write-MigrationLog -Level "PLAN" -Message "Planned action" -Data @{
            Action = $action
            Status = $status
            SourceFile = $_.FullName
            RelativePath = $relativePath
            SharePointPath = $sharePointPath
            TargetUrl = $checkUrl
            ServerModified = $_.LastWriteTime
            SharePointModified = $spModified
            ModifiedDeltaSeconds = if ($deltaSeconds -ne $null) { [Math]::Round($deltaSeconds, 3) } else { $null }
            TimestampToleranceSeconds = $TimestampToleranceSeconds
        }
    }

    if ($Migrate -and ($action -eq "Migrate" -or $action -eq "CanMigrate")) {
        $attemptedUploads++
        Write-Host ("  UPLOAD [{0}] ({1}) {2}" -f $attemptedUploads, $action, $relativePath) -ForegroundColor Cyan

        Write-MigrationLog -Level "UPLOAD_START" -Message "Uploading file" -Data @{
            Attempt = $attemptedUploads
            Action = $action
            Overwrite = [bool]$overwrite
            SourceFile = $_.FullName
            SharePointPath = $sharePointPath
        }

        $result = Copy-FileToSharePoint -SourcePath $_.FullName -SharePointPath $sharePointPath -List $spList -Overwrite:$overwrite
        if ($result.Success) {
            $uploadedOk++
            Write-Host "    ✓ Uploaded" -ForegroundColor Green
            Write-MigrationLog -Level "UPLOAD_OK" -Message "Upload succeeded" -Data @{
                Attempt = $attemptedUploads
                Action = $action
                TargetUrl = $result.FileUrl
            }
        }
        else {
            $uploadedFail++
            Write-Host "    ✗ Failed: $($result.Error)" -ForegroundColor Red
            Write-MigrationLog -Level "UPLOAD_FAIL" -Message "Upload failed" -Data @{
                Attempt = $attemptedUploads
                Action = $action
                Error = $result.Error
            }
            if ($FailFast) { throw "FailFast: stopping on first upload failure: $($result.Error)" }
        }

        if ($StopAfter -gt 0 -and $attemptedUploads -ge $StopAfter) {
            Write-Host "StopAfter reached ($StopAfter upload attempts). Stopping early." -ForegroundColor Yellow
            Write-MigrationLog -Level "WARN" -Message "StopAfter reached; stopping early" -Data @{
                StopAfter = $StopAfter
                AttemptedUploads = $attemptedUploads
                UploadedOk = $uploadedOk
                UploadedFail = $uploadedFail
            }
            break
        }
    }

    if ($enumerated % 1000 -eq 0) {
        Write-Host "  Enumerated $enumerated... kept $kept, filtered by date $filteredByDate, uploads attempted $attemptedUploads" -ForegroundColor Gray
    }
}

Write-Host "`n=== Summary ===" -ForegroundColor Yellow
Write-Host "Enumerated: $enumerated" -ForegroundColor Gray
Write-Host "Kept (after date filter): $kept" -ForegroundColor Gray
if ($start -or $end) {
    Write-Host "Filtered by date: $filteredByDate" -ForegroundColor Gray
}
Write-Host "Missing: $missingCount" -ForegroundColor Gray
Write-Host "Newer on server (CanMigrate candidates): $canMigrateCount" -ForegroundColor Gray
Write-Host "Skipped (newer in SharePoint): $skippedNewerInSharePointCount" -ForegroundColor Gray
Write-Host "Skipped (existing/same): $skippedExistingCount" -ForegroundColor Gray
Write-Host "Uploads attempted: $attemptedUploads (ok=$uploadedOk, fail=$uploadedFail)" -ForegroundColor Gray
Write-Host "Log: $LogPath" -ForegroundColor Gray

Write-MigrationLog -Level "INFO" -Message "Run summary" -Data @{
    Enumerated = $enumerated
    Kept = $kept
    FilteredByDate = $filteredByDate
    Missing = $missingCount
    CanMigrateCandidates = $canMigrateCount
    SkippedNewerInSharePoint = $skippedNewerInSharePointCount
    SkippedExisting = $skippedExistingCount
    UploadsAttempted = $attemptedUploads
    UploadsOk = $uploadedOk
    UploadsFail = $uploadedFail
}

if (Get-Module -Name PnP.PowerShell) {
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}

Write-Host "`nDone." -ForegroundColor Green

