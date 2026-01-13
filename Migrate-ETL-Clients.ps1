<#
.SYNOPSIS
    One-time migration script for ETL Clients folder to SharePoint
    Hardcoded for: \\etlrom-dc1\lab\shared\clients (etc - wilco) -> SharePoint ETL/Clients
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$ConfigPath,
    
    [Parameter(Mandatory = $false)]
    [switch]$Migrate
)

# Error handling
$ErrorActionPreference = "Stop"

# Get paths from config (or use hardcoded defaults)
$sourcePath = if ($config.FileServerPath) { $config.FileServerPath } else { "\\etlrom-dc1\lab\shared\clients (etc - wilco)" }
$targetSharePointBasePath = if ($config.SharePointBasePath) { $config.SharePointBasePath } else { "ETL" }
$targetRootFolderName = "Clients"  # Always transform root folder to "Clients"

# Load configuration
Write-Host "Loading configuration from $ConfigPath..." -ForegroundColor Cyan
if (-not (Test-Path $ConfigPath)) {
    throw "Configuration file not found: $ConfigPath"
}

$config = Get-Content $ConfigPath | ConvertFrom-Json

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
        [Microsoft.SharePoint.Client.List]$List
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
        $targetFolderUrl = if ($folderPath) { "$libraryRootUrl/$folderPath" } else { $libraryRootUrl }
        
        $file = Add-PnPFile -Path $SourcePath -Folder $targetFolderUrl -ErrorAction Stop
        
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
Write-Host "Target: $($config.SharePointSiteUrl)/Documents/$targetSharePointBasePath/$targetRootFolderName" -ForegroundColor Cyan
Write-Host ""

# Step 1: Scan source folder
Write-Host "Step 1: Scanning source folder..." -ForegroundColor Cyan
$files = @()
$fileCount = 0

Get-ChildItem -Path $sourcePath -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
    $fileCount++
    $relativePath = $_.FullName -replace [regex]::Escape($sourcePath), ""
    $relativePath = $relativePath.TrimStart('\', '/')
    
    # Build SharePoint path: ETL/Clients/[relative path]
    # Only transform the root folder name, keep subfolders as-is
    $sharePointPath = "$targetSharePointBasePath\$targetRootFolderName\$relativePath"
    
    $files += @{
        FullPath = $_.FullName
        RelativePath = $relativePath
        SharePointPath = $sharePointPath
        Name = $_.Name
        Size = $_.Length
        Modified = $_.LastWriteTime
    }
    
    if ($files.Count % 1000 -eq 0) {
        Write-Host "  Scanned $($files.Count) files..." -ForegroundColor Gray
    }
}

Write-Host "Found $($files.Count) files to process" -ForegroundColor Green

# Step 2: Connect to SharePoint
Write-Host "`nStep 2: Connecting to SharePoint..." -ForegroundColor Cyan
Connect-SharePoint -TenantId $config.TenantId -ClientId $config.ClientId -Thumbprint $config.Thumbprint -SiteUrl $config.SharePointSiteUrl

# Get Documents library
$spList = Get-PnPList -Identity "Documents" -ErrorAction SilentlyContinue
if (-not $spList) {
    throw "Could not find 'Documents' library in SharePoint"
}

# Step 3: Compare files (optional - check what exists)
Write-Host "`nStep 3: Checking existing files in SharePoint..." -ForegroundColor Cyan
$filesToMigrate = @()
$existingCount = 0
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
        $existingFile = Get-PnPFile -Url $checkUrl -ErrorAction SilentlyContinue
        if ($existingFile) {
            $existingCount++
        }
        else {
            $filesToMigrate += $file
        }
    }
    catch {
        # File doesn't exist, add to migration list
        $filesToMigrate += $file
    }
}

Write-Host "Files already in SharePoint: $existingCount" -ForegroundColor Gray
Write-Host "Files to migrate: $($filesToMigrate.Count)" -ForegroundColor Cyan

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
        
        foreach ($file in $filesToMigrate) {
            $currentFile++
            $percent = [Math]::Round(($currentFile / $filesToMigrate.Count) * 100, 1)
            
            Write-Host "  [$currentFile/$($filesToMigrate.Count)] ($percent%) Uploading: $($file.RelativePath)" -ForegroundColor Cyan
            Write-Host "    Starting upload..." -ForegroundColor Gray
            
            # Flush output buffer to ensure message appears immediately
            [Console]::Out.Flush()
            
            $result = Copy-FileToSharePoint -SourcePath $file.FullPath -SharePointPath $file.SharePointPath -List $spList
            
            if ($result.Success) {
                Write-Host "    ✓ Successfully uploaded" -ForegroundColor Green
                $migratedCount++
            }
            else {
                Write-Host "    ✗ Failed: $($result.Error)" -ForegroundColor Red
                $failedCount++
            }
            
            # Flush after each file to ensure progress is visible
            [Console]::Out.Flush()
        }
        
        Write-Host "`n=== Migration Results ===" -ForegroundColor Yellow
        Write-Host "Successfully migrated: $migratedCount" -ForegroundColor Green
        Write-Host "Failed: $failedCount" -ForegroundColor Red
    }
}
else {
    Write-Host "`nNote: Use -Migrate parameter to actually upload files" -ForegroundColor Yellow
    Write-Host "Example: .\Migrate-ETL-Clients.ps1 -ConfigPath config.json -Migrate" -ForegroundColor Gray
}

# Disconnect
if (Get-Module -Name PnP.PowerShell) {
    Disconnect-PnPOnline
}

Write-Host "`nScript complete!" -ForegroundColor Green
