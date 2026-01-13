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
    [string]$ReportPath = $null,
    
    [Parameter(Mandatory = $false)]
    [switch]$Migrate,  # If specified, will migrate files after comparison

    # SharePoint + Windows timestamps can differ by milliseconds / timezone kind while printing the same in CSV.
    # Use a small tolerance to avoid false NewerOnServer/NewerInSharePoint classifications.
    [Parameter(Mandatory = $false)]
    [double]$TimestampToleranceSeconds = 2
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

# Early check: Verify file server path is accessible before proceeding
Write-Host "Verifying file server path is accessible..." -ForegroundColor Cyan
if (-not (Test-Path $config.FileServerPath)) {
    throw "File server path is not accessible: $($config.FileServerPath). Please verify the path exists and you have access permissions."
}
Write-Host "File server path is accessible: $($config.FileServerPath)" -ForegroundColor Green

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
    
    # First, verify the certificate exists
    Write-Host "Checking for certificate with thumbprint: $Thumbprint..." -ForegroundColor Gray
    $cert = Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object { $_.Thumbprint -eq $Thumbprint }
    if (-not $cert) {
        $cert = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Thumbprint -eq $Thumbprint }
    }
    
    if (-not $cert) {
        throw "Certificate with thumbprint '$Thumbprint' not found in certificate store. Please verify:" + 
              "`n  1. The certificate is installed in either CurrentUser\My or LocalMachine\My" + 
              "`n  2. The thumbprint in config.json matches exactly (case-sensitive)" + 
              "`n  3. You have access to the certificate's private key"
    }
    
    Write-Host "Certificate found: $($cert.Subject) (Valid until: $($cert.NotAfter))" -ForegroundColor Gray
    
    try {
        # Connect using certificate authentication (matching existing pattern)
        Write-Host "Connecting to SharePoint site..." -ForegroundColor Gray
        Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $TenantId -WarningAction SilentlyContinue -ErrorAction Stop
        
        # Verify connection
        $web = Get-PnPWeb
        Write-Host "Connected to: $($web.Title)" -ForegroundColor Green
        return $true
    }
    catch {
        $errorDetails = $_.Exception.Message
        if ($errorDetails -like "*keyset*" -or $errorDetails -like "*key*") {
            throw "Failed to connect: Certificate access error. The certificate exists but the private key may not be accessible. " +
                  "`nError details: $errorDetails" +
                  "`n`nTroubleshooting:" +
                  "`n  1. Ensure you have permission to access the certificate's private key" +
                  "`n  2. If using a service account, ensure it has access to the certificate" +
                  "`n  3. Try running PowerShell as Administrator" +
                  "`n  4. Verify the certificate hasn't expired"
        }
        else {
            throw "Failed to connect with PnP.PowerShell: $errorDetails"
        }
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

# Helper function to transform folder names based on configuration
function Transform-FolderName {
    param(
        [string]$FolderName,
        [object]$TransformConfig
    )
    
    $transformed = $FolderName
    
    # Hardcoded replacement: "Clients" followed by anything -> "Clients"
    # This handles "Clients (ETC - Wilco)", "Clients - Old", "Clients(anything)", etc.
    # This always runs, regardless of TransformConfig
    if ($transformed -match '^Clients') {
        # If it's exactly "Clients", keep it; otherwise replace with just "Clients"
        if ($transformed -ne 'Clients') {
            $transformed = 'Clients'
        }
    }
    
    # If no TransformConfig, return after hardcoded replacement
    if (-not $TransformConfig) {
        return $transformed
    }
    
    # Apply name mappings first (if specified) - exact match replacements
    if ($TransformConfig.NameMappings) {
        foreach ($mapping in $TransformConfig.NameMappings.PSObject.Properties) {
            if ($transformed -eq $mapping.Name) {
                $transformed = $mapping.Value
                break
            }
        }
    }
    
    # Apply folder name simplification - keep only the base name before delimiters
    # This handles cases like "Clients (ETC - Wilco)" -> "Clients"
    if ($TransformConfig.SimplifyFolders) {
        # Common delimiters that indicate additional info: space+paren, space+dash, just paren
        # Remove everything after: " (", " -", "(", " -"
        $simplified = $transformed -replace '\s*\(.*$', ''  # Remove " (anything)"
        $simplified = $simplified -replace '\s+-.*$', ''    # Remove " - anything"
        $simplified = $simplified.Trim()
        if ($simplified) {
            $transformed = $simplified
        }
    }
    
    # Apply regex pattern to remove parts (if specified)
    if ($TransformConfig.RemovePattern) {
        $transformed = $transformed -replace $TransformConfig.RemovePattern, ''
    }
    
    # Trim any extra spaces that might result from pattern removal
    $transformed = $transformed.Trim()
    
    # Safety check: if transformation resulted in empty string, return original
    if ([string]::IsNullOrWhiteSpace($transformed)) {
        Write-Warning "Folder name transformation resulted in empty string for '$FolderName', using original name"
        return $FolderName
    }
    
    return $transformed
}

# Helper function to transform a full path by transforming each folder name
function Transform-Path {
    param(
        [string]$Path,
        [object]$TransformConfig
    )
    
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $Path
    }
    
    # Split path into parts
    $pathParts = $Path -split '[\\/]'
    $transformedParts = @()
    
    foreach ($part in $pathParts) {
        if ($part) {
            # Transform-FolderName now always applies hardcoded "Clients" transformation
            # even if TransformConfig is null
            $transformedPart = Transform-FolderName -FolderName $part -TransformConfig $TransformConfig
            # Only add non-empty transformed parts
            if ($transformedPart -and -not [string]::IsNullOrWhiteSpace($transformedPart)) {
                $transformedParts += $transformedPart
            }
            else {
                # If transformation resulted in empty, keep original to avoid breaking path
                Write-Warning "Path part '$part' transformed to empty, keeping original"
                $transformedParts += $part
            }
        }
    }
    
    # Rejoin with backslashes
    $result = $transformedParts -join '\'
    
    # Safety check: if result is empty, return original path
    if ([string]::IsNullOrWhiteSpace($result)) {
        Write-Warning "Path transformation resulted in empty path for '$Path', using original"
        return $Path
    }
    
    return $result
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
    
    # Debug: Show what we're working with
    Write-Host "      Library root URL: $libraryRootUrl" -ForegroundColor DarkGray
    Write-Host "      Library name: $libraryName" -ForegroundColor DarkGray
    Write-Host "      SharePoint path: $spPath" -ForegroundColor DarkGray
    
    # Construct the full server-relative URL for the file
    # The libraryRootUrl already includes the library name (e.g., /sites/ETCFiles/Documents or /sites/ETCFiles/Shared Documents)
    # However, if the library is "Shared Documents" and files are actually under a "Documents" subfolder,
    # we need to account for that. Try the most likely path first based on library name.
    
    # If library is "Shared Documents", files are often under a "Documents" subfolder
    # So try: /sites/ETCFiles/Shared Documents/Documents/ETL/Clients/file.pdf
    # Otherwise try: /sites/ETCFiles/Documents/ETL/Clients/file.pdf
    if ($libraryName -eq "Shared Documents" -and $spPath -notmatch '^Documents/') {
        $fullServerRelativeUrl = "$libraryRootUrl/Documents/$spPath"
    }
    else {
        $fullServerRelativeUrl = "$libraryRootUrl/$spPath"
    }
    
    # Remove any duplicate library names if they appear in the path
    # Sometimes the path might already include "Documents" or "Shared Documents"
    $fullServerRelativeUrl = $fullServerRelativeUrl -replace "/(Documents|Shared Documents)/(Documents|Shared Documents)/", '/$1/'
    
    Write-Host "      Primary URL: $fullServerRelativeUrl" -ForegroundColor DarkGray
    
    # Store the primary URL being checked for logging
    $primaryUrl = $fullServerRelativeUrl
    
    # Try to get the file directly using the constructed URL
    # Direct file access bypasses the 5000 item list view threshold
    try {
        # First, try Get-PnPFile with -AsListItem to get both file and metadata in one call
        $fileItem = Get-PnPFile -Url $fullServerRelativeUrl -AsListItem -ErrorAction Stop
        if ($fileItem) {
            return @{
                Name = $fileItem.FieldValues.FileLeafRef
                Size = [long]$fileItem.FieldValues.File_x0020_Size
                Modified = [DateTime]$fileItem.FieldValues.Modified
                Path = $SharePointPath
                NormalizedPath = (Normalize-Path -Path $SharePointPath)
                CheckedUrl = $fullServerRelativeUrl
            }
        }
    }
    catch {
        # If -AsListItem doesn't work, try the regular Get-PnPFile
        try {
            $file = Get-PnPFile -Url $fullServerRelativeUrl -ErrorAction Stop
            if ($file) {
                # Get list item metadata using the file's server-relative URL
                $fileItem = Get-PnPListItem -List $List -Query "<View><Query><Where><Eq><FieldRef Name='FileRef'/><Value Type='Text'>$($file.ServerRelativeUrl)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($fileItem) {
                    return @{
                        Name = $fileItem.FieldValues.FileLeafRef
                        Size = [long]$fileItem.FieldValues.File_x0020_Size
                        Modified = [DateTime]$fileItem.FieldValues.Modified
                        Path = $SharePointPath
                        NormalizedPath = (Normalize-Path -Path $SharePointPath)
                        CheckedUrl = $fullServerRelativeUrl
                    }
                }
            }
        }
        catch {
            # File not found at that exact path, try alternative approaches
            Write-Host "      Error: $($_.Exception.Message)" -ForegroundColor DarkRed
        }
    }
    
    # Alternative: Try with different library name prefixes (only if initial path failed)
    # The issue is that SharePoint might have nested structures like:
    # - Library: "Shared Documents" -> Root: /sites/ETCFiles/Shared Documents
    # - But files are actually at: /sites/ETCFiles/Shared Documents/Documents/ETL/Clients/...
    # So we need to try adding "Documents" if the library is "Shared Documents"
    
    $libraryRootLower = $libraryRootUrl.ToLower()
    $libraryNameLower = $libraryName.ToLower()
    $rootEndsWithLibrary = $libraryRootLower.EndsWith("/$libraryNameLower") -or $libraryRootLower.EndsWith("/shared documents")
    
    $possiblePaths = @()
    
    # If library is "Shared Documents" and path doesn't start with "Documents", try adding it
    # This handles the case where files are under Shared Documents/Documents/...
    if ($libraryNameLower -eq "shared documents" -and $spPath -notmatch '^Documents/') {
        $possiblePaths += "Documents/$spPath"
    }
    
    # Only try other alternatives if the root doesn't already include the library name
    if (-not $rootEndsWithLibrary) {
        if ($spPath -notmatch '^Documents/') {
            $possiblePaths += "Documents/$spPath"
        }
        if ($spPath -notmatch '^Shared Documents/') {
            $possiblePaths += "Shared Documents/$spPath"
        }
        if ($spPath -notmatch "^$([regex]::Escape($libraryName))/") {
            $possiblePaths += "$libraryName/$spPath"
        }
    }
    # Always try the path as-is (without library prefix) as last resort
    $possiblePaths += $spPath
    
    foreach ($testPath in $possiblePaths) {
        try {
            $testUrl = "$libraryRootUrl/$testPath"
            Write-Host "      Trying alternative: $testUrl" -ForegroundColor DarkGray
            # Try with -AsListItem first
            $fileItem = Get-PnPFile -Url $testUrl -AsListItem -ErrorAction Stop
            if ($fileItem) {
                Write-Host "      ✓ Found at alternative URL" -ForegroundColor Green
                return @{
                    Name = $fileItem.FieldValues.FileLeafRef
                    Size = [long]$fileItem.FieldValues.File_x0020_Size
                    Modified = [DateTime]$fileItem.FieldValues.Modified
                    Path = $SharePointPath
                    NormalizedPath = (Normalize-Path -Path $SharePointPath)
                    CheckedUrl = $testUrl
                }
            }
        }
        catch {
            # Try regular Get-PnPFile if -AsListItem failed
            try {
                $file = Get-PnPFile -Url $testUrl -ErrorAction Stop
                if ($file) {
                    $fileItem = Get-PnPListItem -List $List -Query "<View><Query><Where><Eq><FieldRef Name='FileRef'/><Value Type='Text'>$($file.ServerRelativeUrl)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue | Select-Object -First 1
                    if ($fileItem) {
                        Write-Host "      ✓ Found at alternative URL" -ForegroundColor Green
                        return @{
                            Name = $fileItem.FieldValues.FileLeafRef
                            Size = [long]$fileItem.FieldValues.File_x0020_Size
                            Modified = [DateTime]$fileItem.FieldValues.Modified
                            Path = $SharePointPath
                            NormalizedPath = (Normalize-Path -Path $SharePointPath)
                            CheckedUrl = $testUrl
                        }
                    }
                }
            }
            catch {
                continue
            }
        }
    }
    
    # Last resort: Try using Get-PnPFolder to navigate to the folder, then get the file
    try {
        $pathParts = $spPath -split '/'
        $fileName = $pathParts[-1]
        $folderPathParts = $pathParts[0..($pathParts.Length - 2)]
        $folderPath = $folderPathParts -join '/'
        
        if ($folderPath) {
            $folderUrl = "$libraryRootUrl/$folderPath"
            Write-Host "      Trying folder navigation: $folderUrl" -ForegroundColor DarkGray
            $folder = Get-PnPFolder -Url $folderUrl -ErrorAction Stop
            if ($folder) {
                $fileUrl = "$folderUrl/$fileName"
                $fileItem = Get-PnPFile -Url $fileUrl -AsListItem -ErrorAction Stop
                if ($fileItem) {
                    Write-Host "      ✓ Found via folder navigation" -ForegroundColor Green
                    return @{
                        Name = $fileItem.FieldValues.FileLeafRef
                        Size = [long]$fileItem.FieldValues.File_x0020_Size
                        Modified = [DateTime]$fileItem.FieldValues.Modified
                        Path = $SharePointPath
                        NormalizedPath = (Normalize-Path -Path $SharePointPath)
                        CheckedUrl = $fileUrl
                    }
                }
            }
        }
    }
    catch {
        # Folder navigation failed
    }
    
    # Note: We don't use list queries as a fallback because they can hit the 5000 item limit
    # Direct file access (Get-PnPFile -Url) is the only reliable method for large folders
    # If the file wasn't found via direct URL, it likely doesn't exist at that path
    # Return null with the URL that was checked for logging
    return $null
}

# Function to ensure a folder path exists in SharePoint
function Ensure-SharePointFolder {
    param(
        [Microsoft.SharePoint.Client.List]$List,
        [string]$FolderPath  # Relative path like "clients\client1" (with backslashes)
    )
    
    if (-not $List -or -not $FolderPath) {
        return $false
    }
    
    # Convert backslashes to forward slashes for SharePoint
    $spFolderPath = $FolderPath -replace '\\', '/'
    
    # Get library root
    $libraryRootUrl = $List.RootFolder.ServerRelativeUrl.TrimEnd('/')
    $fullFolderPath = "$libraryRootUrl/$spFolderPath"
    
    try {
        # Try to get the folder - if it exists, we're done
        $folder = Get-PnPFolder -Url $fullFolderPath -ErrorAction Stop
        if ($folder) {
            return $true
        }
    }
    catch {
        # Folder doesn't exist, need to create it
        # Split path into parts and create each level
        $pathParts = $spFolderPath -split '/'
        $currentPath = $libraryRootUrl
        
        foreach ($part in $pathParts) {
            if ($part) {
                $currentPath = "$currentPath/$part"
                try {
                    # Try to get folder at this level
                    $folder = Get-PnPFolder -Url $currentPath -ErrorAction Stop
                }
                catch {
                    # Folder doesn't exist, create it
                    try {
                        $parentPath = $currentPath.Substring(0, $currentPath.LastIndexOf('/'))
                        $folderName = $part
                        Add-PnPFolder -Name $folderName -Folder $parentPath -ErrorAction Stop
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

# Function to upload a file to SharePoint
function Copy-FileToSharePoint {
    param(
        [string]$SourcePath,
        [string]$SharePointPath,  # Relative path like "clients\client1\file.pdf"
        [Microsoft.SharePoint.Client.List]$List,
        [string]$LibraryName
    )
    
    if (-not (Test-Path $SourcePath)) {
        return @{
            Success = $false
            Error = "Source file not found: $SourcePath"
        }
    }
    
    try {
        # Convert SharePoint path to forward slashes and get folder path
        $spPath = $SharePointPath -replace '\\', '/'
        $fileName = Split-Path -Leaf $spPath
        $folderPath = $spPath.Substring(0, $spPath.LastIndexOf('/'))
        
        # Ensure folder exists
        if ($folderPath) {
            $folderCreated = Ensure-SharePointFolder -List $List -FolderPath ($folderPath -replace '/', '\')
            if (-not $folderCreated) {
                return @{
                    Success = $false
                    Error = "Failed to create folder structure: $folderPath"
                }
            }
        }
        
        # Get library root
        $libraryRootUrl = $List.RootFolder.ServerRelativeUrl.TrimEnd('/')
        $targetFolderUrl = if ($folderPath) { "$libraryRootUrl/$folderPath" } else { $libraryRootUrl }
        
        # Show file size for context
        $fileSize = (Get-Item $SourcePath).Length
        $fileSizeMB = [Math]::Round($fileSize / 1MB, 2)
        Write-Host "      File size: $fileSizeMB MB - Uploading to SharePoint..." -ForegroundColor DarkGray
        [Console]::Out.Flush()
        
        # Upload the file (this can take time for larger files)
        $file = Add-PnPFile -Path $SourcePath -Folder $targetFolderUrl -ErrorAction Stop
        
        if ($file) {
            return @{
                Success = $true
                FileUrl = $file.ServerRelativeUrl
                Error = $null
            }
        }
        else {
            return @{
                Success = $false
                Error = "Upload completed but file object not returned"
            }
        }
    }
    catch {
        return @{
            Success = $false
            Error = $_.Exception.Message
        }
    }
}

# Function to get all files from file server
# Can optionally process files as they're found via ProcessFileCallback
function Get-FileServerFiles {
    param(
        [string]$RootPath,
        [DateTime]$StartDate = $null,
        [DateTime]$EndDate = $null,
        [string]$SharePointBasePath = $null,  # Optional prefix path in SharePoint (e.g., "etc")
        [object]$FolderNameTransform = $null,  # Optional folder name transformation config
        [scriptblock]$ProcessFileCallback = $null  # Optional callback to process each file as it's found
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
    
    # Apply transformation to root folder name for display
    $transformedRootFolderName = if ($FolderNameTransform) {
        Transform-FolderName -FolderName $rootFolderName -TransformConfig $FolderNameTransform
    } else {
        $rootFolderName
    }
    
    Write-Host "Root folder name: $rootFolderName" -ForegroundColor Gray
    if ($transformedRootFolderName -ne $rootFolderName) {
        Write-Host "  Transformed to: $transformedRootFolderName" -ForegroundColor Gray
    }
    
    # Show the correct SharePoint path mapping including SharePointBasePath if provided
    if ($SharePointBasePath) {
        Write-Host "  Files will be mapped to SharePoint path: $SharePointBasePath\$transformedRootFolderName\[relative path]" -ForegroundColor Gray
        Write-Host "    Example: $SharePointBasePath\$transformedRootFolderName\subfolder\file.pdf" -ForegroundColor DarkGray
    }
    else {
        Write-Host "  Files will be mapped to SharePoint path: $transformedRootFolderName\[relative path]" -ForegroundColor Gray
        Write-Host "    Example: $transformedRootFolderName\subfolder\file.pdf" -ForegroundColor DarkGray
    }
    
    if ($FolderNameTransform) {
        Write-Host "  Note: Folder name transformations are applied to all paths" -ForegroundColor DarkGray
    }
    
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
            
            # Apply folder name transformations if configured
            if ($FolderNameTransform) {
                $transformedRootFolderName = Transform-FolderName -FolderName $rootFolderName -TransformConfig $FolderNameTransform
                $transformedRelativePath = Transform-Path -Path $relativePath -TransformConfig $FolderNameTransform
            }
            else {
                $transformedRootFolderName = $rootFolderName
                $transformedRelativePath = $relativePath
            }
            
            # Create SharePoint path with optional base path
            if ($SharePointBasePath) {
                $sharePointRelativePath = "$SharePointBasePath\$transformedRootFolderName\$transformedRelativePath"
            }
            else {
                $sharePointRelativePath = "$transformedRootFolderName\$transformedRelativePath"
            }
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
        
        # Apply folder name transformations if configured
        if ($FolderNameTransform) {
            # Transform the root folder name
            $transformedRootFolderName = Transform-FolderName -FolderName $rootFolderName -TransformConfig $FolderNameTransform
            # Transform the relative path (all folder names in the path)
            $transformedRelativePath = Transform-Path -Path $relativePath -TransformConfig $FolderNameTransform
            
            # Debug: Show transformation for first few files
            if ($files.Count -lt 3) {
                Write-Host "  DEBUG: Root '$rootFolderName' -> '$transformedRootFolderName'" -ForegroundColor Cyan
                Write-Host "  DEBUG: Relative '$relativePath' -> '$transformedRelativePath'" -ForegroundColor Cyan
            }
            
            # Safety check: ensure we have valid folder names after transformation
            if ([string]::IsNullOrWhiteSpace($transformedRootFolderName)) {
                Write-Warning "Transformation resulted in empty root folder name for '$rootFolderName', using original"
                $transformedRootFolderName = $rootFolderName
            }
            if ([string]::IsNullOrWhiteSpace($transformedRelativePath)) {
                Write-Warning "Transformation resulted in empty relative path for '$relativePath', using original"
                $transformedRelativePath = $relativePath
            }
        }
        else {
            $transformedRootFolderName = $rootFolderName
            $transformedRelativePath = $relativePath
        }
        
        # Create SharePoint path: [base path] + transformed root folder name + transformed relative path
        # If SharePointBasePath is specified (e.g., "etc"), path becomes "etc\clients\file.pdf"
        # Otherwise, just "clients\file.pdf"
        if ($SharePointBasePath) {
            $sharePointRelativePath = "$SharePointBasePath\$transformedRootFolderName\$transformedRelativePath"
        }
        else {
            $sharePointRelativePath = "$transformedRootFolderName\$transformedRelativePath"
        }
        
        # Final safety check: ensure we have a valid SharePoint path
        if ([string]::IsNullOrWhiteSpace($sharePointRelativePath)) {
            Write-Warning "Skipping file due to empty SharePoint path: $($fileInfo.FullName)"
            Write-Warning "  Original relative path: $relativePath"
            Write-Warning "  Transformed root: $transformedRootFolderName"
            Write-Warning "  Transformed relative: $transformedRelativePath"
            return
        }
        
        # Normalize path for comparison (lowercase, backslashes)
        $normalizedPath = Normalize-Path -Path $sharePointRelativePath
        
        # Debug: Log first few files to verify they're being processed
        if ($files.Count -lt 5) {
            Write-Verbose "Processing file: $($fileInfo.Name)" -Verbose
            Write-Verbose "  Original relative path: $relativePath" -Verbose
            Write-Verbose "  Transformed relative path: $transformedRelativePath" -Verbose
            Write-Verbose "  Final SharePoint path: $sharePointRelativePath" -Verbose
        }
        
        $fileData = @{
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
        
        # Store in dictionary for return value
        $files[$normalizedPath] = $fileData
        
        # If callback provided, process file immediately
        if ($ProcessFileCallback) {
            try {
                & $ProcessFileCallback -FileData $fileData
            }
            catch {
                Write-Warning "Error processing file $($fileData.FullPath): $_"
            }
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

# Step 1: Connect to SharePoint early (before scanning, so we can process files as we find them)
Write-Host "Step 1: Connecting to SharePoint..." -ForegroundColor Cyan
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

# Step 2: Scan file server and process files as they're found
Write-Host "`nStep 2: Scanning file server and checking files in SharePoint (streaming mode)..." -ForegroundColor Cyan
Write-Host "Note: Files will be checked and migrated as they're found (no need to wait for full scan)" -ForegroundColor Gray

# Get optional SharePoint base path (prefix folder in SharePoint)
$sharePointBasePath = if ($config.SharePointBasePath) { $config.SharePointBasePath } else { $null }

# Get optional folder name transformation config
$folderNameTransform = if ($config.FolderNameTransform) { $config.FolderNameTransform } else { $null }
if ($folderNameTransform) {
    Write-Host "Folder name transformation enabled:" -ForegroundColor Cyan
    if ($folderNameTransform.RemovePattern) {
        Write-Host "  Remove pattern: $($folderNameTransform.RemovePattern)" -ForegroundColor Gray
    }
    if ($folderNameTransform.NameMappings) {
        Write-Host "  Name mappings configured: $($folderNameTransform.NameMappings.PSObject.Properties.Count) mappings" -ForegroundColor Gray
    }
}

# Initialize tracking variables for streaming processing
$results = @()
$missingCount = 0
$newerOnServerCount = 0
$newerInSharePointCount = 0
$identicalCount = 0
$checkedCount = 0
$fileServerFiles = @{}
$lockedFiles = @{}

# Create callback to process each file as it's found
$processFileCallback = {
    param($FileData)
    
    $checkedCount++
    $script:checkedCount = $checkedCount
    
    # Check if this specific file exists in SharePoint
    $spFile = Get-SharePointFile -SiteUrl $config.SharePointSiteUrl -List $spList -SharePointPath $FileData.SharePointPath
    
    if ($spFile) {
        # File exists in SharePoint - compare modification dates
        $modifiedDeltaSeconds = ($FileData.Modified - $spFile.Modified).TotalSeconds

        if ($modifiedDeltaSeconds -gt $TimestampToleranceSeconds) {
            # File is newer on server - can be migrated
            $result = [PSCustomObject]@{
                Status = "NewerOnServer"
                FilePath = $FileData.NormalizedPath
                SharePointPath = $FileData.SharePointPath
                ServerSize = $FileData.Size
                ServerModified = $FileData.Modified
                SharePointSize = $spFile.Size
                SharePointModified = $spFile.Modified
                ModifiedDeltaSeconds = [Math]::Round($modifiedDeltaSeconds, 3)
                Action = "CanMigrate"
            }
            $script:results += $result
            $script:newerOnServerCount++
            
            # If migrating, migrate immediately
            if ($Migrate) {
                Write-Host "  [$checkedCount] Migrating (newer on server): $($FileData.SharePointPath)" -ForegroundColor Cyan
                $migrationResult = Copy-FileToSharePoint -SourcePath $FileData.FullPath -SharePointPath $FileData.SharePointPath -List $spList
                if ($migrationResult.Success) {
                    Write-Host "    ✓ Migrated successfully" -ForegroundColor Green
                }
                else {
                    Write-Host "    ✗ Migration failed: $($migrationResult.Error)" -ForegroundColor Red
                }
            }
        }
        elseif ($modifiedDeltaSeconds -lt (-1 * $TimestampToleranceSeconds)) {
            # File is newer in SharePoint - skip to avoid overwriting
            $result = [PSCustomObject]@{
                Status = "NewerInSharePoint"
                FilePath = $FileData.NormalizedPath
                SharePointPath = $FileData.SharePointPath
                ServerSize = $FileData.Size
                ServerModified = $FileData.Modified
                SharePointSize = $spFile.Size
                SharePointModified = $spFile.Modified
                ModifiedDeltaSeconds = [Math]::Round($modifiedDeltaSeconds, 3)
                Action = "Skip"
            }
            $script:results += $result
            $script:newerInSharePointCount++
        }
        else {
            # Same modification time (or very close)
            if ($FileData.Size -ne $spFile.Size) {
                # Same time but different size - might need migration
                $result = [PSCustomObject]@{
                    Status = "SizeMismatch"
                    FilePath = $FileData.NormalizedPath
                    SharePointPath = $FileData.SharePointPath
                    ServerSize = $FileData.Size
                    ServerModified = $FileData.Modified
                    SharePointSize = $spFile.Size
                    SharePointModified = $spFile.Modified
                    ModifiedDeltaSeconds = [Math]::Round($modifiedDeltaSeconds, 3)
                    Action = "Review"
                }
                $script:results += $result
            }
            else {
                $script:identicalCount++
            }
        }
    }
    else {
        # File doesn't exist in SharePoint - needs migration
        $result = [PSCustomObject]@{
            Status = "Missing"
            FilePath = $FileData.NormalizedPath
            SharePointPath = $FileData.SharePointPath
            ServerSize = $FileData.Size
            ServerModified = $FileData.Modified
            SharePointSize = $null
            SharePointModified = $null
            Action = "Migrate"
        }
        $script:results += $result
        $script:missingCount++
        
        # If migrating, migrate immediately
        if ($Migrate) {
            Write-Host "  [$checkedCount] Migrating (missing): $($FileData.SharePointPath)" -ForegroundColor Cyan
            $migrationResult = Copy-FileToSharePoint -SourcePath $FileData.FullPath -SharePointPath $FileData.SharePointPath -List $spList
            if ($migrationResult.Success) {
                Write-Host "    ✓ Migrated successfully" -ForegroundColor Green
            }
            else {
                Write-Host "    ✗ Migration failed: $($migrationResult.Error)" -ForegroundColor Red
            }
        }
    }
    
    # Progress update every 100 files
    if ($checkedCount % 100 -eq 0) {
        Write-Host "  Processed $checkedCount files (Missing: $script:missingCount, Newer: $script:newerOnServerCount, Identical: $script:identicalCount)..." -ForegroundColor Gray
    }
}

# Scan file server with callback for streaming processing
$fileServerResult = Get-FileServerFiles -RootPath $config.FileServerPath -StartDate $startDate -EndDate $endDate -SharePointBasePath $sharePointBasePath -FolderNameTransform $folderNameTransform -ProcessFileCallback $processFileCallback
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

# Add locked files to results (process them after scan completes)
$lockedCount = $fileServerResult.LockedCount
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
try {
    $results | Export-Csv -Path $ReportPath -NoTypeInformation -ErrorAction Stop
    
    # Excel row limit is 1,048,576 rows (Excel 2007+)
    $excelRowLimit = 1048576
    if ($results.Count -gt $excelRowLimit) {
        Write-Host "`n⚠️  WARNING: Report has $($results.Count) records, which exceeds Excel's limit of $excelRowLimit rows" -ForegroundColor Yellow
        Write-Host "The CSV file was generated successfully, but Excel will only display the first $excelRowLimit rows." -ForegroundColor Yellow
        Write-Host "Consider using a different tool (like Power BI, or split the report) to view all records." -ForegroundColor Yellow
    }
    
    Write-Host "`nReport saved to: $ReportPath ($($results.Count) records)" -ForegroundColor Green
}
catch {
    Write-Host "`nWarning: Failed to generate comparison report CSV: $_" -ForegroundColor Yellow
    Write-Host "Comparison completed, but report generation failed. Results are still available in memory." -ForegroundColor Yellow
}

# Generate SPMT-compatible migration manifest in JSON format (for files that need migration)
$spmtManifestPath = $ReportPath -replace '\.csv$', '-spmt-manifest.json'
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
            # Convert SharePoint path from backslashes to forward slashes for SPMT
            $targetPath = $result.SharePointPath -replace '\\', '/'
            
            # SPMT JSON format: Array of objects with source and target information
            $spmtManifest += @{
                SourcePath = $fullPath
                TargetUrl = $config.SharePointSiteUrl
                TargetDocumentLibrary = $libraryName
                TargetPath = $targetPath
                FileName = Split-Path -Leaf $fullPath
            }
        }
    }
}

if ($spmtManifest.Count -gt 0) {
    # Convert to JSON with proper formatting
    $jsonContent = $spmtManifest | ConvertTo-Json -Depth 10
    $jsonContent | Out-File -FilePath $spmtManifestPath -Encoding UTF8
    Write-Host "SPMT migration manifest (JSON) saved to: $spmtManifestPath ($($spmtManifest.Count) files)" -ForegroundColor Green
    Write-Host "Note: Review the JSON format and adjust if needed for your SPMT version." -ForegroundColor Yellow
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

# Step 3: Migration summary (files were migrated during scan if -Migrate was specified)
if ($Migrate) {
    Write-Host "`n=== Migration Summary ===" -ForegroundColor Yellow
    Write-Host "Files were migrated as they were found during the scan." -ForegroundColor Gray
    Write-Host "Total files processed: $checkedCount" -ForegroundColor Cyan
    Write-Host "Files migrated: $($missingCount + $newerOnServerCount)" -ForegroundColor Green
    
    # Get files that were migrated
    $filesToMigrate = $results | Where-Object { $_.Action -eq "Migrate" -or $_.Action -eq "CanMigrate" }
    
    if ($filesToMigrate.Count -eq 0) {
        Write-Host "No files needed migration. All files are either already in SharePoint or should be skipped." -ForegroundColor Gray
    }
    else {
        Write-Host "Files that needed migration: $($filesToMigrate.Count)" -ForegroundColor Cyan
        
        # Warn if there are a lot of files
        if ($filesToMigrate.Count -gt 10000) {
            Write-Host "`n⚠️  WARNING: Large migration detected ($($filesToMigrate.Count) files)" -ForegroundColor Yellow
            Write-Host "This migration may take a very long time and could hit SharePoint rate limits." -ForegroundColor Yellow
            Write-Host "Estimated time: ~$([Math]::Round($filesToMigrate.Count * 2 / 60, 0)) minutes (assuming 2 seconds per file)" -ForegroundColor Yellow
            Write-Host ""
            $confirm = Read-Host "Do you want to continue? (Y/N)"
            if ($confirm -ne 'Y' -and $confirm -ne 'y') {
                Write-Host "Migration cancelled by user." -ForegroundColor Gray
                exit 0
            }
        }
        elseif ($filesToMigrate.Count -gt 1000) {
            Write-Host "Note: Large migration ($($filesToMigrate.Count) files) - this may take some time" -ForegroundColor Yellow
        }
        
        Write-Host ""
        
        $migrationResults = @()
        $migratedCount = 0
        $failedCount = 0
        $skippedCount = 0
        $currentFile = 0
        
        foreach ($fileToMigrate in $filesToMigrate) {
            $currentFile++
            $currentFilePercent = [Math]::Round(($currentFile / $filesToMigrate.Count) * 100, 1)
            
            # Find the full source path
            $sourcePath = $null
            foreach ($filePath in $fileServerFiles.Keys) {
                if ($fileServerFiles[$filePath].NormalizedPath -eq $fileToMigrate.FilePath) {
                    $sourcePath = $fileServerFiles[$filePath].FullPath
                    break
                }
            }
            
            if (-not $sourcePath) {
                Write-Host "  [$currentFile/$($filesToMigrate.Count)] ✗ Skipped: $($fileToMigrate.SharePointPath) - Source file not found" -ForegroundColor Yellow
                $migrationResults += [PSCustomObject]@{
                    SharePointPath = $fileToMigrate.SharePointPath
                    SourcePath = $null
                    Status = "Skipped"
                    Error = "Source file not found"
                }
                $skippedCount++
                continue
            }
            
            # Check if file already exists (might have been uploaded by another process)
            $existingFile = Get-SharePointFile -SiteUrl $config.SharePointSiteUrl -List $spList -SharePointPath $fileToMigrate.SharePointPath
            if ($existingFile) {
                Write-Host "  [$currentFile/$($filesToMigrate.Count)] ⊙ Skipped: $($fileToMigrate.SharePointPath) - Already exists in SharePoint" -ForegroundColor Gray
                $migrationResults += [PSCustomObject]@{
                    SharePointPath = $fileToMigrate.SharePointPath
                    SourcePath = $sourcePath
                    Status = "Skipped"
                    Error = "File already exists in SharePoint"
                }
                $skippedCount++
                continue
            }
            
            $fileSizeMB = [Math]::Round((Get-Item $sourcePath).Length / 1MB, 2)
            Write-Host "  [$currentFile/$($filesToMigrate.Count)] ($currentFilePercent%) Uploading: $($fileToMigrate.SharePointPath) ($fileSizeMB MB)" -ForegroundColor Cyan
            Write-Host "    Starting upload..." -ForegroundColor Gray
            [Console]::Out.Flush()
            
            # Upload the file
            $uploadResult = Copy-FileToSharePoint -SourcePath $sourcePath -SharePointPath $fileToMigrate.SharePointPath -List $spList -LibraryName $libraryName
            
            if ($uploadResult.Success) {
                Write-Host "    ✓ Successfully uploaded" -ForegroundColor Green
                $migrationResults += [PSCustomObject]@{
                    SharePointPath = $fileToMigrate.SharePointPath
                    SourcePath = $sourcePath
                    Status = "Success"
                    FileUrl = $uploadResult.FileUrl
                    Error = $null
                }
                $migratedCount++
            }
            else {
                Write-Host "    ✗ Failed: $($uploadResult.Error)" -ForegroundColor Red
                $migrationResults += [PSCustomObject]@{
                    SharePointPath = $fileToMigrate.SharePointPath
                    SourcePath = $sourcePath
                    Status = "Failed"
                    FileUrl = $null
                    Error = $uploadResult.Error
                }
                $failedCount++
            }
        }
        
        # Generate migration report
        $migrationReportPath = $ReportPath -replace '\.csv$', '-migration-results.csv'
        try {
            $migrationResults | Export-Csv -Path $migrationReportPath -NoTypeInformation -ErrorAction Stop
            Write-Host "`n=== Migration Results ===" -ForegroundColor Yellow
            Write-Host "Total files processed: $($filesToMigrate.Count)" -ForegroundColor White
            Write-Host "Successfully migrated: $migratedCount" -ForegroundColor Green
            Write-Host "Failed: $failedCount" -ForegroundColor Red
            Write-Host "Skipped: $skippedCount" -ForegroundColor Gray
            Write-Host "`nMigration report saved to: $migrationReportPath ($($migrationResults.Count) records)" -ForegroundColor Green
        }
        catch {
            Write-Host "`nWarning: Failed to generate migration report CSV: $_" -ForegroundColor Yellow
            Write-Host "Migration completed successfully, but report generation failed." -ForegroundColor Yellow
            Write-Host "Summary: $migratedCount migrated, $failedCount failed, $skippedCount skipped" -ForegroundColor Gray
        }
    }
}

# Disconnect from SharePoint
if (Get-Module -Name PnP.PowerShell) {
    Disconnect-PnPOnline
}

Write-Host "`nComparison complete!" -ForegroundColor Green
if ($Migrate) {
    Write-Host "Migration complete!" -ForegroundColor Green
}
