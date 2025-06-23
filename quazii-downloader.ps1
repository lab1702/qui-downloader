<#
.SYNOPSIS
    Downloads and installs the latest QuaziiUI addon for World of Warcraft.

.DESCRIPTION
    This script downloads the latest release of QuaziiUI from GitHub and installs it
    to the World of Warcraft AddOns directory. It includes automatic WoW path detection,
    comprehensive error handling, and security validations.

.PARAMETER WoWPath
    Custom path to World of Warcraft installation directory. If not specified, the script
    will attempt to auto-detect the installation.

.PARAMETER Force
    Skip confirmation prompts for destructive operations.

.PARAMETER LogFile
    Path to log file for detailed operation logging. If not specified, logs to console only.


.EXAMPLE
    .\quazii-downloader.ps1
    Downloads and installs QuaziiUI with auto-detected WoW path.

.EXAMPLE
    .\quazii-downloader.ps1 -WoWPath "D:\Games\World of Warcraft" -Force
    Installs to custom WoW path without confirmation prompts.

.NOTES
    Author: Snackington
    Version: 2.0
    Requires: PowerShell 5.1 or higher
#>

#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false, HelpMessage = "Path to World of Warcraft installation directory")]
    [ValidateScript({
        if ($_ -and -not (Test-Path $_ -PathType Container)) {
            throw "WoW path '$_' does not exist or is not a directory"
        }
        return $true
    })]
    [string]$WoWPath,

    [Parameter(Mandatory = $false, HelpMessage = "Skip confirmation prompts")]
    [switch]$Force,

    [Parameter(Mandatory = $false, HelpMessage = "Path to log file")]
    [string]$LogFile
)

# Set strict mode and error action preference
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Script-level variables
$script:LogFile = $LogFile
$script:StartTime = Get-Date

# Exit codes
$ExitCodes = @{
    Success = 0
    GeneralError = 1
    NetworkError = 2
    FileSystemError = 3
    ValidationError = 4
    UserCancelled = 5
}

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Info', 'Warning', 'Error', 'Success', 'Verbose')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Console output with colors
    switch ($Level) {
        'Info' { Write-Host $Message -ForegroundColor White }
        'Warning' { Write-Host $Message -ForegroundColor Yellow }
        'Error' { Write-Host $Message -ForegroundColor Red }
        'Success' { Write-Host $Message -ForegroundColor Green }
        'Verbose' { 
            if ($VerbosePreference -eq 'Continue') {
                Write-Host $Message -ForegroundColor Cyan
            }
        }
    }
    
    # File logging
    if ($script:LogFile) {
        try {
            Add-Content -Path $script:LogFile -Value $logMessage -ErrorAction Stop
        }
        catch {
            Write-Warning "Failed to write to log file: $_"
        }
    }
}

function Test-PathSecurity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        
        [Parameter(Mandatory = $false)]
        [string]$Purpose = "operation"
    )
    
    try {
        # Resolve to absolute path
        $resolvedPath = [System.IO.Path]::GetFullPath($Path)
        
        # Check for path traversal attempts
        if ($resolvedPath.Contains('..') -or $resolvedPath.Contains('~')) {
            throw "Path traversal detected in path: $Path"
        }
        
        # Validate path length
        if ($resolvedPath.Length -gt 260) {
            throw "Path too long: $resolvedPath"
        }
        
        Write-Log "Path validated for $Purpose : $resolvedPath" -Level Verbose
        return $resolvedPath
    }
    catch {
        Write-Log "Path validation failed for $Purpose : $($_.Exception.Message)" -Level Error
        throw
    }
}

function Get-WoWInstallPath {
    [CmdletBinding()]
    param()
    
    Write-Log "Auto-detecting World of Warcraft installation..." -Level Verbose
    
    $commonPaths = @(
        "${env:ProgramFiles(x86)}\World of Warcraft",
        "${env:ProgramFiles}\World of Warcraft",
        "C:\Program Files (x86)\World of Warcraft",
        "C:\Program Files\World of Warcraft",
        "D:\World of Warcraft",
        "E:\World of Warcraft"
    )
    
    foreach ($path in $commonPaths) {
        if (Test-Path $path -PathType Container) {
            $retailPath = Join-Path $path "_retail_"
            if (Test-Path $retailPath -PathType Container) {
                Write-Log "Found WoW installation at: $path" -Level Success
                return $path
            }
        }
    }
    
    # Check registry for Battle.net installation
    try {
        $regPath = "HKLM:\SOFTWARE\WOW6432Node\Blizzard Entertainment\World of Warcraft"
        if (Test-Path $regPath) {
            $installPath = Get-ItemPropertyValue -Path $regPath -Name "InstallPath" -ErrorAction SilentlyContinue
            if ($installPath -and (Test-Path $installPath -PathType Container)) {
                Write-Log "Found WoW installation via registry: $installPath" -Level Success
                return $installPath
            }
        }
    }
    catch {
        Write-Log "Registry check failed: $($_.Exception.Message)" -Level Verbose
    }
    
    throw "Could not auto-detect World of Warcraft installation. Please specify -WoWPath parameter."
}

function Get-LatestRelease {
    [CmdletBinding()]
    param()
    
    Write-Log "Fetching latest release information from GitHub..." -Level Info
    
    try {
        $apiUrl = "https://api.github.com/repos/imquazii/QuaziiUI/releases/latest"
        
        # Configure web request with security settings
        $webRequestParams = @{
            Uri = $apiUrl
            TimeoutSec = 30
            Headers = @{
                'User-Agent' = 'QuaziiUI-Installer/2.0'
                'Accept' = 'application/vnd.github.v3+json'
            }
            UseBasicParsing = $true
        }
        
        $releaseInfo = Invoke-RestMethod @webRequestParams
        
        if (-not $releaseInfo.zipball_url -or -not $releaseInfo.tag_name) {
            throw "Invalid release information received from GitHub API"
        }
        
        Write-Log "Latest release found: $($releaseInfo.tag_name)" -Level Success
        return @{
            ZipUrl = $releaseInfo.zipball_url
            TagName = $releaseInfo.tag_name
            PublishedAt = $releaseInfo.published_at
        }
    }
    catch [System.Net.WebException] {
        Write-Log "Network error fetching release information: $($_.Exception.Message)" -Level Error
        throw [System.Exception]::new("Failed to fetch release information from GitHub", $_.Exception)
    }
    catch {
        Write-Log "Error fetching release information: $($_.Exception.Message)" -Level Error
        throw
    }
}

function Get-FallbackRelease {
    [CmdletBinding()]
    param()
    
    Write-Log "Using fallback main branch download..." -Level Warning
    return @{
        ZipUrl = "https://github.com/imquazii/QuaziiUI/archive/refs/heads/main.zip"
        TagName = "main"
        PublishedAt = $null
    }
}

function Invoke-SecureDownload {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Url,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputPath,
        
        [Parameter(Mandatory = $false)]
        [int]$TimeoutSeconds = 300
    )
    
    Write-Log "Downloading from: $Url" -Level Verbose
    Write-Log "Saving to: $OutputPath" -Level Verbose
    
    try {
        # Validate and secure the output path
        $secureOutputPath = Test-PathSecurity -Path $OutputPath -Purpose "download"
        
        # Ensure parent directory exists
        $parentDir = Split-Path $secureOutputPath -Parent
        if (-not (Test-Path $parentDir -PathType Container)) {
            New-Item -Path $parentDir -ItemType Directory -Force | Out-Null
        }
        
        # Configure download with progress reporting
        $webRequestParams = @{
            Uri = $Url
            OutFile = $secureOutputPath
            TimeoutSec = $TimeoutSeconds
            Headers = @{
                'User-Agent' = 'QuaziiUI-Installer/2.0'
            }
            UseBasicParsing = $true
        }
        
        # Download with progress
        Write-Progress -Activity "Downloading QuaziiUI" -Status "Starting download..." -PercentComplete 0
        Invoke-WebRequest @webRequestParams
        Write-Progress -Activity "Downloading QuaziiUI" -Completed
        
        # Verify download
        if (-not (Test-Path $secureOutputPath)) {
            throw "Download verification failed - file not found at $secureOutputPath"
        }
        
        $fileSize = (Get-Item $secureOutputPath).Length
        if ($fileSize -eq 0) {
            throw "Download verification failed - file is empty"
        }
        
        Write-Log "Download completed successfully ($([math]::Round($fileSize / 1MB, 2)) MB)" -Level Success
        return $secureOutputPath
    }
    catch [System.Net.WebException] {
        Write-Log "Network error during download: $($_.Exception.Message)" -Level Error
        throw [System.Exception]::new("Download failed due to network error", $_.Exception)
    }
    catch {
        Write-Log "Download error: $($_.Exception.Message)" -Level Error
        throw
    }
}

function Expand-SecureArchive {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ArchivePath,
        
        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )
    
    Write-Log "Extracting archive: $ArchivePath" -Level Verbose
    
    try {
        # Validate paths
        $secureArchivePath = Test-PathSecurity -Path $ArchivePath -Purpose "archive extraction"
        $secureDestinationPath = Test-PathSecurity -Path $DestinationPath -Purpose "archive extraction"
        
        # Verify archive exists and is not empty
        if (-not (Test-Path $secureArchivePath)) {
            throw "Archive file not found: $secureArchivePath"
        }
        
        $archiveSize = (Get-Item $secureArchivePath).Length
        if ($archiveSize -eq 0) {
            throw "Archive file is empty: $secureArchivePath"
        }
        
        Write-Progress -Activity "Extracting QuaziiUI" -Status "Extracting files..." -PercentComplete 0
        Expand-Archive -Path $secureArchivePath -DestinationPath $secureDestinationPath -Force
        Write-Progress -Activity "Extracting QuaziiUI" -Completed
        
        Write-Log "Archive extracted successfully" -Level Success
    }
    catch [System.IO.InvalidDataException] {
        Write-Log "Archive corruption error: $($_.Exception.Message)" -Level Error
        throw [System.Exception]::new("Archive file is corrupted or invalid", $_.Exception)
    }
    catch {
        Write-Log "Extraction error: $($_.Exception.Message)" -Level Error
        throw
    }
}

function Find-ExtractedFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SearchPath,
        
        [Parameter(Mandatory = $true)]
        [string]$ReleaseTag
    )
    
    Write-Log "Searching for extracted QuaziiUI folder..." -Level Verbose
    
    try {
        $secureSearchPath = Test-PathSecurity -Path $SearchPath -Purpose "folder search"
        
        # Search patterns for different GitHub archive formats
        $searchPatterns = @(
            "imquazii-QuaziiUI-*",
            "QuaziiUI-$ReleaseTag",
            "QuaziiUI-main"
        )
        
        foreach ($pattern in $searchPatterns) {
            Write-Log "Searching for pattern: $pattern" -Level Verbose
            $folders = @(Get-ChildItem -Path $secureSearchPath -Directory | Where-Object { $_.Name -like $pattern })
            
            if ($folders.Count -gt 0) {
                $selectedFolder = $folders[0]
                Write-Log "Found extracted folder: $($selectedFolder.Name)" -Level Success
                return $selectedFolder.FullName
            }
        }
        
        throw "No extracted QuaziiUI folder found in $secureSearchPath"
    }
    catch {
        Write-Log "Folder search error: $($_.Exception.Message)" -Level Error
        throw
    }
}

function Confirm-DestructiveOperation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Operation,
        
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
    
    if ($Force) {
        Write-Log "Skipping confirmation due to -Force parameter" -Level Verbose
        return $true
    }
    
    Write-Host "`nWARNING: About to perform destructive operation" -ForegroundColor Yellow
    Write-Host "Operation: $Operation" -ForegroundColor Yellow
    Write-Host "Path: $Path" -ForegroundColor Yellow
    Write-Host "`nDo you want to continue? (Y/N): " -ForegroundColor Yellow -NoNewline
    
    $response = Read-Host
    return ($response -match '^[Yy]([Ee][Ss])?$')
}

function Remove-SecureItem {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        
        [Parameter(Mandatory = $false)]
        [string]$Description = "item"
    )
    
    try {
        if (-not (Test-Path $Path)) {
            Write-Log "$Description not found at $Path - skipping removal" -Level Verbose
            return
        }
        
        $securePath = Test-PathSecurity -Path $Path -Purpose "removal"
        
        if (-not (Confirm-DestructiveOperation -Operation "Remove $Description" -Path $securePath)) {
            Write-Log "User cancelled removal of $Description" -Level Warning
            throw [System.OperationCanceledException]::new("Operation cancelled by user")
        }
        
        Write-Log "Removing $Description : $securePath" -Level Info
        Remove-Item -Path $securePath -Recurse -Force
        Write-Log "$Description removed successfully" -Level Success
    }
    catch [System.OperationCanceledException] {
        throw
    }
    catch [System.UnauthorizedAccessException] {
        Write-Log "Access denied removing $Description : $($_.Exception.Message)" -Level Error
        throw [System.Exception]::new("Insufficient permissions to remove $Description", $_.Exception)
    }
    catch {
        Write-Log "Error removing $Description : $($_.Exception.Message)" -Level Error
        throw
    }
}

function Copy-AddonFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,
        
        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )
    
    Write-Log "Installing QuaziiUI files..." -Level Info
    
    try {
        # Validate paths
        $secureSourcePath = Test-PathSecurity -Path $SourcePath -Purpose "addon copy source"
        $secureDestinationPath = Test-PathSecurity -Path $DestinationPath -Purpose "addon copy destination"
        
        # Verify source exists and contains QuaziiUI folder
        $addonSourcePath = Join-Path $secureSourcePath "QuaziiUI"
        if (-not (Test-Path $addonSourcePath -PathType Container)) {
            throw "QuaziiUI folder not found in extracted archive at: $addonSourcePath"
        }
        
        # Create destination directory
        $parentDir = Split-Path $secureDestinationPath -Parent
        if (-not (Test-Path $parentDir -PathType Container)) {
            Write-Log "Creating AddOns directory: $parentDir" -Level Info
            New-Item -Path $parentDir -ItemType Directory -Force | Out-Null
        }
        
        # Create QuaziiUI destination directory
        if (-not (Test-Path $secureDestinationPath -PathType Container)) {
            New-Item -Path $secureDestinationPath -ItemType Directory -Force | Out-Null
        }
        
        # Copy files with progress
        Write-Progress -Activity "Installing QuaziiUI" -Status "Copying addon files..." -PercentComplete 50
        $copyParams = @{
            Path = Join-Path $addonSourcePath "*"
            Destination = $secureDestinationPath
            Recurse = $true
            Force = $true
        }
        Copy-Item @copyParams
        Write-Progress -Activity "Installing QuaziiUI" -Completed
        
        # Verify installation
        $copiedFiles = Get-ChildItem -Path $secureDestinationPath -Recurse -File
        if ($copiedFiles.Count -eq 0) {
            throw "No files were copied to destination"
        }
        
        Write-Log "QuaziiUI installed successfully ($($copiedFiles.Count) files copied)" -Level Success
    }
    catch [System.UnauthorizedAccessException] {
        Write-Log "Access denied during file copy: $($_.Exception.Message)" -Level Error
        throw [System.Exception]::new("Insufficient permissions to install addon files", $_.Exception)
    }
    catch {
        Write-Log "File copy error: $($_.Exception.Message)" -Level Error
        throw
    }
}

function Invoke-Cleanup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string[]]$PathsToClean = @()
    )
    
    Write-Log "Performing cleanup..." -Level Info
    
    foreach ($path in $PathsToClean) {
        if ($path -and (Test-Path $path)) {
            try {
                $securePath = Test-PathSecurity -Path $path -Purpose "cleanup"
                Remove-Item -Path $securePath -Recurse -Force -ErrorAction SilentlyContinue
                Write-Log "Cleaned up: $securePath" -Level Verbose
            }
            catch {
                Write-Log "Warning: Could not clean up $path : $($_.Exception.Message)" -Level Warning
            }
        }
    }
    
    Write-Log "Cleanup completed" -Level Success
}

function Main {
    [CmdletBinding()]
    param()
    
    $tempFiles = @()
    
    try {
        Write-Log "=== QuaziiUI Installer v2.0 Starting ===" -Level Info
        Write-Log "Start time: $($script:StartTime)" -Level Verbose
        
        # Determine WoW installation path
        if ($WoWPath) {
            $wowInstallPath = Test-PathSecurity -Path $WoWPath -Purpose "WoW installation"
            Write-Log "Using provided WoW path: $wowInstallPath" -Level Info
        } else {
            $wowInstallPath = Get-WoWInstallPath
        }
        
        $wowAddonsPath = Join-Path $wowInstallPath "_retail_\Interface\AddOns\QuaziiUI"
        $homeDirectory = $env:USERPROFILE
        
        # Remove existing installation if present
        if (Test-Path $wowAddonsPath) {
            Remove-SecureItem -Path $wowAddonsPath -Description "existing QuaziiUI installation"
        } else {
            Write-Log "No existing QuaziiUI installation found" -Level Info
        }
        
        # Get release information
        try {
            $releaseInfo = Get-LatestRelease
        }
        catch {
            Write-Log "Failed to get latest release, using fallback" -Level Warning
            $releaseInfo = Get-FallbackRelease
        }
        
        # Download release
        $zipFileName = "QuaziiUI-$($releaseInfo.TagName).zip"
        $zipPath = Join-Path $homeDirectory $zipFileName
        $tempFiles += $zipPath
        
        # Remove existing download if present
        if (Test-Path $zipPath) {
            Remove-SecureItem -Path $zipPath -Description "existing download file"
        }
        
        $downloadedPath = Invoke-SecureDownload -Url $releaseInfo.ZipUrl -OutputPath $zipPath
        
        # Extract archive
        Expand-SecureArchive -ArchivePath $downloadedPath -DestinationPath $homeDirectory
        
        # Find extracted folder
        $extractedPath = Find-ExtractedFolder -SearchPath $homeDirectory -ReleaseTag $releaseInfo.TagName
        $tempFiles += $extractedPath
        
        # Install addon files
        Copy-AddonFiles -SourcePath $extractedPath -DestinationPath $wowAddonsPath
        
        # Final verification
        if (-not (Test-Path $wowAddonsPath -PathType Container)) {
            throw "Installation verification failed - QuaziiUI directory not found"
        }
        
        $installedFiles = Get-ChildItem -Path $wowAddonsPath -Recurse -File
        Write-Log "=== Installation completed successfully! ===" -Level Success
        Write-Log "Files installed: $($installedFiles.Count)" -Level Info
        Write-Log "Installation path: $wowAddonsPath" -Level Info
        Write-Log "Release version: $($releaseInfo.TagName)" -Level Info
        
        if ($releaseInfo.PublishedAt) {
            Write-Log "Release date: $($releaseInfo.PublishedAt)" -Level Info
        }
        
        Write-Host "`nYou can now launch World of Warcraft and enable the QuaziiUI addon." -ForegroundColor Cyan
        
        $exitCode = $ExitCodes.Success
    }
    catch [System.OperationCanceledException] {
        Write-Log "Operation cancelled by user" -Level Warning
        $exitCode = $ExitCodes.UserCancelled
    }
    catch [System.Net.WebException] {
        Write-Log "Network error occurred: $($_.Exception.Message)" -Level Error
        $exitCode = $ExitCodes.NetworkError
    }
    catch [System.UnauthorizedAccessException] {
        Write-Log "Permission error occurred: $($_.Exception.Message)" -Level Error
        Write-Log "Try running as Administrator" -Level Warning
        $exitCode = $ExitCodes.FileSystemError
    }
    catch [System.IO.IOException] {
        Write-Log "File system error occurred: $($_.Exception.Message)" -Level Error
        $exitCode = $ExitCodes.FileSystemError
    }
    catch {
        Write-Log "Unexpected error occurred: $($_.Exception.Message)" -Level Error
        Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level Verbose
        $exitCode = $ExitCodes.GeneralError
    }
    finally {
        # Guaranteed cleanup
        Invoke-Cleanup -PathsToClean $tempFiles
        
        $endTime = Get-Date
        $duration = $endTime - $script:StartTime
        Write-Log "=== QuaziiUI Installer completed ===" -Level Info
        Write-Log "End time: $endTime" -Level Verbose
        Write-Log "Duration: $($duration.TotalSeconds) seconds" -Level Verbose
        Write-Log "Exit code: $exitCode" -Level Verbose
    }
    
    exit $exitCode
}

# Execute main function
Main