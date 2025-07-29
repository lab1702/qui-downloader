# QuaziiUI Downloader

A PowerShell script that automatically downloads and installs the latest QuaziiUI addon for World of Warcraft.

## Quick Start

To run the script with default options without first dowloading it you can do this:

```powershell
Invoke-Expression (Invoke-WebRequest -Uri "https://raw.githubusercontent.com/lab1702/qui-downloader/refs/heads/main/quazii-downloader.ps1").Content
```

## Features

- **Automatic WoW Detection**: Detects your World of Warcraft installation automatically
- **Latest Release Download**: Fetches the most recent QuaziiUI release from GitHub
- **Security Focused**: Includes path validation, secure downloads, and comprehensive error handling
- **Progress Tracking**: Shows download and installation progress
- **Flexible Logging**: Optional file logging with console output
- **Safe Installation**: Removes old versions before installing new ones
- **Fallback Support**: Uses main branch if latest release fails

## Requirements

- PowerShell 5.1 or higher
- World of Warcraft installed
- Internet connection

## Usage

### Basic Installation
```powershell
.\quazii-downloader.ps1
```

### Custom WoW Path
```powershell
.\quazii-downloader.ps1 -WoWPath "D:\Games\World of Warcraft"
```

### Skip Confirmations
```powershell
.\quazii-downloader.ps1 -Force
```

### Enable Logging
```powershell
.\quazii-downloader.ps1 -LogFile "C:\Logs\quazii-install.log"
```

### Combined Parameters
```powershell
.\quazii-downloader.ps1 -WoWPath "D:\Games\World of Warcraft" -Force -LogFile "install.log"
```

## Parameters

| Parameter | Type | Description |
|-----------|------|-------------|
| `-WoWPath` | String | Custom path to World of Warcraft installation directory |
| `-Force` | Switch | Skip confirmation prompts for destructive operations |
| `-LogFile` | String | Path to log file for detailed operation logging |

## Exit Codes

| Code | Meaning |
|------|---------|
| 0 | Success |
| 1 | General Error |
| 2 | Network Error |
| 3 | File System Error |
| 4 | Validation Error |
| 5 | User Cancelled |

## Security Features

- Path traversal protection
- Secure file operations
- Download verification
- Archive integrity checks
- Permission validation

## Troubleshooting

### Common Issues

**"Could not auto-detect World of Warcraft installation"**
- Use the `-WoWPath` parameter to specify your WoW installation directory
- Ensure World of Warcraft is properly installed

**"Access denied" errors**
- Run PowerShell as Administrator
- Check file/folder permissions

**Network errors**
- Verify internet connection
- Check firewall settings
- Try again later if GitHub is experiencing issues

**Archive extraction failures**
- Ensure sufficient disk space
- Check for antivirus interference

## What It Does

1. Detects your World of Warcraft installation (or uses provided path)
2. Removes any existing QuaziiUI installation
3. Downloads the latest QuaziiUI release from GitHub
4. Extracts the downloaded archive
5. Installs QuaziiUI to the correct AddOns directory
6. Cleans up temporary files
7. Provides installation confirmation

## License

This installer script is provided as-is for the QuaziiUI addon community.

## Author

Snackington
