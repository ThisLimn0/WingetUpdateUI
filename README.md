# WingetUpdateUI

A lightweight PowerShell GUI for Windows Package Manager updates.

## Quick Start

```powershell
powershell -ExecutionPolicy Bypass -File .\src\WingetUI.ps1
```

## Requirements

- Windows 10/11
- winget (Windows Package Manager)

## Features

- **Package Management**
  - View available updates
  - Select/deselect packages for update
  - Batch update selected packages
  - Automatic retry with elevated privileges

- **Detail Panel**
  - Package information (Publisher, Author, Description)
  - License and support links
  - Version info and release notes
  - Installer download URL

- **Installation Options**
  - Standard mode (default)
  - Silent mode (`--silent`)
  - Interactive mode (`--interactive`)
  - Force installation (`--force`)

- **Network Settings**
  - HTTP/HTTPS/SOCKS proxy support
  - Proxy authentication

- **User Interface**
  - Dark mode
  - 6 languages (DE, EN, FR, ES, IT, NL)
  - Real-time progress tracking
  - Activity log

## Version

1.0.0
