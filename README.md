# 🚀 WinGet Application Manager

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue.svg)](https://github.com/PowerShell/PowerShell)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)
[![WinGet](https://img.shields.io/badge/WinGet-Required-orange.svg)](https://github.com/microsoft/winget-cli)

> A modern, feature-rich GUI for managing Windows applications with WinGet

![WinGet Application Manager](https://img.shields.io/badge/Status-Production%20Ready-brightgreen)

## 📋 Executive Summary

**WinGet Application Manager** is a powerful Windows application that provides a modern graphical interface for Microsoft's WinGet package manager. Designed for both power users and IT professionals, it streamlines the process of installing, updating, and managing Windows applications through an intuitive interface.

### ✨ Key Features

- 🎨 **Modern Dark/Light Theme** - Professional interface with seamless theme switching
- 📦 **Package Management** - Install, update, and uninstall applications with checkboxes
- 🔍 **Smart Search** - Search installed packages or discover new apps from the WinGet repository
- 📤 **Export/Import** - Backup and restore your application configurations
- 🎯 **Batch Operations** - Update or install multiple applications simultaneously
- 📊 **Real-time Progress** - Live progress tracking with detailed activity logs
- ⚡ **Auto-Update Detection** - Automatically identifies packages with available updates
- 🎛️ **Selective Installation** - Choose exactly which applications to install during import
- 🛑 **Cancellation Support** - Stop operations safely at any time

---

## 🎯 Quick Start

### Prerequisites

- ✅ **Windows 10/11** (Build 1809 or later)
- ✅ **PowerShell 5.1** or higher
- ✅ **WinGet** (App Installer from Microsoft Store)

### Installation

1. **Download the script:**
   ```powershell
   # Clone the repository
   git clone https://github.com/ChrisMunnPS/WinGet-Application-Manager.git
   cd WinGet-Application-Manager
   ```

2. **Run the application:**
   ```powershell
   .\WingetManager.ps1
   ```

   Or right-click the script and select **"Run with PowerShell"**

---

## 💡 Features Overview

### 🎛️ Package Manager Tab

Manage all your installed applications from a single interface:

- ✅ **Auto-select packages with updates** - Packages needing updates are automatically checked
- 🔄 **Real-time refresh** - Update your package list on demand
- 🔍 **Search & Filter** - Find packages by name or ID
- ⬆️ **Batch Updates** - Update multiple packages with one click
- 🗑️ **Batch Uninstall** - Remove multiple applications efficiently
- 📥 **Install from Repository** - Search WinGet and install new applications

**Display Information:**
- Package name and ID
- Installed version
- Available version (if update exists)
- Current status (Installed/Update Available)
- Source repository

### 📤 Import/Export Tab

Backup and restore your application configurations:

**Export Features:**
- 📋 **One-click export** - Save all installed applications to JSON
- 📁 **Auto-naming** - Files named as `WingetApps_HOSTNAME_DATE.json`
- 📊 **Visual progress** - See each application as it's exported
- ✅ **Validation** - Ensures successful export before completion

**Import Features:**
- 📥 **Selective installation** - Choose which apps to install via checkboxes
- ✅ **All selected by default** - Quick restore with option to customize
- 📊 **Live progress tracking** - Watch installations in real-time
- ⚠️ **Error reporting** - Clear indication of successes and failures

### 📝 Activity Log Tab

Monitor all operations with detailed logging:

- 🕐 **Timestamped entries** - Track when each action occurred
- 🎨 **Color-coded messages** - Easy identification of errors, warnings, and successes
- 💾 **Export logs** - Save logs for troubleshooting or record-keeping
- 🗑️ **Clear logs** - Start fresh when needed
- 📜 **Auto-scroll** - Always see the latest activity

### ℹ️ About Section

Quick access to:
- 📦 Version information
- 👤 Author details
- 🔗 GitHub repository
- 💼 LinkedIn profile
- 🌐 Website link

---

## 🖥️ Screenshots

### Package Manager Interface
```
┌──────────────────────────────────────────────────────┐
│ [🔄 Refresh] [✓ All] [✗ None]  55 installed | 2 updates | 2 selected │
│ [Search box...................] [🔍 Search] [✕ Clear] │
├───┬────────┬──────┬─────┬─────┬──────┬────────┤
│☑ │ Name   │ ID   │Inst │Avail│Status│Source  │
├───┼────────┼──────┼─────┼─────┼──────┼────────┤
│[✓]│Discord │Disc..│1.0.9│1.0.x│Update│winget  │
│[✓]│Git     │Git...│2.43 │2.44 │Update│winget  │
│[ ]│Node.js │Node..│20.11│     │Inst. │winget  │
└───┴────────┴──────┴─────┴─────┴──────┴────────┘
          [⬆ UPDATE] [⬇ INSTALL] [🗑 UNINSTALL]
```

### Import with Selective Installation
```
┌──────────────────────────────────────────────────────┐
│ Select applications to install:                       │
├───┬────────────┬─────────────────┬─────────┬────────┤
│☑ │ Name       │ ID              │ Version │ Status │
├───┼────────────┼─────────────────┼─────────┼────────┤
│[✓]│ Discord    │ Discord.Discord │ 1.0.9225│ Ready  │
│[✓]│ Git        │ Git.Git         │ 2.43.0  │ Ready  │
│[ ]│ Node.js    │ OpenJS.NodeJS   │ 20.11.0 │ Skip   │
└───┴────────────┴─────────────────┴─────────┴────────┘
             [⬇ Install Selected (2)]
```

---

## 🔧 Usage Examples

### Basic Operations

**Update Applications:**
```powershell
1. Go to Package Manager tab
2. Packages with updates are auto-checked
3. Click "⬆ Update Selected"
4. Monitor progress in Activity Log
```

**Search and Install:**
```powershell
1. Type application name (e.g., "firefox")
2. Press Enter or click "🔍 Search"
3. If not installed, searches WinGet repository
4. Check desired packages
5. Click "⬇ Install Selected"
```

**Export Configuration:**
```powershell
1. Go to Import/Export tab
2. Click "⬆ Export Applications"
3. Choose save location
4. File saved as WingetApps_PC_2026-02-20.json
```

**Selective Import:**
```powershell
1. Go to Import/Export tab
2. Browse to JSON file (or drag & drop)
3. Click "⬇ Import Applications"
4. Uncheck apps you don't want
5. Click "⬇ Install Selected"
6. Monitor progress in Activity Log
```

---

## 🎨 Customization

### Theme Switching
- Click the **theme button** in the header (🌙 Dark Mode / ☀ Light Mode)
- Settings persist between sessions
- Instant theme switching without restart

### Data Persistence
- Recent file paths remembered
- Theme preference saved
- No data loss when switching tabs
- Search results preserved until new search

---

## 📊 Technical Details

### Architecture
- **Language**: PowerShell 5.1+
- **UI Framework**: WPF (Windows Presentation Foundation)
- **Threading**: Runspace-based async operations
- **Storage**: JSON-based settings and exports

### Features
- ✅ Non-blocking UI during operations
- ✅ Real-time progress updates
- ✅ Graceful cancellation support
- ✅ Comprehensive error handling
- ✅ Drag-and-drop file support
- ✅ Keyboard shortcuts (Ctrl+E, Ctrl+I)

### Requirements
```powershell
# Verify WinGet installation
winget --version

# Verify PowerShell version
$PSVersionTable.PSVersion
```

---

## 🐛 Troubleshooting

### Common Issues

**WinGet not found:**
```powershell
# Install App Installer from Microsoft Store
# Or download from: https://github.com/microsoft/winget-cli/releases
```

**PowerShell execution policy:**
```powershell
# Run PowerShell as Administrator
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

**Updates not detected:**
```powershell
# Refresh WinGet sources
winget source update
```

**Script won't run:**
```powershell
# Unblock the script
Unblock-File -Path .\WingetManager.ps1
```

---

## 🤝 Contributing

Contributions are welcome! Please feel free to submit issues, feature requests, or pull requests.

### Development Setup
```powershell
# Clone the repository
git clone https://github.com/ChrisMunnPS/WinGet-Application-Manager.git

# Create a feature branch
git checkout -b feature/amazing-feature

# Make your changes and test thoroughly

# Commit your changes
git commit -m "Add amazing feature"

# Push to your fork
git push origin feature/amazing-feature

# Open a Pull Request
```

---

## 📜 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## 👤 Author

**Christopher Munn**

- 🌐 Website: [https://ChrisMunnPS.github.io](https://ChrisMunnPS.github.io)
- 💼 LinkedIn: [Chris Munn](https://www.linkedin.com/in/chris-munn)
- 🐙 GitHub: [@ChrisMunnPS](https://github.com/ChrisMunnPS)

---

## 🙏 Acknowledgments

- Microsoft WinGet team for the excellent package manager
- PowerShell community for invaluable resources
- All contributors and users providing feedback

---

## 📈 Roadmap

- [ ] Multi-language support
- [ ] Package comparison between systems
- [ ] Scheduled automatic updates
- [ ] Custom package sources
- [ ] Export to other formats (CSV, XML)
- [ ] Restore point creation before operations
- [ ] Notification system for updates

---

## 💬 Support

If you find this tool helpful, please consider:
- ⭐ Starring the repository
- 🐛 Reporting bugs
- 💡 Suggesting features
- 📢 Sharing with others

---

<div align="center">

**Made with ❤️ by Christopher Munn**

[⬆ Back to Top](#-winget-application-manager)

</div>
