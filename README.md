# WinGet Application Manager

<div align="center">

![Windows](https://img.shields.io/badge/Windows-0078D6?style=for-the-badge&logo=windows&logoColor=white)
![PowerShell](https://img.shields.io/badge/PowerShell-5391FE?style=for-the-badge&logo=powershell&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green.svg?style=for-the-badge)
![Version](https://img.shields.io/badge/version-1.0.0-blue.svg?style=for-the-badge)
![Status](https://img.shields.io/badge/status-active-success.svg?style=for-the-badge)

**A modern, professional GUI for managing Windows applications with WinGet**

[Features](#-features) • [Installation](#-installation) • [Usage](#-usage) • [Screenshots](#-screenshots) • [Contributing](#-contributing)

---

</div>

## 📋 Executive Summary

**WinGet Application Manager** transforms Microsoft's command-line WinGet package manager into a beautiful, user-friendly desktop application. Built entirely in PowerShell with WPF, it provides enterprise-grade package management with zero dependencies beyond Windows 10/11.

### 🎯 Why Use This?

- **🖱️ Click Instead of Type** - Manage hundreds of applications through an intuitive interface
- **📊 Visual Overview** - See all installed applications, available updates, and statuses at a glance
- **🔄 Bulk Operations** - Update, install, or remove multiple applications simultaneously
- **💾 Import/Export** - Save and restore application configurations across machines
- **🎨 Modern UI** - Beautiful dark/light themes with responsive design
- **🛡️ Professional** - Production-ready with comprehensive error handling and logging

---

## ✨ Features

### 🔍 **Package Management**
- 📦 **Browse** all installed applications with version info
- 🔄 **Auto-detect** available updates with one-click updates
- 🔍 **Search** WinGet repository (1000+ applications)
- ⚡ **Bulk operations** - Update/install/uninstall/repair multiple apps at once
- 🔧 **Repair** broken installations with WinGet repair feature
- 📌 **Status tracking** for each operation

### 💾 **Import/Export**
- 📤 **Export** your application list to JSON
- 📥 **Import** application lists from JSON files
- 🎯 **Selective** import - choose which apps to install
- 📁 **Browse** for JSON files or paste file paths directly
- 💻 **Machine migration** - replicate setups across computers
- 🏢 **Team deployments** - standardize application stacks

### 📊 **Activity Logging**
- 📝 **Real-time** operation logs with timestamps
- 🎨 **Color-coded** messages (Success/Error/Warning/Info)
- 📋 **Export** logs as Markdown or plain text
- 🔍 **Detailed** error messages with 100+ official WinGet codes
- 📅 **Timestamp** every action for audit trails

### 🎨 **User Experience**
- 🌓 **Dark/Light** themes with perfect readability
- 🚀 **Fast** - Async operations never block the UI
- 📁 **Browse** and paste support for file paths
- 🔔 **Status** indicators for all operations
- ❌ **Cancel** long-running operations safely
- ✅ **WinGet check** validates installation at startup

### 🛡️ **Reliability**
- ✅ **100+ WinGet error codes** mapped with official Microsoft descriptions
- 🔒 **Safe** - No destructive operations without confirmation
- 🔄 **Auto-close** applications before updates
- 🚦 **Progress** tracking with per-app status
- 📊 **Success rates** - Know exactly what worked/failed
- 🔍 **Version check** - Validates WinGet availability at startup

---

## 🚀 Installation

### Prerequisites

- ✅ **Windows 10** (1809+) or **Windows 11**
- ✅ **PowerShell 5.1+** (pre-installed on Windows)
- ✅ **WinGet** ([Install Guide](https://learn.microsoft.com/en-us/windows/package-manager/winget/))
- ✅ **.NET Framework 4.7.2+** (pre-installed on Windows 10/11)

### Quick Start

1. **Download** the latest release:
   ```powershell
   # Clone the repository
   git clone https://github.com/ChrisMunnPS/WinGet-Application-Manager.git
   cd WinGet-Application-Manager
   ```

2. **Run** the application:
   ```powershell
   # Right-click WingetManager.ps1 → Run with PowerShell
   # OR from PowerShell:
   .\WingetManager.ps1
   ```

3. **First Launch**:
   - The app automatically checks for WinGet
   - If WinGet is missing, you'll get installation instructions
   - Once verified, you're ready to go!

### 🔧 Optional: Execution Policy

If you encounter execution policy errors:

```powershell
# Run PowerShell as Administrator, then:
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

---

## 📖 Usage

### Package Manager Tab

**View Installed Applications:**
1. Open the **Package Manager** tab
2. Click **🔄 Refresh Installed** to load your applications
3. Applications with updates are **auto-selected**

**Update Applications:**
1. Select packages to update (or keep auto-selection)
2. Click **⬆ Update Selected**
3. Monitor progress in **Activity Log** tab
4. Refresh to see updated versions

**Install New Applications:**
1. Enter application name in search box
2. Click **🔍 Search**
3. Select packages to install
4. Click **⬇ Install Selected**

**Repair Installations:**
1. Select broken installations
2. Click **🔧 Repair Selected**
3. WinGet attempts to fix the installation

### Import/Export Tab

**Export Your Setup:**
1. Go to **Import / Export** tab
2. Click **⬆ Export Applications**
3. Choose save location
4. JSON file contains all installed apps

**Import Configuration:**
1. Click **📁 Browse** (or paste JSON file path in the text box)
2. Click **"⬇ Load / Install"** to load packages
3. Grid shows all packages with checkboxes (all checked by default)
4. Use **Select All** / **Deselect All** buttons or click individual checkboxes
5. Click **"⬇ Install Checked"** to install selected packages
6. Monitor installation in Activity Log

**Migration Workflow:**
```
Old PC: Export → USB Drive → New PC: Import → Done!
```

### Activity Log

**View Operations:**
- All operations logged with timestamps
- Color-coded: 🟢 Success | 🔴 Error | 🟡 Warning | ⚪ Info

**Copy/Export Logs:**
1. Click **📋 Copy Log**
2. Choose format (Markdown or Text)
3. Save for documentation/support

**Clear Log:**
- Click **🗑 Clear** to remove old entries

---

## 🐦‍🔥 Some Screenshots

Main Page - Light Mode
![Main Page – Light Mode](https://github.com/ChrisMunnPS/WinGet-Application-Manager/blob/main/Screenshots/Light%20Mode/1a%20-%20Main%20Page.png?raw=true)

Main Page - Dark Mode
![Main Page – Dark Mode](https://github.com/ChrisMunnPS/WinGet-Application-Manager/blob/main/Screenshots/Dark%20Mode/1a%20-%20Main%20Page.png?raw=true)

![Import Selection - Light Mode](https://raw.githubusercontent.com/ChrisMunnPS/WinGet-Application-Manager/main/Screenshots/Light%20Mode/1c%20-%20Import%20Selection.png)

![Update, Repair, Uninstall - Light Mode](https://raw.githubusercontent.com/ChrisMunnPS/WinGet-Application-Manager/main/Screenshots/Light%20Mode/2%20-%20Update%2C%20Repair%2C%20Uninstall.png)

![Activity Log - Light Mode](https://raw.githubusercontent.com/ChrisMunnPS/WinGet-Application-Manager/main/Screenshots/Light%20Mode/4%20-%20Activity%20Log.png)

![Import Selection - Dark Mode](https://raw.githubusercontent.com/ChrisMunnPS/WinGet-Application-Manager/main/Screenshots/Dark%20Mode/1c%20-%20Import%20Selection.png)

![Update, Repair, Uninstall - Dark Mode](https://raw.githubusercontent.com/ChrisMunnPS/WinGet-Application-Manager/main/Screenshots/Dark%20Mode/2%20-%20Update%2C%20Repair%2C%20Uninstall.png)

![Activity Log - Dark Mode](https://raw.githubusercontent.com/ChrisMunnPS/WinGet-Application-Manager/main/Screenshots/Dark%20Mode/4%20-%20Activity%20Log.png)

---



## 🎯 Key Features Explained

### 🔧 Repair Feature
The **Repair** button uses WinGet's built-in repair functionality to fix broken installations:
- Reinstalls files without uninstalling
- Fixes registry entries
- Repairs file associations
- Works with MSI, MSIX, and EXE installers

### 📁 Drag & Drop
Simply **drag any JSON file** into the Import/Export tab:
- Instant file path population
- Automatic validation
- Visual feedback
- No more browsing for files!

### ✅ WinGet Validation
At startup, the app:
- Checks if WinGet is installed
- Verifies WinGet version
- Offers installation help if missing
- Won't run without WinGet (safety first!)

### 📊 Error Code Mapping
100+ WinGet error codes mapped to human-readable messages:
```
Instead of: "Exit code: -1978334975"
You see: "Application is currently running. Exit the application then try again."
```

---

## ⚙️ Configuration

### Settings File
Settings are automatically saved to:
```
%LOCALAPPDATA%\WingetAppMgr\settings.json
```

**Stored Preferences:**
- 🎨 Theme (Dark/Light)
- 📁 Last export path
- 🪟 Window size and position

### Export File Format

```json
{
  "SchemaVersion": "1.0",
  "CreatedDate": "2026-02-22T15:30:00",
  "Computer": "YOUR-PC",
  "Sources": [
    {"Name": "winget", "Argument": "https://..."}
  ],
  "Packages": [
    {
      "PackageIdentifier": "Microsoft.VisualStudioCode",
      "Version": "1.85.0",
      "Source": "winget"
    }
  ]
}
```

---

## 🔧 Troubleshooting

### WinGet Not Found
**Problem:** "WinGet is not installed"

**Solution:**
1. Install from [Microsoft Store](https://www.microsoft.com/store/productId/9NBLGGH4NNS1)
2. Or download from [GitHub](https://github.com/microsoft/winget-cli/releases)
3. Restart the application

### Updates Fail
**Problem:** "Application is currently running"

**Solution:**
1. Close the application completely
2. Check system tray for hidden instances
3. Try running PowerShell as Administrator

### Admin Rights Required
**Problem:** "Command requires administrator privileges"

**Solution:**
1. Right-click PowerShell
2. Select "Run as Administrator"
3. Launch the application again

### Import Fails
**Problem:** "Could not parse JSON file"

**Solution:**
1. Verify JSON file is valid
2. Ensure it was exported from this app or WinGet
3. Check file isn't corrupted
4. Try using Browse button to select the file

---

## 🚦 Roadmap

- [x] **WinGet Validation** - Check at startup
- [x] **Repair Feature** - Fix broken installations
- [x] **Import Selection** - Choose what to install
- [ ] **Scheduled Updates** - Automatic update checks
- [ ] **Package Details** - View description, homepage, license
- [ ] **Source Management** - Add/remove package sources


---

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

### Development Setup

1. Fork the repository
2. Create your feature branch
   ```bash
   git checkout -b feature/AmazingFeature
   ```
3. Commit your changes
   ```bash
   git commit -m 'Add some AmazingFeature'
   ```
4. Push to the branch
   ```bash
   git push origin feature/AmazingFeature
   ```
5. Open a Pull Request

---

## 📜 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## 👤 Author

**Christopher Munn**

- 🐙 GitHub: [@ChrisMunnPS](https://github.com/ChrisMunnPS)
- 🌐 Website: [ChrisMunnPS.github.io](https://ChrisMunnPS.github.io)
- 💼 LinkedIn: [Chris Munn](https://www.linkedin.com/in/chrismunn/)

---

## 🙏 Acknowledgments

- **Microsoft** - For creating WinGet and making it open-source
- **WinGet Community** - For maintaining the package repository
- **PowerShell Community** - For excellent WPF/XAML resources
- **You** - For using and improving this tool!

---

<div align="center">

**[⬆ Back to Top](#winget-application-manager)**

Made with ❤️ using PowerShell and WPF

![GitHub stars](https://img.shields.io/github/stars/ChrisMunnPS/WinGet-Application-Manager?style=social)
![GitHub forks](https://img.shields.io/github/forks/ChrisMunnPS/WinGet-Application-Manager?style=social)

</div>
