# 🧹 WinTrim 2.0

**Comprehensive Windows Cleanup Tool**

WinTrim 2.0 is a powerful, modular cleanup tool that goes beyond simple duplicate removal. It intelligently cleans your Windows system across multiple areas while keeping your system safe and functional. Like Marie Kondo for your entire Windows installation—keeping only what sparks joy (and functionality).

![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue.svg)
![Windows](https://img.shields.io/badge/Windows-10%2F11-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Version](https://img.shields.io/badge/version-2.0-brightgreen.svg)

## ✨ New in Version 2.0

### 🎯 Six Powerful Cleanup Modules

1. **Duplicate Programs** (Original WinTrim)
   - Removes old software updates and duplicate redistributables
   - Intelligent pattern matching and online version checking
   - Keeps different Visual C++ years and editions separate

2. **Temp Files Cleanup**
   - Cleans system and user temporary directories
   - Safe removal of old temporary files
   - Typical savings: 1-10 GB

3. **Windows Update Cache**
   - Clears Windows Update download cache
   - Safely stops and restarts update services
   - Typical savings: 5-20 GB

4. **Browser Cache Cleanup**
   - Supports Brave, Chrome, Edge, and Firefox
   - Clears code cache and regular cache
   - Typical savings: 500 MB - 5 GB

5. **System Cache Cleanup**
   - Thumbnail cache, shader cache, icon cache
   - Windows error reports and old logs
   - Prefetch files older than 7 days
   - Typical savings: 1-3 GB

6. **Advanced DISM Cleanup**
   - Component store analysis and cleanup
   - Superseded component removal
   - Most thorough but slowest (10-30 minutes)
   - Typical savings: 2-10 GB

### 🎮 Interactive Module Selection
- Run all modules or pick specific ones
- Easy-to-use menu interface
- Each module runs independently with its own confirmation

### 📊 Enhanced Reporting
- Comprehensive final statistics
- Per-module breakdown
- Time tracking
- Total space freed across all modules

## 🚀 Quick Start

### Prerequisites
- Windows 10 or Windows 11
- PowerShell 5.1 or later
- Administrator privileges

### Installation

1. **Download WinTrim**
   ```powershell
   # Clone the repository
   git clone https://github.com/kdsmith10000/WinTrim.git
   cd WinTrim
   ```

2. **Run as Administrator**
   - Right-click PowerShell
   - Select "Run as Administrator"
   - Navigate to the WinTrim directory
   - Run the script:
   ```powershell
   .\WinTrim.ps1
   ```

## 📖 Usage

### Basic Usage

Simply run the script and follow the interactive prompts:

```powershell
.\WinTrim.ps1
```

The script will:
1. ✅ Prompt to create a system restore point (recommended!)
2. 📋 Show available cleanup modules
3. 🎯 Let you select which modules to run (or run all)
4. 🔍 Scan and analyze each selected module
5. 💾 Display estimated disk space savings per module
6. ⚠️ Ask for confirmation before each cleanup operation
7. 🗑️ Perform cleanup operations
8. 📊 Show comprehensive final statistics

### Interactive Module Selection

```
Available modules:
  1. Duplicate Programs - Remove old software updates (original WinTrim)
  2. Temp Files - Clean system and user temp directories
  3. Windows Update Cache - Clear update download cache
  4. Browser Cache - Clear browser caches (Brave, Chrome, Edge, Firefox)
  5. System Cache - Clean thumbnails, logs, error reports, shader cache
  6. Advanced DISM - Component store cleanup (slow but thorough)
  A. All modules
  Q. Quit

Enter modules to run (e.g., '1,2,3' or 'A' for all):
```

### What Gets Cleaned by Module

**Module 1: Duplicate Programs**
- ✅ Visual C++ Redistributables (old versions of same year/edition)
- ✅ Microsoft .NET Runtime (outdated versions)
- ✅ Microsoft Edge Updates (old update helpers)
- ✅ MongoDB, Java, Adobe updates (old versions)
- ❌ NEVER removes different Visual C++ years (coexist by design)
- ❌ NEVER removes Python components (part of one installation)
- ❌ NEVER removes single installations

**Module 2-6: System Cleanup**
- Temporary files (system and user)
- Windows Update download cache
- Browser caches (all major browsers)
- Thumbnail, shader, and icon caches
- Windows error reports and old logs
- Superseded Windows components (DISM)

## 💡 Example Output

```
========================================
  WinTrim v2.0
  Comprehensive Windows Cleanup
========================================

Available modules:
  1. Duplicate Programs - Remove old software updates (original WinTrim)
  2. Temp Files - Clean system and user temp directories
  3. Windows Update Cache - Clear update download cache
  4. Browser Cache - Clear browser caches (Brave, Chrome, Edge, Firefox)
  5. System Cache - Clean thumbnails, logs, error reports, shader cache
  6. Advanced DISM - Component store cleanup (slow but thorough)
  A. All modules
  Q. Quit

Enter modules to run (e.g., '1,2,3' or 'A' for all): A

Modules to run:
  ✓ Duplicate Programs
  ✓ Temp Files
  ✓ Windows Update Cache
  ✓ Browser Cache
  ✓ System Cache
  ✓ Advanced DISM

========================================
  MODULE: Temporary Files Cleanup
========================================

Scanning temp directories...
Total potential space to free: 3.45 GB

Proceed with temp files cleanup? (Y/N) Y

Temp Cleanup Results:
  Files removed: 15,234
  Space freed: 3.45 GB

... [other modules run] ...

========================================
  WinTrim Cleanup Complete!
========================================

FINAL STATISTICS:
-------------------
Total space freed: 24.67 GB
Programs removed: 8
Temp files removed: 15,234
Time elapsed: 12m 34s

CLEANUP BY MODULE:
-------------------
Temp Files: 15,234 files removed
Windows Update: 8.23 GB
Browser Cache: 4.12 GB
System Cache: 2.34 GB
Duplicate Programs: 8 removed
DISM Cleanup: 6.53 GB
```

## 🔧 Advanced Configuration

### Execution Policy

If you encounter execution policy errors:

```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
.\WinTrim.ps1
```

### Custom Patterns

You can modify the `$patterns` array in the script to add your own software patterns:

```powershell
$patterns = @(
    "Microsoft Edge",
    "Your Custom Pattern"
)
```

## 📊 Statistics & Reporting

After cleanup, WinTrim provides:

```
REMOVAL STATISTICS:
-------------------
Programs targeted for removal: 8
Successfully removed: 8
Failed to remove: 0

DISK SPACE SAVINGS:
-------------------
Total disk space freed: 102.95 MB

ESTIMATED MEMORY SAVINGS:
-------------------------
Potential RAM savings: ~600 MB
(Old update services no longer running)

ADDITIONAL BENEFITS:
--------------------
- Cleaner Add/Remove Programs list
- Reduced registry clutter
- Faster system scans and updates
```

## 🛡️ Safety & Security

### Security Features

- ✅ **Requires Administrator Privileges**: Prevents unauthorized execution
- ✅ **No Internet Dependencies**: Works offline (uses cached version data as fallback)
- ✅ **No Data Collection**: Completely private, nothing sent anywhere
- ✅ **Open Source**: Fully auditable code

### Security Audit

WinTrim has been designed with security in mind:

- **No arbitrary code execution**: Only uses uninstall strings from registry
- **Input validation**: All registry data is validated before use
- **Safe defaults**: Conservative approach, keeps items when in doubt
- **Restore point integration**: Easy rollback if needed

## ⚠️ Important Notes

### Before Running

1. **Create a backup** of important data
2. **Close all programs** before running
3. **Read the preview** carefully before confirming removal
4. **Accept the restore point** prompt (strongly recommended)

### Known Limitations

- Some programs may fail to uninstall silently (exit code 1603)
- Requires manual intervention for programs with custom uninstallers
- Online version checking requires internet connection (falls back to cached data)

## 🐛 Troubleshooting

### "Script not recognized" error
```powershell
# Use full path
C:\path\to\WinTrim.ps1
```

### "Execution policy" error
```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
```

### "Not running as administrator"
Right-click PowerShell → "Run as Administrator"

### Items failed to remove
Check the exit code in the output. Common codes:
- `1603`: Installation failed (dependency issue)
- `1605`: Product not found (already removed)

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

### Areas for Contribution

- 🌐 Additional online version sources
- 🎯 More software patterns
- 🌍 Internationalization
- 📝 Documentation improvements
- 🐛 Bug fixes

## 📜 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🙏 Acknowledgments

- Inspired by the need for safer Windows cleanup tools
- Built with feedback from the Windows power user community
- Thanks to all contributors and testers

## 📞 Support

- **Issues**: [GitHub Issues](https://github.com/kdsmith10000/WinTrim/issues)
- **Discussions**: [GitHub Discussions](https://github.com/kdsmith10000/WinTrim/discussions)

## ⚡ Roadmap

### ✅ Completed (v2.0)
- [x] Modular cleanup system
- [x] Temp files cleanup
- [x] Windows Update cache cleanup
- [x] Browser cache cleanup
- [x] System cache cleanup
- [x] DISM component cleanup
- [x] Comprehensive statistics and reporting

### 🔮 Future Enhancements
- [ ] GUI version
- [ ] Scheduled cleanup tasks
- [ ] Export/import cleanup profiles
- [ ] Detailed HTML/JSON reports
- [ ] Per-module rollback
- [ ] Portable version (no installation required)
- [ ] Recycle Bin cleanup module
- [ ] Windows.old removal module
- [ ] Driver store cleanup
- [ ] More granular control per cleanup type

---

**Made with ❤️ for the Windows community**

*WinTrim - Keep your Windows installation lean and clean!*
