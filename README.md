# 🧹 WinTrim

**Smart Windows Update & Redistributable Cleanup Tool**

WinTrim intelligently removes duplicate and outdated software updates while keeping your system safe and functional. It's like Marie Kondo for your Windows installation—keeping only what sparks joy (and functionality).

![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue.svg)
![Windows](https://img.shields.io/badge/Windows-10%2F11-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)

## ✨ Features

### 🎯 Smart Detection
- **Intelligent Pattern Matching**: Finds duplicate Visual C++ redistributables, .NET components, and other updates
- **Online Version Checking**: Compares installed versions against latest releases from official sources
- **Dependency-Aware**: Protects essential components from accidental removal

### 🛡️ Safety First
- **Automatic Restore Points**: Creates system restore point before making changes
- **Protected Installations**: Never removes single installations or essential components
- **Preview Mode**: Shows what will be removed before you confirm
- **Coexistence Support**: Keeps different Visual C++ years (2005, 2008, 2010, etc.) as they're needed by different programs

### 📊 Comprehensive Reporting
- Disk space savings estimation
- Memory savings projection
- Success/failure statistics
- Detailed removal logs

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

Simply run the script and follow the prompts:

```powershell
.\WinTrim.ps1
```

The script will:
1. ✅ Prompt to create a system restore point (recommended!)
2. 🔍 Scan your installed programs
3. 🌐 Check online for latest versions
4. 📋 Show you what will be removed and kept
5. 💾 Display estimated disk space savings
6. ⚠️ Ask for confirmation before removing anything
7. 🗑️ Remove old versions
8. 📊 Show detailed statistics

### What Gets Cleaned

WinTrim targets these types of duplicates:

- ✅ **Visual C++ Redistributables** (old versions of same year/edition)
- ✅ **Microsoft .NET Runtime** (outdated versions)
- ✅ **Microsoft Edge Updates** (old update helpers)
- ✅ **MongoDB Updates** (old versions)
- ✅ **Java Updates** (outdated versions)
- ✅ **Adobe Reader/Acrobat** (old updates)

### What Stays Safe

WinTrim **NEVER** removes:

- ❌ **Different Visual C++ Years** (2005, 2008, 2010, etc. coexist by design)
- ❌ **Python Components** (Standard Library, pip, etc. are part of one installation)
- ❌ **Single Installations** (only removes duplicates)
- ❌ **System-Critical Updates**

## 💡 Example Output

```
========================================
  WinTrim - Smart Cleanup
========================================

Total programs found in registry: 99
Found 60 update-related programs.

Checking for latest versions online...
  ✓ Found online version for 'Microsoft Visual C++ 2022': 14.44.35211
  ✓ Skipping 'Python Standard Library' - part of active installation

Found 8 old versions to remove:

  - Microsoft Visual C++ 2022 X64 Debug Runtime - 14.36.32532
    Version: 14.36.32532
    Size: 30.31 MB

Estimated Disk Space to be Freed: 102.95 MB

Do you want to proceed? (Y/N)
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

- [ ] GUI version
- [ ] Scheduled cleanup tasks
- [ ] Export/import cleanup profiles
- [ ] Detailed HTML reports
- [ ] Rollback specific items
- [ ] Portable version (no installation required)

---

**Made with ❤️ for the Windows community**

*WinTrim - Keep your Windows installation lean and clean!*
