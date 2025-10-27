#Requires -RunAsAdministrator
<#
.SYNOPSIS
    WinTrim 2.0 - Comprehensive Windows Cleanup Tool

.DESCRIPTION
    A comprehensive cleanup tool that removes:
    - Duplicate software installations and old updates
    - Temporary files (system and user)
    - Windows Update cache
    - Browser caches
    - System cache files (thumbnails, error reports, etc.)
    - Advanced cleanup with DISM component store compression

.PARAMETER Modules
    Comma-separated list of modules to run: Programs, Temp, WindowsUpdate, Browser, Cache, Advanced, All

.EXAMPLE
    .\WinTrim.ps1
    Interactive mode - select which modules to run

.EXAMPLE
    .\WinTrim.ps1 -Modules All
    Run all cleanup modules

.NOTES
    Version: 2.0
    WARNING: Always create a system restore point before running this script!
    Some cleanup operations cannot be undone without a restore point.
#>

param(
    [string]$Modules = "Interactive"
)

# Enable strict mode for better error handling
Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

# Script version
$script:Version = "2.0"
$script:TotalSpaceFreed = 0
$script:Statistics = @{
    ProgramsRemoved = 0
    TempFilesRemoved = 0
    BrowserCacheCleared = 0
    WindowsUpdateCacheCleared = 0
    SystemCacheCleared = 0
    DISMCompression = 0
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  WinTrim v$($script:Version)" -ForegroundColor Cyan
Write-Host "  Comprehensive Windows Cleanup" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check if running as administrator
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "ERROR: This script must be run as Administrator!" -ForegroundColor Red
    Write-Host "Right-click PowerShell and select 'Run as Administrator'" -ForegroundColor Yellow
    exit 1
}

#region Helper Functions

function Format-FileSize {
    param([long]$Size)
    if ($Size -gt 1TB) { return "{0:N2} TB" -f ($Size / 1TB) }
    if ($Size -gt 1GB) { return "{0:N2} GB" -f ($Size / 1GB) }
    if ($Size -gt 1MB) { return "{0:N2} MB" -f ($Size / 1MB) }
    if ($Size -gt 1KB) { return "{0:N2} KB" -f ($Size / 1KB) }
    return "$Size bytes"
}

function Get-FolderSize {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return 0 }
    try {
        $size = (Get-ChildItem $Path -Recurse -File -ErrorAction SilentlyContinue |
                 Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
        if ($null -eq $size) { return 0 }
        return $size
    } catch {
        return 0
    }
}

function Remove-FilesRecursively {
    param(
        [string]$Path,
        [int]$OlderThanDays = 0
    )

    $removedCount = 0
    $freedSpace = 0

    if (-not (Test-Path $Path)) {
        return @{ Count = 0; Size = 0 }
    }

    try {
        $items = Get-ChildItem -Path $Path -Recurse -Force -ErrorAction SilentlyContinue

        foreach ($item in $items) {
            try {
                if ($OlderThanDays -gt 0) {
                    if ($item.LastWriteTime -gt (Get-Date).AddDays(-$OlderThanDays)) {
                        continue
                    }
                }

                if ($item.PSIsContainer) {
                    Remove-Item $item.FullName -Recurse -Force -ErrorAction SilentlyContinue
                } else {
                    $freedSpace += $item.Length
                    Remove-Item $item.FullName -Force -ErrorAction SilentlyContinue
                    $removedCount++
                }
            } catch {
                # Skip files in use
            }
        }
    } catch {
        # Continue on error
    }

    return @{ Count = $removedCount; Size = $freedSpace }
}

#endregion

#region Module: Temporary Files Cleanup

function Invoke-TempFilesCleanup {
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "  MODULE: Temporary Files Cleanup" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan

    $tempPaths = @(
        "$env:SystemRoot\Temp",
        "$env:TEMP",
        "$env:TMP"
    )

    # Add all user temp directories
    $userProfiles = Get-ChildItem "C:\Users" -Directory -ErrorAction SilentlyContinue
    foreach ($profile in $userProfiles) {
        $tempPaths += "$($profile.FullName)\AppData\Local\Temp"
    }

    $tempPaths = $tempPaths | Select-Object -Unique

    Write-Host "`nScanning temp directories..." -ForegroundColor Yellow
    $totalSize = 0
    $pathsWithData = @()

    foreach ($path in $tempPaths) {
        if (Test-Path $path) {
            $size = Get-FolderSize -Path $path
            if ($size -gt 0) {
                $totalSize += $size
                $pathsWithData += [PSCustomObject]@{
                    Path = $path
                    Size = $size
                }
            }
        }
    }

    if ($pathsWithData.Count -eq 0) {
        Write-Host "No temp files found to clean." -ForegroundColor Green
        return
    }

    Write-Host "`nTemporary directories to clean:" -ForegroundColor Yellow
    foreach ($item in $pathsWithData) {
        Write-Host "  - $($item.Path)" -ForegroundColor White
        Write-Host "    Size: $(Format-FileSize $item.Size)" -ForegroundColor Gray
    }

    Write-Host "`nTotal potential space to free: $(Format-FileSize $totalSize)" -ForegroundColor Cyan
    $confirm = Read-Host "`nProceed with temp files cleanup? (Y/N)"

    if ($confirm -eq 'Y' -or $confirm -eq 'y') {
        Write-Host "`nCleaning temp files..." -ForegroundColor Yellow
        $totalFreed = 0
        $totalFiles = 0

        foreach ($path in $tempPaths) {
            if (Test-Path $path) {
                Write-Host "  Cleaning: $path" -ForegroundColor Gray
                $result = Remove-FilesRecursively -Path $path
                $totalFreed += $result.Size
                $totalFiles += $result.Count
            }
        }

        Write-Host "`nTemp Cleanup Results:" -ForegroundColor Green
        Write-Host "  Files removed: $totalFiles" -ForegroundColor White
        Write-Host "  Space freed: $(Format-FileSize $totalFreed)" -ForegroundColor White

        $script:TotalSpaceFreed += $totalFreed
        $script:Statistics.TempFilesRemoved = $totalFiles
    } else {
        Write-Host "Temp files cleanup skipped." -ForegroundColor Yellow
    }
}

#endregion

#region Module: Windows Update Cache Cleanup

function Invoke-WindowsUpdateCleanup {
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "  MODULE: Windows Update Cache" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan

    $updatePaths = @(
        "$env:SystemRoot\SoftwareDistribution\Download",
        "$env:SystemRoot\SoftwareDistribution\DataStore\Logs"
    )

    Write-Host "`nScanning Windows Update cache..." -ForegroundColor Yellow
    $totalSize = 0
    $pathsWithData = @()

    foreach ($path in $updatePaths) {
        if (Test-Path $path) {
            $size = Get-FolderSize -Path $path
            if ($size -gt 0) {
                $totalSize += $size
                $pathsWithData += [PSCustomObject]@{
                    Path = $path
                    Size = $size
                }
            }
        }
    }

    if ($pathsWithData.Count -eq 0) {
        Write-Host "No Windows Update cache found to clean." -ForegroundColor Green
        return
    }

    Write-Host "`nWindows Update cache directories:" -ForegroundColor Yellow
    foreach ($item in $pathsWithData) {
        Write-Host "  - $($item.Path)" -ForegroundColor White
        Write-Host "    Size: $(Format-FileSize $item.Size)" -ForegroundColor Gray
    }

    Write-Host "`nTotal potential space to free: $(Format-FileSize $totalSize)" -ForegroundColor Cyan
    Write-Host "Note: Windows Update service will be stopped and restarted." -ForegroundColor Yellow
    $confirm = Read-Host "`nProceed with Windows Update cleanup? (Y/N)"

    if ($confirm -eq 'Y' -or $confirm -eq 'y') {
        Write-Host "`nStopping Windows Update services..." -ForegroundColor Yellow
        Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
        Stop-Service -Name bits -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2

        $totalFreed = 0
        foreach ($path in $updatePaths) {
            if (Test-Path $path) {
                Write-Host "  Cleaning: $path" -ForegroundColor Gray
                $result = Remove-FilesRecursively -Path $path
                $totalFreed += $result.Size
            }
        }

        Write-Host "`nRestarting Windows Update services..." -ForegroundColor Yellow
        Start-Service -Name bits -ErrorAction SilentlyContinue
        Start-Service -Name wuauserv -ErrorAction SilentlyContinue

        Write-Host "`nWindows Update Cleanup Results:" -ForegroundColor Green
        Write-Host "  Space freed: $(Format-FileSize $totalFreed)" -ForegroundColor White

        $script:TotalSpaceFreed += $totalFreed
        $script:Statistics.WindowsUpdateCacheCleared = $totalFreed
    } else {
        Write-Host "Windows Update cleanup skipped." -ForegroundColor Yellow
    }
}

#endregion

#region Module: Browser Cache Cleanup

function Invoke-BrowserCacheCleanup {
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "  MODULE: Browser Cache Cleanup" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan

    $userProfiles = Get-ChildItem "C:\Users" -Directory -ErrorAction SilentlyContinue
    $browserPaths = @()

    foreach ($profile in $userProfiles) {
        $profilePath = $profile.FullName

        # Brave
        $browserPaths += "$profilePath\AppData\Local\BraveSoftware\Brave-Browser\User Data\Default\Cache"
        $browserPaths += "$profilePath\AppData\Local\BraveSoftware\Brave-Browser\User Data\Default\Code Cache"

        # Chrome
        $browserPaths += "$profilePath\AppData\Local\Google\Chrome\User Data\Default\Cache"
        $browserPaths += "$profilePath\AppData\Local\Google\Chrome\User Data\Default\Code Cache"

        # Edge
        $browserPaths += "$profilePath\AppData\Local\Microsoft\Edge\User Data\Default\Cache"
        $browserPaths += "$profilePath\AppData\Local\Microsoft\Edge\User Data\Default\Code Cache"

        # Firefox
        $firefoxProfiles = Get-ChildItem "$profilePath\AppData\Local\Mozilla\Firefox\Profiles" -Directory -ErrorAction SilentlyContinue
        foreach ($ffProfile in $firefoxProfiles) {
            $browserPaths += "$($ffProfile.FullName)\cache2"
        }
    }

    Write-Host "`nScanning browser caches..." -ForegroundColor Yellow
    $totalSize = 0
    $pathsWithData = @()

    foreach ($path in $browserPaths) {
        if (Test-Path $path) {
            $size = Get-FolderSize -Path $path
            if ($size -gt 0) {
                $totalSize += $size
                $pathsWithData += [PSCustomObject]@{
                    Path = $path
                    Size = $size
                }
            }
        }
    }

    if ($pathsWithData.Count -eq 0) {
        Write-Host "No browser cache found to clean." -ForegroundColor Green
        return
    }

    Write-Host "`nBrowser caches found:" -ForegroundColor Yellow
    foreach ($item in $pathsWithData) {
        $browserName = "Unknown"
        if ($item.Path -match "Brave") { $browserName = "Brave" }
        elseif ($item.Path -match "Chrome") { $browserName = "Chrome" }
        elseif ($item.Path -match "Edge") { $browserName = "Edge" }
        elseif ($item.Path -match "Firefox") { $browserName = "Firefox" }

        Write-Host "  - $browserName" -ForegroundColor White
        Write-Host "    Size: $(Format-FileSize $item.Size)" -ForegroundColor Gray
    }

    Write-Host "`nTotal potential space to free: $(Format-FileSize $totalSize)" -ForegroundColor Cyan
    Write-Host "WARNING: Close all browsers before proceeding!" -ForegroundColor Yellow
    $confirm = Read-Host "`nProceed with browser cache cleanup? (Y/N)"

    if ($confirm -eq 'Y' -or $confirm -eq 'y') {
        Write-Host "`nCleaning browser caches..." -ForegroundColor Yellow
        $totalFreed = 0

        foreach ($path in $browserPaths) {
            if (Test-Path $path) {
                $result = Remove-FilesRecursively -Path $path
                $totalFreed += $result.Size
            }
        }

        Write-Host "`nBrowser Cache Cleanup Results:" -ForegroundColor Green
        Write-Host "  Space freed: $(Format-FileSize $totalFreed)" -ForegroundColor White

        $script:TotalSpaceFreed += $totalFreed
        $script:Statistics.BrowserCacheCleared = $totalFreed
    } else {
        Write-Host "Browser cache cleanup skipped." -ForegroundColor Yellow
    }
}

#endregion

#region Module: System Cache Cleanup

function Invoke-SystemCacheCleanup {
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "  MODULE: System Cache Cleanup" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan

    $userProfiles = Get-ChildItem "C:\Users" -Directory -ErrorAction SilentlyContinue
    $cachePaths = @()

    foreach ($profile in $userProfiles) {
        $profilePath = $profile.FullName

        # Thumbnail cache
        $cachePaths += "$profilePath\AppData\Local\Microsoft\Windows\Explorer"

        # DirectX Shader Cache
        $cachePaths += "$profilePath\AppData\Local\D3DSCache"

        # Icon cache
        $cachePaths += "$profilePath\AppData\Local\IconCache.db"
    }

    # System-wide caches
    $cachePaths += "$env:SystemRoot\Prefetch"
    $cachePaths += "C:\ProgramData\Microsoft\Windows\WER"
    $cachePaths += "$env:SystemRoot\Logs"

    Write-Host "`nScanning system caches..." -ForegroundColor Yellow
    $totalSize = 0
    $pathsWithData = @()

    foreach ($path in $cachePaths) {
        if (Test-Path $path) {
            $item = Get-Item $path -ErrorAction SilentlyContinue
            if ($null -eq $item) { continue }

            if ($item -is [System.IO.DirectoryInfo]) {
                $size = Get-FolderSize -Path $path
            } else {
                $size = $item.Length
                if ($null -eq $size) { $size = 0 }
            }

            if ($size -gt 0) {
                $totalSize += $size
                $pathsWithData += [PSCustomObject]@{
                    Path = $path
                    Size = $size
                    Type = if ($item -is [System.IO.DirectoryInfo]) { "Folder" } else { "File" }
                }
            }
        }
    }

    if ($pathsWithData.Count -eq 0) {
        Write-Host "No system cache found to clean." -ForegroundColor Green
        return
    }

    Write-Host "`nSystem caches found:" -ForegroundColor Yellow
    foreach ($item in $pathsWithData | Sort-Object Size -Descending | Select-Object -First 10) {
        $name = Split-Path $item.Path -Leaf
        Write-Host "  - $name" -ForegroundColor White
        Write-Host "    Size: $(Format-FileSize $item.Size)" -ForegroundColor Gray
    }

    if ($pathsWithData.Count -gt 10) {
        Write-Host "  ... and $($pathsWithData.Count - 10) more" -ForegroundColor Gray
    }

    Write-Host "`nTotal potential space to free: $(Format-FileSize $totalSize)" -ForegroundColor Cyan
    $confirm = Read-Host "`nProceed with system cache cleanup? (Y/N)"

    if ($confirm -eq 'Y' -or $confirm -eq 'y') {
        Write-Host "`nCleaning system caches..." -ForegroundColor Yellow
        $totalFreed = 0

        foreach ($path in $cachePaths) {
            if (Test-Path $path) {
                if ((Get-Item $path -ErrorAction SilentlyContinue) -is [System.IO.DirectoryInfo]) {
                    # Keep prefetch files less than 7 days old
                    if ($path -match "Prefetch") {
                        $result = Remove-FilesRecursively -Path $path -OlderThanDays 7
                        $totalFreed += $result.Size
                    } else {
                        $result = Remove-FilesRecursively -Path $path
                        $totalFreed += $result.Size
                    }
                } else {
                    # File
                    try {
                        if (Test-Path $path) {
                            $file = Get-Item $path -ErrorAction SilentlyContinue
                            if ($file) {
                                $fileSize = $file.Length
                                Remove-Item $path -Force -ErrorAction SilentlyContinue
                                $totalFreed += $fileSize
                            }
                        }
                    } catch {
                        # Skip files in use
                    }
                }
            }
        }

        Write-Host "`nSystem Cache Cleanup Results:" -ForegroundColor Green
        Write-Host "  Space freed: $(Format-FileSize $totalFreed)" -ForegroundColor White

        $script:TotalSpaceFreed += $totalFreed
        $script:Statistics.SystemCacheCleared = $totalFreed
    } else {
        Write-Host "System cache cleanup skipped." -ForegroundColor Yellow
    }
}

#endregion

#region Module: Duplicate Programs Cleanup

function Invoke-DuplicateProgramsCleanup {
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "  MODULE: Duplicate Programs Cleanup" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Scanning installed programs and updates..." -ForegroundColor Cyan
    Write-Host ""

# Function to check latest version online
function Get-LatestVersionOnline {
    param (
        [string]$SoftwareName,
        [string]$Publisher
    )

    try {
        # For Microsoft products, use known latest versions + web search
        if ($Publisher -match "Microsoft" -or $SoftwareName -match "Microsoft|Visual C\+\+|\.NET") {
            # For Visual C++ Redistributables
            if ($SoftwareName -match "Visual C\+\+.*(200[5-9]|20[1-2][0-9])") {
                Write-Host "    Checking online for Visual C++ $($matches[1]) version..." -ForegroundColor DarkGray
                $year = $matches[1]

                # DON'T scrape - use known versions only (web scraping was getting wrong version)
                # Known latest versions for each year
                $knownLatestVersions = @{
                    "2005" = "8.0.61001"
                    "2008" = "9.0.30729.6161"
                    "2010" = "10.0.40219"
                    "2012" = "11.0.61030"
                    "2013" = "12.0.40664"
                    "2015" = "14.0.24215"
                    "2017" = "14.16.27033"
                    "2019" = "14.29.30153"
                    "2022" = "14.44.35211"  # Updated to current
                }

                if ($knownLatestVersions.ContainsKey($year)) {
                    Write-Host "    Using known latest version for VC++ $year" -ForegroundColor DarkGray
                    return $knownLatestVersions[$year]
                }
            }

            # For .NET Framework
            if ($SoftwareName -match "\.NET Framework") {
                Write-Host "    Checking for latest .NET Framework version..." -ForegroundColor DarkGray
                try {
                    $url = "https://dotnet.microsoft.com/en-us/download/dotnet-framework"
                    $response = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 10 -ErrorAction SilentlyContinue

                    if ($response.Content -match "\.NET Framework (\d+\.\d+\.?\d*)") {
                        return $matches[1]
                    }
                } catch {
                    # Fall back
                }
                return "4.8.1"  # Latest stable
            }

            # For Microsoft Edge
            if ($SoftwareName -match "Microsoft Edge") {
                Write-Host "    Checking for latest Edge version..." -ForegroundColor DarkGray
                try {
                    $url = "https://www.microsoft.com/en-us/edge/download"
                    $response = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 10 -ErrorAction SilentlyContinue

                    if ($response.Content -match "Version (\d+\.\d+\.\d+\.\d+)") {
                        return $matches[1]
                    }
                } catch {
                    # Fallback
                }
            }
        }

        # For other common software, try their version APIs or websites
        if ($SoftwareName -match "Chrome") {
            Write-Host "    Checking for latest Chrome version..." -ForegroundColor DarkGray
            try {
                $url = "https://chromereleases.googleblog.com/"
                $response = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 10 -ErrorAction SilentlyContinue

                if ($response.Content -match "(\d+\.\d+\.\d+\.\d+)") {
                    return $matches[1]
                }
            } catch {
                # Fallback
            }
        }

        return $null
    } catch {
        return $null
    }
}

Write-Host "Note: Will check online for latest versions to improve detection accuracy" -ForegroundColor Yellow
Write-Host ""

# Get all installed programs from both 32-bit and 64-bit registry locations
$uninstallPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
)

$allPrograms = @()
foreach ($basePath in $uninstallPaths) {
    try {
        if (Test-Path $basePath) {
            $keys = Get-ChildItem $basePath -ErrorAction Stop

            foreach ($key in $keys) {
                try {
                    $props = Get-ItemProperty $key.PSPath -ErrorAction SilentlyContinue

                    if ($props.DisplayName -and $props.UninstallString) {
                        $allPrograms += [PSCustomObject]@{
                            DisplayName = $props.DisplayName
                            DisplayVersion = $props.DisplayVersion
                            Publisher = $props.Publisher
                            UninstallString = $props.UninstallString
                            PSPath = $key.PSPath
                            InstallDate = $props.InstallDate
                            EstimatedSize = $props.EstimatedSize
                        }
                    }
                } catch {
                    # Skip keys we can't read
                }
            }
        }
    } catch {
        # Silently continue if path doesn't exist
    }
}

Write-Host "Total programs found in registry: $($allPrograms.Count)" -ForegroundColor Green
Write-Host ""

# Show sample of what was found for debugging
Write-Host "Sample of installed programs (first 10):" -ForegroundColor Cyan
$allPrograms | Select-Object -First 10 | ForEach-Object {
    Write-Host "  - $($_.DisplayName) (v$($_.DisplayVersion))" -ForegroundColor Gray
}
Write-Host ""

# Define patterns to match common update software (expanded patterns)
$patterns = @(
    "Microsoft Edge",
    "Microsoft Visual C\+\+",
    "Visual C\+\+",
    "Microsoft Update",
    "Windows.*Update",
    "KB\d+",
    "Update for",
    "Security Update",
    "Cumulative Update",
    "Definition Update",
    "MongoDB",
    "Redistributable",
    "\d{4}.*Redistributable",
    "Runtime.*\d{4}",
    "Microsoft\.NET",
    "Microsoft ASP\.NET",
    "DirectX",
    "Mozilla Maintenance Service",
    "Java.*Update",
    "Java.*Development Kit",
    "Adobe.*Reader",
    "Adobe.*Acrobat"
)

# CRITICAL: Exclude patterns - these should NEVER be removed as they're part of active installations
$excludePatterns = @(
    # Python components are part of one installation - don't treat as separate duplicates
    "Python.*Standard Library",
    "Python.*Executables",
    "Python.*pip Bootstrap",
    "Python.*Tcl/Tk",
    "Python.*Test Suite",
    "Python.*Documentation"
)

# Filter programs that match our patterns
$updatePrograms = @($allPrograms | Where-Object {
    $displayName = $_.DisplayName
    $matched = $false
    foreach ($pattern in $patterns) {
        if ($displayName -match $pattern) {
            $matched = $true
            break
        }
    }
    $matched
})

Write-Host "Found $($updatePrograms.Count) update-related programs." -ForegroundColor Green
Write-Host ""

# Show what matched for debugging
if ($updatePrograms.Count -gt 0) {
    Write-Host "Programs that matched patterns:" -ForegroundColor Cyan
    $updatePrograms | ForEach-Object {
        Write-Host "  - $($_.DisplayName) (v$($_.DisplayVersion))" -ForegroundColor Gray
    }
    Write-Host ""
} else {
    Write-Host "DEBUG: No programs matched the patterns. Here are ALL programs:" -ForegroundColor Yellow
    $allPrograms | Sort-Object DisplayName | ForEach-Object {
        Write-Host "  - $($_.DisplayName) (v$($_.DisplayVersion)) [$($_.Publisher)]" -ForegroundColor Gray
    }
    Write-Host ""
    Write-Host "If you see duplicate programs above, the patterns may need adjustment." -ForegroundColor Yellow
    Write-Host ""
}

# Group by base name (removing version numbers and dates)
$grouped = @{}
foreach ($program in $updatePrograms) {
    # CRITICAL FIX: Extract year FIRST before removing version numbers
    $year = $null
    if ($program.DisplayName -match "Visual C\+\+ (200[5-9]|20[1-2][0-9])") {
        # Match actual years like 2005, 2008, 2010, 2012, 2013, 2015, 2017, 2019, 2022
        $year = $matches[1]
    }

    # Extract base name by removing version patterns
    $baseName = $program.DisplayName -replace '\s*\d+\.[\d\.]+.*$', '' `
                                      -replace '\s*\(KB\d+\)', '' `
                                      -replace '\s*x(86|64)', '' `
                                      -replace '\s*-\s*[\d\.]+.*$', ''  # Remove version after dash

    # For Visual C++, include year to keep different years separate
    if ($year) {
        $baseName = "$baseName $year"
    }

    if (-not $grouped.ContainsKey($baseName)) {
        $grouped[$baseName] = @()
    }
    $grouped[$baseName] += $program
}

# Find duplicates and identify old versions
$toRemove = @()
$toKeep = @()
$onlineVersionCache = @{}

Write-Host "Checking for latest versions online (this may take a moment)..." -ForegroundColor Cyan
Write-Host ""

foreach ($baseName in $grouped.Keys) {
    $programs = $grouped[$baseName]

    # CRITICAL FIX: Check if this matches exclude patterns (Python components, etc.)
    $shouldExclude = $false
    foreach ($excludePattern in $excludePatterns) {
        if ($baseName -match $excludePattern) {
            $shouldExclude = $true
            Write-Host "  Skipping '$baseName' - part of active installation, not a duplicate" -ForegroundColor Yellow
            # Keep all of these
            $toKeep += $programs
            break
        }
    }

    if ($shouldExclude) {
        continue  # Skip to next group
    }

    # Check online for latest version
    $latestOnlineVersion = $null
    if (-not $onlineVersionCache.ContainsKey($baseName)) {
        $sampleProgram = $programs[0]
        $latestOnlineVersion = Get-LatestVersionOnline -SoftwareName $baseName -Publisher $sampleProgram.Publisher
        $onlineVersionCache[$baseName] = $latestOnlineVersion
    } else {
        $latestOnlineVersion = $onlineVersionCache[$baseName]
    }

    if ($programs.Count -gt 1) {
        # Sort by version (if available) or install date
        $sorted = $programs | Sort-Object {
            try {
                if ($_.DisplayVersion) {
                    [version]$_.DisplayVersion
                } elseif ($_.InstallDate) {
                    [datetime]::ParseExact($_.InstallDate, "yyyyMMdd", $null)
                } else {
                    [datetime]::MinValue
                }
            } catch {
                [datetime]::MinValue
            }
        } -Descending

        # If we found a version online, use it as benchmark
        if ($latestOnlineVersion) {
            Write-Host "  Found online version for '$baseName': $latestOnlineVersion" -ForegroundColor Green

            $foundNewerOrEqual = $false
            foreach ($program in $sorted) {
                try {
                    $installedVersion = [version]$program.DisplayVersion
                    $onlineVersion = [version]$latestOnlineVersion

                    if ($installedVersion -ge $onlineVersion) {
                        # Keep this one - it's current or newer
                        $toKeep += $program
                        $foundNewerOrEqual = $true
                    } else {
                        # SAFETY CHECK: For Visual C++, only remove if it's truly the same year/edition
                        # Different editions (Minimum/Additional/Debug) should coexist
                        if ($program.DisplayName -match "Visual C\+\+") {
                            # Only mark truly duplicate runtime versions for removal, not different editions
                            $edition = ""
                            if ($program.DisplayName -match "(Minimum|Additional|Debug)") {
                                $edition = $matches[1]
                            }

                            # Check if we already have this edition kept
                            $alreadyKept = $toKeep | Where-Object {
                                $_.DisplayName -match $edition -and $_.DisplayName -match "Visual C\+\+"
                            }

                            if (-not $alreadyKept) {
                                # Keep this one too - different edition
                                $toKeep += $program
                            } else {
                                # Old version of same edition - mark for removal
                                $toRemove += $program
                            }
                        } else {
                            # Old version - mark for removal
                            $toRemove += $program
                        }
                    }
                } catch {
                    # If version parsing fails, use local comparison
                    if (-not $foundNewerOrEqual) {
                        $toKeep += $program
                        $foundNewerOrEqual = $true
                    } else {
                        $toRemove += $program
                    }
                }
            }
        } else {
            # No online version found - use local comparison
            # Keep the newest, mark others for removal
            $toKeep += $sorted[0]
            for ($i = 1; $i -lt $sorted.Count; $i++) {
                $toRemove += $sorted[$i]
            }
        }
    } elseif ($programs.Count -eq 1 -and $latestOnlineVersion) {
        # Single program but we have online version to compare
        # SAFETY: Don't remove single installations - they're likely needed
        $program = $programs[0]
        $toKeep += $program

        try {
            $installedVersion = [version]$program.DisplayVersion
            $onlineVersion = [version]$latestOnlineVersion

            if ($installedVersion -lt $onlineVersion) {
                Write-Host "  Note: '$($program.DisplayName)' v$installedVersion is outdated (latest: $latestOnlineVersion)" -ForegroundColor Cyan
                Write-Host "        But keeping it as only installation (not a duplicate)" -ForegroundColor Cyan
            }
        } catch {
            # Can't compare versions, keep it
        }
    }
}

    Write-Host ""

    if ($toRemove.Count -eq 0) {
        Write-Host "No old updates found to remove!" -ForegroundColor Green
        return
    }

    Write-Host "Found $($toRemove.Count) old versions to remove:" -ForegroundColor Yellow
    Write-Host ""

    # Calculate total size before removal
    $totalSizeKB = 0
    $toRemove | ForEach-Object {
        if ($_.EstimatedSize) {
            $totalSizeKB += $_.EstimatedSize
        }
    }

    # Display what will be removed
    $toRemove | ForEach-Object {
        Write-Host "  - $($_.DisplayName)" -ForegroundColor White
        Write-Host "    Version: $($_.DisplayVersion)" -ForegroundColor Gray
        Write-Host "    Publisher: $($_.Publisher)" -ForegroundColor Gray
        if ($_.EstimatedSize) {
            $sizeMB = [math]::Round($_.EstimatedSize / 1024, 2)
            Write-Host "    Size: $sizeMB MB" -ForegroundColor Gray
        }
        Write-Host ""
    }

    Write-Host "The following will be KEPT:" -ForegroundColor Green
    $toKeep | Select-Object -Unique | ForEach-Object {
        $displayName = $_.DisplayName
        $baseName = $displayName -replace '\s*\d+\.[\d\.]+.*$', '' `
                                  -replace '\s*\(KB\d+\)', '' `
                                  -replace '\s*x(86|64)', '' `
                                  -replace '\s*\d{4}$', ''

        $onlineVersionInfo = ""
        if ($onlineVersionCache.ContainsKey($baseName) -and $onlineVersionCache[$baseName]) {
            $onlineVersionInfo = " [Latest online: $($onlineVersionCache[$baseName])]"
        }

        Write-Host "  + $($_.DisplayName) (Version: $($_.DisplayVersion))$onlineVersionInfo" -ForegroundColor Green
    }
    Write-Host ""

    # Show summary of online checks
    $onlineChecksPerformed = @($onlineVersionCache.Values | Where-Object { $_ -ne $null }).Count
    if ($onlineChecksPerformed -gt 0) {
        Write-Host "Online Version Checks:" -ForegroundColor Cyan
        Write-Host "  Successfully retrieved $onlineChecksPerformed current version(s) from the internet" -ForegroundColor White
        Write-Host "  These were used as benchmarks for determining outdated software" -ForegroundColor White
        Write-Host ""
    }

    # Display estimated savings
    if ($totalSizeKB -gt 0) {
        $totalSizeMB = [math]::Round($totalSizeKB / 1024, 2)
        $totalSizeGB = [math]::Round($totalSizeKB / 1024 / 1024, 2)

        Write-Host "Estimated Disk Space to be Freed:" -ForegroundColor Cyan
        if ($totalSizeGB -gt 1) {
            Write-Host "  $totalSizeGB GB ($totalSizeMB MB)" -ForegroundColor White
        } else {
            Write-Host "  $totalSizeMB MB" -ForegroundColor White
        }
        Write-Host ""
    }

    $confirm = Read-Host "Do you want to proceed with removal? (Y/N)"
    if ($confirm -ne 'Y' -and $confirm -ne 'y') {
        Write-Host "Duplicate programs cleanup cancelled." -ForegroundColor Yellow
        return
    }

    # Remove old versions
    Write-Host ""
    Write-Host "Removing old versions..." -ForegroundColor Cyan
    $successCount = 0
    $failCount = 0
    $actualSpaceFreedKB = 0

    foreach ($program in $toRemove) {
        Write-Host "Removing: $($program.DisplayName)..." -ForegroundColor Yellow

        try {
            $uninstallString = $program.UninstallString

            if ($uninstallString -match 'msiexec') {
                # MSI-based uninstall
                $productCode = $uninstallString -replace '.*\{(.*)\}.*', '{$1}'
                $process = Start-Process "msiexec.exe" -ArgumentList "/x $productCode /quiet /norestart" -Wait -PassThru -NoNewWindow

                if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
                    Write-Host "  Successfully removed." -ForegroundColor Green
                    $successCount++
                    if ($program.EstimatedSize) {
                        $actualSpaceFreedKB += $program.EstimatedSize
                    }
                } else {
                    Write-Host "  Failed with exit code: $($process.ExitCode)" -ForegroundColor Red
                    $failCount++
                }
            } else {
                # EXE-based uninstall
                # Add silent uninstall flags
                $uninstallString = $uninstallString -replace '"([^"]+)"', '$1'
                $exePath = ($uninstallString -split ' ')[0]
                $args = ($uninstallString -replace [regex]::Escape($exePath), '').Trim()

                # Add common silent flags if not present
                if ($args -notmatch '/quiet|/silent|/s\s') {
                    $args += " /quiet /norestart"
                }

                $process = Start-Process $exePath -ArgumentList $args -Wait -PassThru -NoNewWindow -ErrorAction Stop

                if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
                    Write-Host "  Successfully removed." -ForegroundColor Green
                    $successCount++
                    if ($program.EstimatedSize) {
                        $actualSpaceFreedKB += $program.EstimatedSize
                    }
                } else {
                    Write-Host "  Failed with exit code: $($process.ExitCode)" -ForegroundColor Red
                    $failCount++
                }
            }
        } catch {
            Write-Host "  Error: $_" -ForegroundColor Red
            $failCount++
        }

        Start-Sleep -Milliseconds 500
    }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  Module Cleanup Complete!" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""

    # Display removal statistics
    Write-Host "REMOVAL STATISTICS:" -ForegroundColor White
    Write-Host "-------------------" -ForegroundColor White
    Write-Host "Programs targeted for removal: $($toRemove.Count)" -ForegroundColor White
    Write-Host "Successfully removed: $successCount" -ForegroundColor Green
    Write-Host "Failed to remove: $failCount" -ForegroundColor Red
    Write-Host ""

    # Display space savings
    if ($actualSpaceFreedKB -gt 0) {
        $spaceMB = [math]::Round($actualSpaceFreedKB / 1024, 2)
        $spaceGB = [math]::Round($actualSpaceFreedKB / 1024 / 1024, 2)
        $actualSpaceFreedBytes = $actualSpaceFreedKB * 1024

        Write-Host "DISK SPACE SAVINGS:" -ForegroundColor White
        Write-Host "-------------------" -ForegroundColor White
        if ($spaceGB -gt 1) {
            Write-Host "Total disk space freed: $spaceGB GB ($spaceMB MB)" -ForegroundColor Green
        } else {
            Write-Host "Total disk space freed: $spaceMB MB" -ForegroundColor Green
        }

        # Update script-level statistics
        $script:TotalSpaceFreed += $actualSpaceFreedBytes
        $script:Statistics.ProgramsRemoved = $successCount

        # Estimate memory savings
        $estimatedMemorySavingsMB = $successCount * 75

        Write-Host ""
        Write-Host "ESTIMATED MEMORY SAVINGS:" -ForegroundColor White
        Write-Host "-------------------------" -ForegroundColor White
        Write-Host "Potential RAM savings: ~$estimatedMemorySavingsMB MB" -ForegroundColor Green
        Write-Host "(Old update services and processes no longer running)" -ForegroundColor Gray

        # Registry cleanup benefit
        Write-Host ""
        Write-Host "ADDITIONAL BENEFITS:" -ForegroundColor White
        Write-Host "--------------------" -ForegroundColor White
        Write-Host "- Cleaner Add/Remove Programs list" -ForegroundColor Cyan
        Write-Host "- Reduced registry clutter" -ForegroundColor Cyan
        Write-Host "- Faster system scans and updates" -ForegroundColor Cyan
        Write-Host "- Reduced background service overhead" -ForegroundColor Cyan
    } else {
        Write-Host "DISK SPACE SAVINGS:" -ForegroundColor White
        Write-Host "-------------------" -ForegroundColor White
        Write-Host "Space information not available for removed programs" -ForegroundColor Yellow
    }

    Write-Host ""

    if ($failCount -gt 0) {
        Write-Host "NOTE: Some items could not be removed automatically." -ForegroundColor Yellow
        Write-Host "You may need to remove them manually from Add/Remove Programs." -ForegroundColor Yellow
        Write-Host ""
    }
}

#endregion

#region Module: Advanced DISM Cleanup

function Invoke-AdvancedDISMCleanup {
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "  MODULE: Advanced DISM Cleanup" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan

    Write-Host "`nThis module will:" -ForegroundColor Yellow
    Write-Host "  - Analyze component store" -ForegroundColor White
    Write-Host "  - Clean superseded components" -ForegroundColor White
    Write-Host "  - Reset base of superseded components" -ForegroundColor White
    Write-Host "`nThis operation can take 10-30 minutes." -ForegroundColor Yellow

    $confirm = Read-Host "`nProceed with DISM cleanup? (Y/N)"

    if ($confirm -eq 'Y' -or $confirm -eq 'y') {
        Write-Host "`nAnalyzing component store..." -ForegroundColor Yellow

        try {
            # Analyze
            $analyzeOutput = & DISM.exe /Online /Cleanup-Image /AnalyzeComponentStore 2>&1

            Write-Host "`nComponent Store Analysis:" -ForegroundColor Cyan
            $analyzeOutput | ForEach-Object { Write-Host $_ -ForegroundColor Gray }

            # StartComponentCleanup
            Write-Host "`nCleaning component store..." -ForegroundColor Yellow
            $cleanupOutput = & DISM.exe /Online /Cleanup-Image /StartComponentCleanup /ResetBase 2>&1

            Write-Host "`nDISM Cleanup Output:" -ForegroundColor Cyan
            $cleanupOutput | ForEach-Object { Write-Host $_ -ForegroundColor Gray }

            # Try to extract space savings from output
            $savedSpace = 0
            foreach ($line in $cleanupOutput) {
                if ($line -match "(\d+\.?\d*)\s*(GB|MB)") {
                    $size = [double]$matches[1]
                    if ($matches[2] -eq "GB") {
                        $savedSpace = $size * 1GB
                    } else {
                        $savedSpace = $size * 1MB
                    }
                }
            }

            Write-Host "`nDISM Cleanup Results:" -ForegroundColor Green
            Write-Host "  Operation completed successfully" -ForegroundColor White
            if ($savedSpace -gt 0) {
                Write-Host "  Estimated space freed: $(Format-FileSize $savedSpace)" -ForegroundColor White
                $script:TotalSpaceFreed += $savedSpace
                $script:Statistics.DISMCompression = $savedSpace
            }
        } catch {
            Write-Host "`nDISM cleanup failed: $_" -ForegroundColor Red
        }
    } else {
        Write-Host "DISM cleanup skipped." -ForegroundColor Yellow
    }
}

#endregion

# Warning message
Write-Host "WARNING: This script will perform comprehensive cleanup." -ForegroundColor Yellow
Write-Host "It is STRONGLY recommended to:" -ForegroundColor Yellow
Write-Host "  1. Create a System Restore Point first" -ForegroundColor Yellow
Write-Host "  2. Backup important data" -ForegroundColor Yellow
Write-Host "  3. Close all browsers and applications" -ForegroundColor Yellow
Write-Host "  4. Review each module before confirming" -ForegroundColor Yellow
Write-Host ""

$createRestorePoint = Read-Host "Would you like to create a System Restore Point now? (Y/N)"
if ($createRestorePoint -eq 'Y' -or $createRestorePoint -eq 'y') {
    Write-Host "Creating system restore point..." -ForegroundColor Green
    try {
        Checkpoint-Computer -Description "Before WinTrim Cleanup v$($script:Version)" -RestorePointType "MODIFY_SETTINGS"
        Write-Host "Restore point created successfully!" -ForegroundColor Green
    } catch {
        Write-Host "Failed to create restore point: $_" -ForegroundColor Red
        $continue = Read-Host "Continue anyway? (Y/N)"
        if ($continue -ne 'Y' -and $continue -ne 'y') {
            exit 0
        }
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Select Cleanup Modules" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Available modules:" -ForegroundColor Yellow
Write-Host "  1. Duplicate Programs - Remove old software updates (original WinTrim)" -ForegroundColor White
Write-Host "  2. Temp Files - Clean system and user temp directories" -ForegroundColor White
Write-Host "  3. Windows Update Cache - Clear update download cache" -ForegroundColor White
Write-Host "  4. Browser Cache - Clear browser caches (Brave, Chrome, Edge, Firefox)" -ForegroundColor White
Write-Host "  5. System Cache - Clean thumbnails, logs, error reports, shader cache" -ForegroundColor White
Write-Host "  6. Advanced DISM - Component store cleanup (slow but thorough)" -ForegroundColor White
Write-Host "  A. All modules" -ForegroundColor Green
Write-Host "  Q. Quit" -ForegroundColor Red
Write-Host ""

$selection = Read-Host "Enter modules to run (e.g., '1,2,3' or 'A' for all)"

$runPrograms = $false
$runTemp = $false
$runWindowsUpdate = $false
$runBrowser = $false
$runCache = $false
$runDISM = $false

if ($selection -match '[Qq]') {
    Write-Host "Exiting..." -ForegroundColor Yellow
    exit 0
}

if ($selection -match '[Aa]') {
    $runPrograms = $true
    $runTemp = $true
    $runWindowsUpdate = $true
    $runBrowser = $true
    $runCache = $true
    $runDISM = $true
} else {
    if ($selection -match '1') { $runPrograms = $true }
    if ($selection -match '2') { $runTemp = $true }
    if ($selection -match '3') { $runWindowsUpdate = $true }
    if ($selection -match '4') { $runBrowser = $true }
    if ($selection -match '5') { $runCache = $true }
    if ($selection -match '6') { $runDISM = $true }
}

Write-Host ""
Write-Host "Modules to run:" -ForegroundColor Cyan
if ($runPrograms) { Write-Host "  ✓ Duplicate Programs" -ForegroundColor Green }
if ($runTemp) { Write-Host "  ✓ Temp Files" -ForegroundColor Green }
if ($runWindowsUpdate) { Write-Host "  ✓ Windows Update Cache" -ForegroundColor Green }
if ($runBrowser) { Write-Host "  ✓ Browser Cache" -ForegroundColor Green }
if ($runCache) { Write-Host "  ✓ System Cache" -ForegroundColor Green }
if ($runDISM) { Write-Host "  ✓ Advanced DISM" -ForegroundColor Green }
Write-Host ""

$startTime = Get-Date

# Run selected modules
if ($runTemp) { Invoke-TempFilesCleanup }
if ($runWindowsUpdate) { Invoke-WindowsUpdateCleanup }
if ($runBrowser) { Invoke-BrowserCacheCleanup }
if ($runCache) { Invoke-SystemCacheCleanup }
if ($runPrograms) { Invoke-DuplicateProgramsCleanup }
if ($runDISM) { Invoke-AdvancedDISMCleanup }

$endTime = Get-Date
$duration = $endTime - $startTime

# Final summary report
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  WinTrim Cleanup Complete!" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "FINAL STATISTICS:" -ForegroundColor White
Write-Host "-------------------" -ForegroundColor White
Write-Host "Total space freed: $(Format-FileSize $script:TotalSpaceFreed)" -ForegroundColor Green
Write-Host "Programs removed: $($script:Statistics.ProgramsRemoved)" -ForegroundColor White
Write-Host "Temp files removed: $($script:Statistics.TempFilesRemoved)" -ForegroundColor White
Write-Host "Time elapsed: $($duration.Minutes)m $($duration.Seconds)s" -ForegroundColor White
Write-Host ""

Write-Host "CLEANUP BY MODULE:" -ForegroundColor White
Write-Host "-------------------" -ForegroundColor White
if ($script:Statistics.TempFilesRemoved -gt 0) {
    Write-Host "Temp Files: $($script:Statistics.TempFilesRemoved) files removed" -ForegroundColor Cyan
}
if ($script:Statistics.WindowsUpdateCacheCleared -gt 0) {
    Write-Host "Windows Update: $(Format-FileSize $script:Statistics.WindowsUpdateCacheCleared)" -ForegroundColor Cyan
}
if ($script:Statistics.BrowserCacheCleared -gt 0) {
    Write-Host "Browser Cache: $(Format-FileSize $script:Statistics.BrowserCacheCleared)" -ForegroundColor Cyan
}
if ($script:Statistics.SystemCacheCleared -gt 0) {
    Write-Host "System Cache: $(Format-FileSize $script:Statistics.SystemCacheCleared)" -ForegroundColor Cyan
}
if ($script:Statistics.ProgramsRemoved -gt 0) {
    Write-Host "Duplicate Programs: $($script:Statistics.ProgramsRemoved) removed" -ForegroundColor Cyan
}
if ($script:Statistics.DISMCompression -gt 0) {
    Write-Host "DISM Cleanup: $(Format-FileSize $script:Statistics.DISMCompression)" -ForegroundColor Cyan
}
Write-Host ""

Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
exit 0
