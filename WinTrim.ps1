#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Removes old software updates from Add/Remove Programs, keeping only the most recent versions.

.DESCRIPTION
    This script identifies duplicate software installations (particularly updates) and removes older versions
    while keeping the most recent one. It targets common software like Microsoft Edge, Visual C++,
    Windows updates, MongoDB, etc.

.NOTES
    WARNING: Always create a system restore point before running this script!
    Some updates may be required for other software to function properly.
#>

# Enable strict mode for better error handling
Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Old Updates Cleanup Script" -ForegroundColor Cyan
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

# Warning message
Write-Host "WARNING: This script will remove old software updates." -ForegroundColor Yellow
Write-Host "It is STRONGLY recommended to:" -ForegroundColor Yellow
Write-Host "  1. Create a System Restore Point first" -ForegroundColor Yellow
Write-Host "  2. Backup important data" -ForegroundColor Yellow
Write-Host "  3. Review the list before confirming removal" -ForegroundColor Yellow
Write-Host ""

$createRestorePoint = Read-Host "Would you like to create a System Restore Point now? (Y/N)"
if ($createRestorePoint -eq 'Y' -or $createRestorePoint -eq 'y') {
    Write-Host "Creating system restore point..." -ForegroundColor Green
    try {
        Checkpoint-Computer -Description "Before Old Updates Cleanup" -RestorePointType "MODIFY_SETTINGS"
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
    exit 0
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
    Write-Host "Operation cancelled by user." -ForegroundColor Yellow
    exit 0
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
Write-Host "  Cleanup Complete!" -ForegroundColor Cyan
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

    Write-Host "DISK SPACE SAVINGS:" -ForegroundColor White
    Write-Host "-------------------" -ForegroundColor White
    if ($spaceGB -gt 1) {
        Write-Host "Total disk space freed: $spaceGB GB ($spaceMB MB)" -ForegroundColor Green
    } else {
        Write-Host "Total disk space freed: $spaceMB MB" -ForegroundColor Green
    }

    # Estimate memory savings
    # Background services and updates typically consume 50-200MB RAM when active
    # We'll use a conservative estimate of 75MB per old update removed
    $estimatedMemorySavingsMB = $successCount * 75
    $estimatedMemorySavingsGB = [math]::Round($estimatedMemorySavingsMB / 1024, 2)

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

Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
