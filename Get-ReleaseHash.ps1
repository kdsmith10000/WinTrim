# Get-ReleaseHash.ps1
# Generates SHA256 hashes for WinTrim release files

<#
.SYNOPSIS
    Generates SHA256 hashes for WinTrim release files.

.DESCRIPTION
    This script calculates SHA256 hashes for the main WinTrim.ps1 file
    and displays them in a format suitable for release notes and verification.

.EXAMPLE
    .\Get-ReleaseHash.ps1
    Generates and displays SHA256 hash for WinTrim.ps1

.EXAMPLE
    .\Get-ReleaseHash.ps1 -OutputFile checksums.txt
    Generates hash and saves to a file
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$OutputFile = $null
)

# Set console output encoding to UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "WinTrim Release Hash Generator" -ForegroundColor Cyan
Write-Host "================================" -ForegroundColor Cyan
Write-Host ""

# Get the script directory
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$releaseFile = Join-Path $scriptPath "WinTrim.ps1"

# Check if the file exists
if (-not (Test-Path $releaseFile)) {
    Write-Host "ERROR: WinTrim.ps1 not found in $scriptPath" -ForegroundColor Red
    exit 1
}

# Calculate SHA256 hash
Write-Host "Calculating SHA256 hash..." -ForegroundColor Yellow
$hash = Get-FileHash -Path $releaseFile -Algorithm SHA256

# Get file size
$fileInfo = Get-Item $releaseFile
$fileSizeKB = [math]::Round($fileInfo.Length / 1KB, 2)
$fileSizeMB = [math]::Round($fileInfo.Length / 1MB, 2)

# Display results
Write-Host ""
Write-Host "File Information:" -ForegroundColor Green
Write-Host "  Name: $($fileInfo.Name)"
Write-Host "  Size: $fileSizeKB KB ($fileSizeMB MB)"
Write-Host "  Modified: $($fileInfo.LastWriteTime)"
Write-Host ""
Write-Host "SHA256 Hash:" -ForegroundColor Green
Write-Host "  $($hash.Hash)" -ForegroundColor White
Write-Host ""

# Prepare output content
$outputContent = @"
WinTrim Release Checksums
Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")

File: $($fileInfo.Name)
Size: $fileSizeKB KB ($fileSizeMB MB)
SHA256: $($hash.Hash)
"@

# Save to file if requested
if ($OutputFile) {
    $outputPath = Join-Path $scriptPath $OutputFile
    $outputContent | Out-File -FilePath $outputPath -Encoding UTF8
    Write-Host "Checksums saved to: $OutputFile" -ForegroundColor Green
    Write-Host ""
}

# Also display in markdown format for release notes
Write-Host "Markdown Format (for GitHub releases):" -ForegroundColor Cyan
Write-Host "--------------------------------------" -ForegroundColor Cyan
Write-Host ""
Write-Host "## Checksums"
Write-Host ""
Write-Host "| File | SHA256 |"
Write-Host "|------|--------|"
Write-Host "| ``$($fileInfo.Name)`` | ``$($hash.Hash)`` |"
Write-Host ""
Write-Host "**File Size:** $fileSizeKB KB ($fileSizeMB MB)"
Write-Host ""

# Copy hash to clipboard if possible
try {
    $hash.Hash | Set-Clipboard
    Write-Host "Hash copied to clipboard!" -ForegroundColor Green
} catch {
    Write-Host "Note: Could not copy to clipboard" -ForegroundColor Yellow
}
