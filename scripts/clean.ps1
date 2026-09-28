<#
.SYNOPSIS
    Cleans build artifacts and temporary files safely without disrupting unrelated user Excel sessions.
.DESCRIPTION
    Removes dist/, temp/, __pycache__ directories, and *.pyc/*.tmp files.
    Optionally terminates a specific tracked PID if provided via -TargetPid.
    NEVER kills arbitrary or global Excel instances.
.PARAMETER TargetPid
    Optional specific process ID to terminate if an automated build process was orphaned.
.PARAMETER Force
    Optional switch to force cleanup of read-only files.
#>
[CmdletBinding()]
param (
    [int]$TargetPid = 0,
    [switch]$Force
)

$ErrorActionPreference = "Continue"

Write-Host "=== BWPConvertTTNVN Build Cleanup ===" -ForegroundColor Cyan

# 1. Terminate ONLY specific tracked PID if explicitly requested
if ($TargetPid -gt 0) {
    try {
        $proc = Get-Process -Id $TargetPid -ErrorAction SilentlyContinue
        if ($proc -and $proc.ProcessName -eq "EXCEL") {
            Write-Host "Terminating build-owned Excel process (PID: $TargetPid)..." -ForegroundColor Yellow
            Stop-Process -Id $TargetPid -Force -ErrorAction SilentlyContinue
        } else {
            Write-Host "Process PID $TargetPid is not running or not EXCEL. Skipping process termination." -ForegroundColor Gray
        }
    } catch {
        Write-Warning "Could not terminate process $($TargetPid): $_"
    }
} else {
    Write-Host "Safe mode: No specific TargetPid specified. Existing Excel sessions preserved." -ForegroundColor Green
}

# 2. Determine project root directory
$RootDir = Split-Path -Parent $PSScriptRoot
if (-not $RootDir -or -not (Test-Path (Join-Path $RootDir "VERSION"))) {
    $RootDir = (Get-Item .).FullName
}

Write-Host "Cleaning directory: $RootDir" -ForegroundColor Gray

# 3. Clean distribution and temporary directories
$dirsToClean = @(
    (Join-Path $RootDir "dist"),
    (Join-Path $RootDir "temp")
)

foreach ($dir in $dirsToClean) {
    if (Test-Path $dir) {
        Write-Host "Removing directory: $dir" -ForegroundColor Yellow
        Remove-Item -Path $dir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# 4. Clean Python cache directories
Get-ChildItem -Path $RootDir -Directory -Recurse -Filter "__pycache__" -ErrorAction SilentlyContinue | ForEach-Object {
    if ($_.FullName -notmatch "\\\.git\\") {
        Write-Host "Removing __pycache__: $($_.FullName)" -ForegroundColor Gray
        Remove-Item -Path $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# 5. Clean Python bytecode and temporary files
$filePatterns = @("*.pyc", "*.pyo", "*.tmp")
foreach ($pattern in $filePatterns) {
    Get-ChildItem -Path $RootDir -File -Recurse -Filter $pattern -ErrorAction SilentlyContinue | ForEach-Object {
        if ($_.FullName -notmatch "\\\.git\\") {
            Remove-Item -Path $_.FullName -Force -ErrorAction SilentlyContinue
        }
    }
}

Write-Host "Cleanup completed successfully." -ForegroundColor Green
