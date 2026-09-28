<#
.SYNOPSIS
    BWPConvertTTNVN Automated PowerShell Build Wrapper

.DESCRIPTION
    Wraps python scripts/build.py with optional automatic AccessVBOM enablement
    and test skipping flags.

.PARAMETER EnableAccessVBOM
    If specified and AccessVBOM = 0 (or not set), temporarily sets AccessVBOM = 1 in HKCU,
    executes the build pipeline, and guarantees restoration of the previous value in a finally block.

.PARAMETER SkipTests
    Skips the Stage 5 automated 3-tier test runner.

.PARAMETER Release
    Packages the release distribution (.zip and .sha256 checksums) in Stage 6.
#>

[CmdletBinding()]
param(
    [switch]$EnableAccessVBOM,
    [switch]$SkipTests,
    [switch]$Release
)

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptDir
Set-Location $ProjectRoot

$officeVersions = @("16.0", "15.0", "14.0", "12.0")
$registryBackups = @()

try {
    if ($EnableAccessVBOM) {
        Write-Host "Checking AccessVBOM registry configuration..." -ForegroundColor Cyan
        foreach ($ver in $officeVersions) {
            $regPath = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security"
            if (Test-Path $regPath) {
                $prop = Get-ItemProperty -Path $regPath -Name "AccessVBOM" -ErrorAction SilentlyContinue
                if ($null -ne $prop -and $prop.AccessVBOM -eq 1) {
                    Write-Host "  Office $ver : AccessVBOM is already enabled (1)." -ForegroundColor Green
                } else {
                    $originalVal = if ($null -ne $prop) { $prop.AccessVBOM } else { $null }
                    $registryBackups += @{
                        Path = $regPath
                        OriginalValue = $originalVal
                        Version = $ver
                    }
                    Set-ItemProperty -Path $regPath -Name "AccessVBOM" -Value 1 -Type DWord -Force
                    Write-Host "  Office $ver : Temporarily enabled AccessVBOM = 1 (was: $originalVal)." -ForegroundColor Yellow
                }
            }
        }
    }

    $buildScript = Join-Path $ScriptDir "build.py"
    $pyArgs = @($buildScript)
    if ($SkipTests) {
        $pyArgs += "--skip-tests"
    }
    if ($Release) {
        $pyArgs += "--release"
    }

    Write-Host "Executing build pipeline: python $($pyArgs -join ' ')" -ForegroundColor Cyan
    & python @pyArgs
    $exitCode = $LASTEXITCODE

    if ($exitCode -ne 0) {
        Write-Error "Build pipeline failed with exit code $exitCode"
        exit $exitCode
    }
}
finally {
    if ($registryBackups.Count -gt 0) {
        Write-Host "`nRestoring original AccessVBOM registry settings..." -ForegroundColor Cyan
        foreach ($item in $registryBackups) {
            if ($null -eq $item.OriginalValue) {
                Remove-ItemProperty -Path $item.Path -Name "AccessVBOM" -ErrorAction SilentlyContinue
                Write-Host "  Office $($item.Version) : Removed temporary AccessVBOM setting." -ForegroundColor Gray
            } else {
                Set-ItemProperty -Path $item.Path -Name "AccessVBOM" -Value $item.OriginalValue -Type DWord -Force
                Write-Host "  Office $($item.Version) : Restored AccessVBOM = $($item.OriginalValue)." -ForegroundColor Gray
            }
        }
    }
}
