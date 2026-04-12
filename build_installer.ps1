param(
    [string]$SourceDir = (Join-Path $PSScriptRoot "dist"),
    [string]$ScriptPath = (Join-Path $PSScriptRoot "installer.iss"),
    [string]$AppVersion = "",
    [string]$OutputDir = (Join-Path $PSScriptRoot "installer")
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $ScriptPath)) {
    throw "Installer script not found: $ScriptPath"
}

if (-not (Test-Path $SourceDir)) {
    throw "Source directory not found: $SourceDir"
}

$mainExe = Join-Path $SourceDir "JingleAllTheDay.exe"
if (-not (Test-Path $mainExe)) {
    throw "Expected bundled EXE not found: $mainExe"
}

if ([string]::IsNullOrWhiteSpace($AppVersion)) {
    $guiPath = Join-Path $PSScriptRoot "gui.py"
    if (-not (Test-Path $guiPath)) {
        throw "Cannot infer app version because gui.py was not found. Provide -AppVersion explicitly."
    }

    $guiText = Get-Content -Raw -Path $guiPath
    $versionMatch = [regex]::Match($guiText, 'APP_VERSION\s*=\s*"([^"]+)"')
    if (-not $versionMatch.Success) {
        throw "Unable to parse APP_VERSION from gui.py. Provide -AppVersion explicitly."
    }
    $AppVersion = $versionMatch.Groups[1].Value
}

function Get-IsccFromRegistry {
    $roots = @(
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )

    foreach ($root in $roots) {
        if (-not (Test-Path $root)) {
            continue
        }

        foreach ($key in (Get-ChildItem -Path $root -ErrorAction SilentlyContinue)) {
            try {
                $props = Get-ItemProperty -Path $key.PSPath -ErrorAction Stop
                if ($props.DisplayName -notlike "*Inno Setup*") {
                    continue
                }

                $displayIcon = [string]$props.DisplayIcon
                if (-not [string]::IsNullOrWhiteSpace($displayIcon)) {
                    $cleanIcon = $displayIcon.Trim('"')
                    $cleanIcon = $cleanIcon -replace ',\d+$', ''
                    $iconDir = Split-Path -Path $cleanIcon -Parent
                    if ($iconDir) {
                        $candidate = Join-Path $iconDir "ISCC.exe"
                        if (Test-Path $candidate) {
                            return $candidate
                        }
                    }
                }

                $installLocation = [string]$props.InstallLocation
                if (-not [string]::IsNullOrWhiteSpace($installLocation)) {
                    $candidate = Join-Path $installLocation "ISCC.exe"
                    if (Test-Path $candidate) {
                        return $candidate
                    }
                }
            }
            catch {
                continue
            }
        }
    }

    return $null
}

function Get-IsccFromWinGetPackages {
    $wingetPackagesRoot = Join-Path $env:LOCALAPPDATA "Microsoft\WinGet\Packages"
    if (-not (Test-Path $wingetPackagesRoot)) {
        return $null
    }

    $packageDirs = Get-ChildItem -Path $wingetPackagesRoot -Directory -Filter "JRSoftware.InnoSetup*" -ErrorAction SilentlyContinue
    foreach ($dir in $packageDirs) {
        $candidate = Join-Path $dir.FullName "ISCC.exe"
        if (Test-Path $candidate) {
            return $candidate
        }

        $toolsCandidate = Join-Path $dir.FullName "tools\ISCC.exe"
        if (Test-Path $toolsCandidate) {
            return $toolsCandidate
        }

        $deepMatch = Get-ChildItem -Path $dir.FullName -Filter "ISCC.exe" -Recurse -File -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($deepMatch) {
            return $deepMatch.FullName
        }
    }

    return $null
}

$isccCandidates = @()

if ($env:ProgramFiles -and (Test-Path $env:ProgramFiles)) {
    $isccCandidates += (Join-Path $env:ProgramFiles "Inno Setup 6\ISCC.exe")
}

if (${env:ProgramFiles(x86)} -and (Test-Path ${env:ProgramFiles(x86)})) {
    $isccCandidates += (Join-Path ${env:ProgramFiles(x86)} "Inno Setup 6\ISCC.exe")
}

if ($env:LOCALAPPDATA -and (Test-Path $env:LOCALAPPDATA)) {
    $isccCandidates += (Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 6\ISCC.exe")
    $isccCandidates += (Join-Path $env:LOCALAPPDATA "Inno Setup 6\ISCC.exe")
    $isccCandidates += (Join-Path $env:LOCALAPPDATA "Microsoft\WinGet\Links\ISCC.exe")
}

if ($env:ChocolateyInstall -and (Test-Path $env:ChocolateyInstall)) {
    $isccCandidates += (Join-Path $env:ChocolateyInstall "bin\iscc.exe")
}

$isccPath = $null
foreach ($candidate in $isccCandidates) {
    if ($candidate -and (Test-Path $candidate)) {
        $isccPath = $candidate
        break
    }
}

if (-not $isccPath) {
    $isccCommand = Get-Command iscc.exe -ErrorAction SilentlyContinue
    if ($isccCommand) {
        $isccPath = $isccCommand.Source
    }
}

if (-not $isccPath) {
    $isccPath = Get-IsccFromRegistry
}

if (-not $isccPath) {
    $isccPath = Get-IsccFromWinGetPackages
}

if (-not $isccPath) {
    throw @"
Inno Setup compiler (ISCC.exe) was not found.

Install it, then run this script again:
  winget install --exact --id JRSoftware.InnoSetup

If already installed, close and reopen VS Code (or your terminal) so PATH updates are picked up.
"@
}

$resolvedSourceDir = (Resolve-Path $SourceDir).Path
$env:JATD_SOURCE_DIR = $resolvedSourceDir

$resolvedOutputDir = $OutputDir
if (-not [System.IO.Path]::IsPathRooted($resolvedOutputDir)) {
    $resolvedOutputDir = Join-Path $PSScriptRoot $resolvedOutputDir
}
if (-not (Test-Path $resolvedOutputDir)) {
    New-Item -ItemType Directory -Path $resolvedOutputDir | Out-Null
}
$resolvedOutputDir = (Resolve-Path $resolvedOutputDir).Path

& $isccPath "/DMyAppVersion=$AppVersion" "/DMyOutputDir=$resolvedOutputDir" $ScriptPath
if ($LASTEXITCODE -ne 0) {
    throw "ISCC failed with exit code $LASTEXITCODE"
}

Write-Host "Installer build complete."
Write-Host "Version: $AppVersion"
Write-Host "Output folder: $resolvedOutputDir"
