param(
    [string]$OutRoot = "$env:LOCALAPPDATA\JingleAllTheDay\pyinstaller"
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$PythonExe = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
$SpecPath = Join-Path $ProjectRoot "JingleAllTheDay.spec"

if (-not (Test-Path $PythonExe)) {
    throw "Python executable not found: $PythonExe"
}

if (-not (Test-Path $SpecPath)) {
    throw "Spec file not found: $SpecPath"
}

$DistPath = Join-Path $OutRoot "dist"
$WorkPath = Join-Path $OutRoot "build"

New-Item -ItemType Directory -Force -Path $DistPath | Out-Null
New-Item -ItemType Directory -Force -Path $WorkPath | Out-Null

& $PythonExe -m PyInstaller --noconfirm --clean --distpath $DistPath --workpath $WorkPath $SpecPath
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller failed with exit code $LASTEXITCODE"
}

$BundleDir = Join-Path $DistPath "JingleAllTheDay"
if (Test-Path $BundleDir) {
    $RootExe = Join-Path $DistPath "JingleAllTheDay.exe"
    $RootInternal = Join-Path $DistPath "_internal"

    if (Test-Path $RootExe) {
        Remove-Item -Force $RootExe
    }
    if (Test-Path $RootInternal) {
        Remove-Item -Recurse -Force $RootInternal
    }

    Get-ChildItem -Force $BundleDir | ForEach-Object {
        Move-Item -Force $_.FullName -Destination $DistPath
    }
    Remove-Item -Recurse -Force $BundleDir
}

$RevLog = Join-Path $ProjectRoot "rev.log"
if (Test-Path $RevLog) {
    Copy-Item -Force $RevLog -Destination $DistPath
    Write-Host "Copied rev.log to $DistPath"
}

Write-Host "Build complete."
Write-Host "Output: $DistPath"
