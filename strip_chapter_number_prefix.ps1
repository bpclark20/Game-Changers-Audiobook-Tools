param(
    [Parameter(Position = 0)]
    [string]$Path = '.',
    [string]$Filter = '*.wav',
    [switch]$Recurse,
    [switch]$Apply
)

$ErrorActionPreference = 'Stop'

$titlePattern = '(?:Chapter\s*\d+|Prologue|Prolouge|Epilogue)(?:\s*-\s*Part\s*\d+)?'

$getChildItemParams = @{
    LiteralPath = $Path
    File = $true
    Filter = $Filter
}

if ($Recurse) {
    $getChildItemParams.Recurse = $true
}

$files = Get-ChildItem @getChildItemParams | Sort-Object FullName
if (-not $files) {
    Write-Host "No files found in '$Path' matching filter: $Filter"
    exit 0
}

$plan = foreach ($file in $files) {
    if ($file.BaseName -match "^(?<prefix>\d+)\s*-\s*(?<title>$titlePattern)$") {
        [pscustomobject]@{
            OldFull = $file.FullName
            OldName = $file.Name
            Temp = "__tmp__$([guid]::NewGuid().ToString('N'))$($file.Extension)"
            NewName = "$($Matches.title)$($file.Extension)"
        }
    }
    else {
        throw "Unexpected filename format: $($file.Name)"
    }
}

if (-not $Apply) {
    Write-Host "Preview mode only. No files will be changed."
    $plan | Select-Object OldName, NewName | Format-Table -AutoSize
    Write-Host ""
    Write-Host "Run again with -Apply to perform the rename."
    exit 0
}

# Two-pass rename avoids name collision/overwrite issues.
foreach ($item in $plan) {
    Rename-Item -LiteralPath $item.OldFull -NewName $item.Temp
}

foreach ($item in $plan) {
    $tempFull = Join-Path (Split-Path $item.OldFull -Parent) $item.Temp
    Rename-Item -LiteralPath $tempFull -NewName $item.NewName
}

Write-Host "Renamed $($plan.Count) files successfully."