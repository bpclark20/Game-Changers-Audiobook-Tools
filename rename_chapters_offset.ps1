param(
    [int]$Offset = 29,
    [string]$Filter = '*.wav',
    [switch]$Apply
)

$ErrorActionPreference = 'Stop'

$files = Get-ChildItem -File -Filter $Filter | Sort-Object Name
if (-not $files) {
    Write-Host "No files found matching filter: $Filter"
    exit 0
}

$plan = foreach ($f in $files) {
    if ($f.BaseName -match '^(?<num>\d+)\s*-\s*Chapter\s*(?<chap>\d+)$') {
        $newNum = [int]$Matches.num - $Offset
        $newChap = [int]$Matches.chap - $Offset

        if ($newNum -lt 1 -or $newChap -lt 1) {
            throw "Invalid result after subtracting offset ${Offset}: $($f.Name)"
        }

        [pscustomobject]@{
            OldFull = $f.FullName
            OldName = $f.Name
            Temp = "__tmp__$([guid]::NewGuid().ToString('N'))$($f.Extension)"
            NewName = ("{0:D3} - Chapter {1}{2}" -f $newNum, $newChap, $f.Extension)
        }
    }
    else {
        throw "Unexpected filename format: $($f.Name)"
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
foreach ($p in $plan) {
    Rename-Item -LiteralPath $p.OldFull -NewName $p.Temp
}

foreach ($p in $plan) {
    $tmpFull = Join-Path (Split-Path $p.OldFull -Parent) $p.Temp
    Rename-Item -LiteralPath $tmpFull -NewName $p.NewName
}

Write-Host "Renamed $($plan.Count) files successfully."
