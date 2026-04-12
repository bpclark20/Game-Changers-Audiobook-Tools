param(
    [Parameter(Position = 0)]
    [string]$Path = '.',
    [string]$Filter = '*',
    [switch]$Apply
)

$ErrorActionPreference = 'Stop'

$bookTitle = Read-Host 'Enter the book title to prepend'
if ([string]::IsNullOrWhiteSpace($bookTitle)) {
    throw 'Book title cannot be empty.'
}

$files = Get-ChildItem -LiteralPath $Path -File -Filter $Filter | Sort-Object Name
if (-not $files) {
    Write-Host "No files found in '$Path' matching filter: $Filter"
    exit 0
}

$normalizedTitle = $bookTitle.Trim()

$plan = foreach ($file in $files) {
    $newName = "$normalizedTitle - $($file.Name)"

    if ($file.Name.StartsWith("$normalizedTitle - ", [System.StringComparison]::OrdinalIgnoreCase)) {
        [pscustomobject]@{
            OldFull = $file.FullName
            OldName = $file.Name
            Temp = $null
            NewName = $file.Name
            Action = 'SkipAlreadyPrefixed'
        }
        continue
    }

    [pscustomobject]@{
        OldFull = $file.FullName
        OldName = $file.Name
        Temp = "__tmp__$([guid]::NewGuid().ToString('N'))$($file.Extension)"
        NewName = $newName
        Action = 'Rename'
    }
}

$renamePlan = $plan | Where-Object Action -eq 'Rename'
$duplicateTargets = $renamePlan |
    Group-Object NewName |
    Where-Object Count -gt 1 |
    Select-Object -ExpandProperty Name

if ($duplicateTargets) {
    throw "Multiple files would be renamed to the same target: $($duplicateTargets -join ', ')"
}

if (-not $Apply) {
    Write-Host "Preview mode only. No files will be changed."
    $plan | Select-Object OldName, NewName, Action | Format-Table -AutoSize
    Write-Host ""
    Write-Host "Run again with -Apply to perform the rename."
    exit 0
}

foreach ($item in $renamePlan) {
    Rename-Item -LiteralPath $item.OldFull -NewName $item.Temp
}

foreach ($item in $renamePlan) {
    $tempFull = Join-Path (Split-Path $item.OldFull -Parent) $item.Temp
    Rename-Item -LiteralPath $tempFull -NewName $item.NewName
}

$skippedCount = ($plan | Where-Object Action -eq 'SkipAlreadyPrefixed').Count
Write-Host "Renamed $($renamePlan.Count) files successfully."
if ($skippedCount -gt 0) {
    Write-Host "Skipped $skippedCount files that already had the title prefix."
}