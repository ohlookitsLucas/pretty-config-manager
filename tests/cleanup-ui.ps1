$filePath = Join-Path $PSScriptRoot '..\src\Ui.ps1'
$lines = [System.IO.File]::ReadAllLines($filePath, [System.Text.Encoding]::UTF8)

$startIdx = -1
$endIdx = -1
for ($i = 0; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -match '^function Apply-Theme') { $startIdx = $i; break }
}
for ($i = $startIdx; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -match '^function Initialize-Ui') { $endIdx = $i; break }
}

if ($startIdx -ge 0 -and $endIdx -gt $startIdx) {
    Write-Host "Removing lines $($startIdx+1) through $($endIdx) (Apply-Theme to before Initialize-Ui)"
    $newLines = $lines[0..($startIdx-1)] + $lines[$endIdx..($lines.Count-1)]
    [System.IO.File]::WriteAllLines($filePath, $newLines, [System.Text.Encoding]::UTF8)
    Write-Host "Done. Removed $($endIdx - $startIdx) lines."
} else {
    Write-Host "Could not find markers: start=$startIdx end=$endIdx"
}
