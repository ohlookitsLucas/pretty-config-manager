$files = @(
    'C:\Users\nony\Documents\vsco\PCM\src\Ui.ps1',
    'C:\Users\nony\Documents\vsco\PCM\src\Ui.Signatures.ps1',
    'C:\Users\nony\Documents\vsco\PCM\src\Ui.Permissions.ps1',
    'C:\Users\nony\Documents\vsco\PCM\src\Ui.Wizard.ps1',
    'C:\Users\nony\Documents\vsco\PCM\src\Ui.Extras.ps1'
)
$allOk = $true
foreach ($f in $files) {
    $errs = $null
    $null = [System.Management.Automation.Language.Parser]::ParseFile($f, [ref]$null, [ref]$errs)
    if ($errs.Count -gt 0) {
        Write-Host "FAIL: $f" -ForegroundColor Red
        foreach ($e in $errs) {
            Write-Host "  Line $($e.Extent.StartLineNumber): $($e.Message)" -ForegroundColor Yellow
        }
        $allOk = $false
    } else {
        Write-Host "OK:   $f" -ForegroundColor Green
    }
}
if ($allOk) { Write-Host "`nAll files parse OK" -ForegroundColor Cyan } else { exit 1 }
