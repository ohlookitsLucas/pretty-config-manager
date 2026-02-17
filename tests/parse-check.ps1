$files = @(
    'C:\Users\nony\Documents\vsco\outlookmAnAger\src\Ui.ps1',
    'C:\Users\nony\Documents\vsco\outlookmAnAger\src\Theme.ps1',
    'C:\Users\nony\Documents\vsco\outlookmAnAger\src\Language.ps1'
)
foreach ($f in $files) {
    $tokens = $null
    $parseErrors = $null
    $null = [System.Management.Automation.Language.Parser]::ParseFile($f, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count -gt 0) {
        Write-Host "=== ERRORS in $f ===" -ForegroundColor Red
        foreach ($e in $parseErrors) {
            Write-Host "  Line $($e.Extent.StartLineNumber): $($e.Message)"
        }
    } else {
        Write-Host "OK: $f" -ForegroundColor Green
    }
}
