# Kill All Extras
# Terminates any running extras background processes (MouseConfusion, CapsLockTroll, etc.)
$scripts = @('MouseConfusion', 'CapsLockTroll', 'InfinitePopup', 'ReverseKeyboard', 'AnnoyingBeeps')
$killed = 0
Get-WmiObject Win32_Process | Where-Object {
    $cmd = $_.CommandLine
    $cmd -and ($scripts | Where-Object { $cmd -like "*$_*" })
} | ForEach-Object {
    Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
    $killed++
}
Write-Host "Killed $killed process(es)." -ForegroundColor Green
