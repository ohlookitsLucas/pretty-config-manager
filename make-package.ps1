<#
  make-package.ps1

  Creates a single self-extracting PowerShell script containing the entire
  outlookmAnAger project. The generated file is a single text PS1 you can email
  to recipients; running it will extract the project to a chosen folder.

  Usage: run this inside the project root (where this file lives):
      .\make-package.ps1

  Output: outlookmAnAger-package.ps1 in the current folder.

  Security note: The generated file contains a base64-encoded ZIP of the project.
  Recipients should verify content before running. Some mail gateways flag PS1
  attachments; if so, use an alternate delivery mechanism (internal share).
#>

Param()

Write-Host "Packaging outlookmAnAger..."

$projRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
$tmpZip = Join-Path $env:TEMP ([System.IO.Path]::GetRandomFileName() + '.zip')
$outFile = Join-Path $projRoot 'outlookmAnAger-package.ps1'

if (Test-Path $tmpZip) { Remove-Item $tmpZip -Force }

Write-Host "Creating temporary ZIP..."
Compress-Archive -Path (Join-Path $projRoot '*') -DestinationPath $tmpZip -Force

Write-Host "Encoding package..."
$bytes = [IO.File]::ReadAllBytes($tmpZip)
$b64 = [Convert]::ToBase64String($bytes)

Write-Host "Generating single-file installer: $outFile"

# Header lines for the generated script (use single-quoted strings to keep $ tokens literal)
$header = @(
    'param([string]$OutDir = "$env:USERPROFILE\Documents\outlookmAnAger")',
    '',
    'Write-Host "Extracting outlookmAnAger to: $OutDir"',
    'if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }',
    '',
    '# Embedded package (base64)'
)

Set-Content -Path $outFile -Value $header -Encoding UTF8

# Write the base64 on a single line (safe: base64 contains no single quotes)
$encLine = '$enc = ''' + $b64 + ''''
Add-Content -Path $outFile -Value $encLine -Encoding UTF8

# Footer: extraction logic
$footer = @(
    '',
    '$bytes = [Convert]::FromBase64String($enc)',
    '$tmp = Join-Path $env:TEMP ([System.IO.Path]::GetRandomFileName() + ''.zip'')',
    '[IO.File]::WriteAllBytes($tmp, $bytes)',
    'Expand-Archive -Path $tmp -DestinationPath $OutDir -Force',
    'Remove-Item $tmp -Force',
    'Write-Host "Done. Run: powershell -NoProfile -ExecutionPolicy Bypass -File \"$OutDir\\build-and-run.ps1\" to launch."'
)

Add-Content -Path $outFile -Value $footer -Encoding UTF8

Remove-Item $tmpZip -Force

Write-Host "Package created: $outFile"
Write-Host "Size: $((Get-Item $outFile).Length / 1KB) KB"
