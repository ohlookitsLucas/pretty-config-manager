<#
  Run-Tests.ps1
  One-command test runner for the outlookmAnAger test suite.

  Usage:
    .\tests\Run-Tests.ps1                           # run all tests
    .\tests\Run-Tests.ps1 -Filter 'SignatureManager' # run only matching tests
    .\tests\Run-Tests.ps1 -Verbose                   # verbose output
#>
param(
    [string]$Filter,
    [switch]$Verbose,
    [switch]$SkipUi
)

$ErrorActionPreference = 'Stop'

# Ensure Pester 5+ is available
$pester = Get-Module -ListAvailable -Name Pester | Sort-Object Version -Descending | Select-Object -First 1
if (-not $pester -or $pester.Version.Major -lt 5) {
    Write-Host 'Pester 5+ not found. Installing from PSGallery...' -ForegroundColor Yellow
    Install-Module -Name Pester -MinimumVersion 5.0 -Force -Scope CurrentUser -SkipPublisherCheck
}
Import-Module Pester -MinimumVersion 5.0 -Force

# Build Pester configuration
$config = New-PesterConfiguration
$config.Run.Path = $PSScriptRoot
$config.Run.Exit = $true
$config.Output.Verbosity = if ($Verbose) { 'Detailed' } else { 'Normal' }

if ($Filter) {
    $config.Filter.FullName = "*$Filter*"
}

if ($SkipUi) {
    $config.Run.Path = Get-ChildItem -Path $PSScriptRoot -Filter '*.Tests.ps1' |
        Where-Object { $_.Name -ne 'Ui.Tests.ps1' } |
        ForEach-Object { $_.FullName }
}

# Run
$result = Invoke-Pester -Configuration $config

# Summary
Write-Host ''
if ($result.FailedCount -gt 0) {
    Write-Host "FAILED: $($result.FailedCount) of $($result.TotalCount) tests failed." -ForegroundColor Red
} else {
    Write-Host "PASSED: All $($result.TotalCount) tests passed." -ForegroundColor Green
}
