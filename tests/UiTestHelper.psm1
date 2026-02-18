<#
  UiTestHelper.psm1
  Infrastructure for automated WPF UI testing of outlookmAnAger.

  Architecture:
  - The WPF app runs in a background STA runspace with sandbox data
  - A synchronized hashtable ($syncHash) bridges the test thread and UI thread
  - To execute code on the UI thread IN THE RUNSPACE'S SCOPE, tests write a
    command string to $syncHash._command and signal via DispatcherTimer
  - This avoids scope issues (scriptblocks from the test thread can't access
    runspace-scoped functions like Set-SelectedSigLabel)

  Usage:
    Import-Module .\tests\UiTestHelper.psm1 -Force
    $ctx = Start-SandboxApp
    Invoke-UiCommand $ctx 'Click-Button BtnNewSignature'
    Stop-SandboxApp $ctx
#>

function Start-SandboxApp {
    $testsDir = $PSScriptRoot
    $srcDir   = Join-Path $testsDir '..\src'

    $syncHash = [hashtable]::Synchronized(@{
        Window           = $null
        Ready            = $false
        Error            = $null
        _inputBoxResult  = ''
        _confirmResult   = $true
        _lastError       = ''
        _commandResult   = $null
        _commandScript   = $null
        _commandReady    = $false
    })

    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = 'STA'
    $runspace.ThreadOptions  = 'ReuseThread'
    $runspace.Open()
    $runspace.SessionStateProxy.SetVariable('syncHash', $syncHash)
    $runspace.SessionStateProxy.SetVariable('testsDir', $testsDir)
    $runspace.SessionStateProxy.SetVariable('srcDir', (Resolve-Path $srcDir).Path)

    $ps = [powershell]::Create()
    $ps.Runspace = $runspace
    $ps.AddScript({
        $ErrorActionPreference = 'Stop'
        try {
            # ── Setup sandbox ──
            $sandbox = & (Join-Path $testsDir 'Reset-TestEnvironment.ps1')

            # ── Load WPF ──
            Add-Type -AssemblyName PresentationFramework

            $xamlPath = Join-Path $srcDir 'MainWindow.xaml'
            [xml]$xaml = Get-Content -Path $xamlPath -Raw
            $reader   = New-Object System.Xml.XmlNodeReader($xaml.DocumentElement)
            $window   = [Windows.Markup.XamlReader]::Load($reader)

            # ── Import modules ──
            Import-Module (Join-Path $srcDir 'SignatureManager.psm1')  -Force
            Import-Module (Join-Path $srcDir 'PermissionsManager.psm1') -Force

            # ── Redirect to sandbox ──
            Set-SignatureManagerPaths `
                -SignaturePath $sandbox.SignaturePath `
                -RegistryBase  $sandbox.RegBase `
                -ProfilesBase  $sandbox.ProfilesBase `
                -LogFile       $sandbox.LogFile `
                -BackupDir     $sandbox.BackupDir

            # ── Inject mocks ──
            . (Join-Path $testsDir 'Mocks\New-MockOutlookNS.ps1')
            $mockNS = New-MockOutlookNS
            Set-OutlookNSFactory { $mockNS }.GetNewClosure()
            Set-ADSearchFactory  (New-MockADSearchFactory)

            # ── Load UI scripts ──
            $themeContent    = Get-Content -Path (Join-Path $srcDir 'Theme.ps1')          -Raw -Encoding UTF8
            $langContent     = Get-Content -Path (Join-Path $srcDir 'Language.ps1')       -Raw -Encoding UTF8
            $uiContent       = Get-Content -Path (Join-Path $srcDir 'Ui.ps1')             -Raw -Encoding UTF8
            $uiSigsContent   = Get-Content -Path (Join-Path $srcDir 'Ui.Signatures.ps1')  -Raw -Encoding UTF8
            $uiPermsContent  = Get-Content -Path (Join-Path $srcDir 'Ui.Permissions.ps1') -Raw -Encoding UTF8
            $uiWizContent    = Get-Content -Path (Join-Path $srcDir 'Ui.Wizard.ps1')      -Raw -Encoding UTF8
            $uiExtrasContent = Get-Content -Path (Join-Path $srcDir 'Ui.Extras.ps1')      -Raw -Encoding UTF8
            Invoke-Expression $themeContent
            Invoke-Expression $langContent
            Invoke-Expression $uiContent
            Invoke-Expression $uiSigsContent
            Invoke-Expression $uiPermsContent
            Invoke-Expression $uiWizContent
            Invoke-Expression $uiExtrasContent

            $script:SettingsPath = $sandbox.SettingsFile

            # ── Initialize UI ──
            Initialize-Ui -Window $window -ScriptRoot $srcDir

            # ── Override modal dialog functions with mock versions ──
            function script:Show-InputBox([string]$prompt, [string]$title, [string]$default = '') {
                return $syncHash._inputBoxResult
            }
            function script:Confirm-Action([string]$msg, [string]$title = 'Confirm') {
                return $syncHash._confirmResult
            }
            function script:Show-Error([string]$msg) {
                $syncHash._lastError = $msg
            }

            # ── Command execution timer ──
            # Tests write a script string to $syncHash._commandScript,
            # then set _commandReady=$true. This timer picks it up and
            # executes it IN THE RUNSPACE'S SCRIPT SCOPE, then writes
            # the result to _commandResult and clears _commandReady.
            $cmdTimer = New-Object System.Windows.Threading.DispatcherTimer
            $cmdTimer.Interval = [TimeSpan]::FromMilliseconds(50)
            $cmdTimer.add_Tick({
                if ($syncHash._commandReady -and $null -ne $syncHash._commandScript) {
                    try {
                        $syncHash._commandResult = Invoke-Expression $syncHash._commandScript
                    } catch {
                        $syncHash._commandResult = "ERROR: $_"
                    }
                    $syncHash._commandScript = $null
                    $syncHash._commandReady  = $false
                }
            })
            $cmdTimer.Start()

            # ── Dispatcher exception handler ──
            $app = [System.Windows.Application]::Current
            if ($null -eq $app) {
                $app = New-Object System.Windows.Application
                $app.ShutdownMode = 'OnExplicitShutdown'
            }
            $app.add_DispatcherUnhandledException({
                param($s, $e)
                $syncHash.Error = $e.Exception.Message
                $e.Handled = $true
            })

            # ── Signal ready once rendered ──
            $window.add_ContentRendered({
                $syncHash.Window = $window
                $syncHash.Ready  = $true
            })

            # ── Show (non-blocking) + pump messages ──
            $window.Show()
            [System.Windows.Threading.Dispatcher]::Run()
        } catch {
            $syncHash.Error = $_.Exception.Message + "`n" + $_.ScriptStackTrace
            $syncHash.Ready = $true
        }
    }) | Out-Null

    $asyncResult = $ps.BeginInvoke()

    # Wait for window to be ready
    $timeout = [datetime]::Now.AddSeconds(30)
    while (-not $syncHash.Ready -and [datetime]::Now -lt $timeout) {
        Start-Sleep -Milliseconds 100
    }

    if (-not $syncHash.Ready) {
        throw "Sandbox app did not become ready within 30 seconds."
    }
    if ($null -ne $syncHash.Error -and $null -eq $syncHash.Window) {
        throw "Sandbox app failed to start: $($syncHash.Error)"
    }

    return [PSCustomObject]@{
        SyncHash    = $syncHash
        Runspace    = $runspace
        PowerShell  = $ps
        AsyncResult = $asyncResult
        TestsDir    = $testsDir
    }
}

function Stop-SandboxApp {
    param([Parameter(Mandatory)] $Context)

    try {
        $w = $Context.SyncHash.Window
        if ($null -ne $w) {
            $closeAction = {
                $w.Close()
                [System.Windows.Threading.Dispatcher]::CurrentDispatcher.InvokeShutdown()
            }.GetNewClosure()
            $w.Dispatcher.Invoke(
                [System.Windows.Threading.DispatcherPriority]::Normal,
                [System.Action]$closeAction
            )
        }
    } catch {}

    try { $Context.PowerShell.EndInvoke($Context.AsyncResult) } catch {}
    try { $Context.PowerShell.Dispose() } catch {}
    try { $Context.Runspace.Close(); $Context.Runspace.Dispose() } catch {}

    try {
        & (Join-Path $Context.TestsDir 'Reset-TestEnvironment.ps1') -Teardown
    } catch {}
}

function Invoke-UiCommand {
    <#
      Executes a PowerShell expression string on the UI thread within the RUNSPACE'S scope.
      This means all functions (Set-SelectedSigLabel, Load-SignaturePreview, etc.),
      $script: variables, and module functions are available.
      Returns the result of the expression.
    #>
    param(
        [Parameter(Mandatory)] $Context,
        [Parameter(Mandatory)] [string] $Script,
        [int] $TimeoutMs = 10000
    )

    $sh = $Context.SyncHash
    $sh['_commandResult'] = $null
    $sh['_commandScript'] = $Script
    $sh['_commandReady']  = $true

    # Wait for the timer to pick up and execute the command
    $deadline = [datetime]::Now.AddMilliseconds($TimeoutMs)
    while ($sh['_commandReady'] -and [datetime]::Now -lt $deadline) {
        Start-Sleep -Milliseconds 25
    }

    if ($sh['_commandReady']) {
        $sh['_commandReady'] = $false
        throw "UI command timed out after ${TimeoutMs}ms: $Script"
    }

    return $sh['_commandResult']
}

# ── Convenience helpers ──────────────────────────────────────────────────────

function Click-Button {
    param(
        [Parameter(Mandatory)] $Context,
        [Parameter(Mandatory)] [string] $ButtonName
    )
    $result = Invoke-UiCommand $Context @"
`$btn = `$window.FindName('$ButtonName')
if (`$null -eq `$btn) { throw "Button '$ButtonName' not found" }
`$a = New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, `$btn)
`$btn.RaiseEvent(`$a)
"@
    if ($result -like 'ERROR:*') { throw $result }
}

function Click-SignatureRow {
    param(
        [Parameter(Mandatory)] $Context,
        [Parameter(Mandatory)] [string] $SigName
    )
    $escaped = $SigName -replace "'", "''"
    $result = Invoke-UiCommand $Context @"
`$found = `$false
`$panel = `$window.FindName('PanelInboxList')
foreach (`$child in `$panel.Children) {
    if (`$child -is [System.Windows.Controls.Border] -and `$child.Tag -eq '$escaped') {
        `$found = `$true; break
    }
}
if (-not `$found) { return 'NOT_FOUND' }
`$script:selectedSigName = '$escaped'
Set-SelectedSigLabel '$escaped'
Load-SignaturePreview '$escaped'
return 'OK'
"@
    if ($result -eq 'NOT_FOUND') {
        throw "Signature row '$SigName' not found in PanelInboxList"
    }
}

function Click-Pill {
    param(
        [Parameter(Mandatory)] $Context,
        [Parameter(Mandatory)] [string] $AccountName
    )
    $escaped = $AccountName -replace "'", "''"
    $result = Invoke-UiCommand $Context @"
`$ctrl = `$window.FindName('TxtSelectedSig')
`$sigName = if (`$null -ne `$ctrl) { `$ctrl.Text } else { `$script:selectedSigName }
`$panel = `$window.FindName('PanelCopyTargets')
foreach (`$pill in `$panel.Children) {
    if (`$pill -is [System.Windows.Controls.Border] -and
        `$null -ne `$pill.Tag -and
        `$pill.Tag.AccountName -eq '$escaped') {
        Toggle-AssignPill `$pill
        return 'OK'
    }
}
return 'NOT_FOUND'
"@
    if ($result -eq 'NOT_FOUND') {
        throw "Pill for '$AccountName' not found in PanelCopyTargets"
    }
}

function Get-ControlText {
    param(
        [Parameter(Mandatory)] $Context,
        [Parameter(Mandatory)] [string] $ControlName
    )
    return Invoke-UiCommand $Context @"
`$ctrl = `$window.FindName('$ControlName')
if (`$null -eq `$ctrl) { return `$null }
if (`$null -ne `$ctrl.Text) { return `$ctrl.Text }
if (`$null -ne `$ctrl.Content) { return "`$(`$ctrl.Content)" }
return `$null
"@
}

function Get-ControlVisibility {
    param(
        [Parameter(Mandatory)] $Context,
        [Parameter(Mandatory)] [string] $ControlName
    )
    return Invoke-UiCommand $Context @"
`$ctrl = `$window.FindName('$ControlName')
if (`$null -eq `$ctrl) { return 'NotFound' }
return `$ctrl.Visibility.ToString()
"@
}

function Wait-ForCondition {
    param(
        [Parameter(Mandatory)] $Context,
        [Parameter(Mandatory)] [string] $ConditionScript,
        [int] $TimeoutMs = 5000,
        [int] $PollMs = 100
    )
    $deadline = [datetime]::Now.AddMilliseconds($TimeoutMs)
    while ([datetime]::Now -lt $deadline) {
        $result = Invoke-UiCommand $Context $ConditionScript
        if ($result) { return $true }
        Start-Sleep -Milliseconds $PollMs
    }
    return $false
}

function Get-PillBadgeText {
    param(
        [Parameter(Mandatory)] $Context,
        [Parameter(Mandatory)] [string] $AccountName
    )
    $escaped = $AccountName -replace "'", "''"
    return Invoke-UiCommand $Context @"
`$panel = `$window.FindName('PanelCopyTargets')
foreach (`$pill in `$panel.Children) {
    if (`$pill -is [System.Windows.Controls.Border] -and
        `$null -ne `$pill.Tag -and
        `$pill.Tag.AccountName -eq '$escaped') {
        `$sp = `$pill.Child
        if (`$sp -is [System.Windows.Controls.StackPanel] -and `$sp.Children.Count -ge 2) {
            `$badge = `$sp.Children[1]
            if (`$badge.Visibility.ToString() -eq 'Visible') {
                return `$badge.Text
            }
        }
        return ''
    }
}
return `$null
"@
}

function Get-InboxListSigNames {
    param([Parameter(Mandatory)] $Context)
    return Invoke-UiCommand $Context @'
$names = @()
$panel = $window.FindName('PanelInboxList')
foreach ($child in $panel.Children) {
    if ($child -is [System.Windows.Controls.Border] -and
        $child.Tag -is [string] -and
        $child.Tag.Length -gt 0 -and
        $child.Tag -notlike '*HKCU:*' -and
        $child.Tag -notlike '*Registry::*') {
        $names += $child.Tag
    }
}
return $names
'@
}

Export-ModuleMember -Function `
    Start-SandboxApp, `
    Stop-SandboxApp, `
    Invoke-UiCommand, `
    Click-Button, `
    Click-SignatureRow, `
    Click-Pill, `
    Get-ControlText, `
    Get-ControlVisibility, `
    Wait-ForCondition, `
    Get-PillBadgeText, `
    Get-InboxListSigNames
