$script:extrasProcs = @{}
# Theme definitions  -> Theme.ps1
# Translations       -> Language.ps1

# ═══ Settings Persistence ════════════════════════════════════════════════════
$script:SettingsPath = Join-Path $env:APPDATA 'outlookmAnAger\settings.json'

function Load-Settings {
    $defaults = [PSCustomObject]@{ Theme = 'Light'; Language = 'en' }
    if (-not (Test-Path $script:SettingsPath)) { return $defaults }
    try {
        $obj = Get-Content -Path $script:SettingsPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if (-not $script:Themes.Contains($obj.Theme)) { $obj.Theme = 'Light' }
        return $obj
    } catch {
        return $defaults
    }
}

function Save-Settings {
    param([PSCustomObject]$Settings)
    $dir = Split-Path $script:SettingsPath -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    try { $Settings | ConvertTo-Json -Depth 3 | Set-Content -Path $script:SettingsPath -Encoding UTF8 } catch {}
}

function Initialize-Ui {
    param(
        [Parameter(Mandatory=$true)] $Window,
        [Parameter(Mandatory=$true)] [string] $ScriptRoot
    )

    # -- Control references (shared/navigation only) --
    $navBtnOutlook       = $Window.FindName('NavBtnOutlook')
    $navBtnSkype         = $Window.FindName('NavBtnSkype')
    $navBtnMUD           = $Window.FindName('NavBtnMUD')
    $navBtnExtras        = $Window.FindName('NavBtnExtras')
    $navBtnSettings      = $Window.FindName('NavBtnSettings')

    $panelOutlookContent     = $Window.FindName('PanelOutlookContent')
    $panelSkypeContent       = $Window.FindName('PanelSkypeContent')
    $panelMUDContent         = $Window.FindName('PanelMUDContent')
    $panelAboutContent       = $Window.FindName('PanelAboutContent')
    $panelExtrasContent      = $Window.FindName('PanelExtrasContent')

    # Settings popup
    $popSettingsMenu         = $Window.FindName('PopSettingsMenu')
    $script:popSettingsMenuRef = $popSettingsMenu

    # Outlook dropdown + sub-views
    $popOutlookMenu          = $Window.FindName('PopOutlookMenu')
    $navMenuSignatures       = $Window.FindName('NavMenuSignatures')
    $navMenuPermissions      = $Window.FindName('NavMenuPermissions')
    $navMenuDev              = $Window.FindName('NavMenuDev')
    $panelSignaturesView     = $Window.FindName('PanelSignaturesView')
    $panelPermissionsView    = $Window.FindName('PanelPermissionsView')
    $panelDevView            = $Window.FindName('PanelDevView')

    # Settings tab
    $lbThemeSelector     = $Window.FindName('LbThemeSelector')
    $cbLanguageSelector  = $Window.FindName('CbLanguageSelector')

    # Status bar
    $txtStatus           = $Window.FindName('TxtStatus')

    # -- State --
    $script:currentSig        = $null
    $script:selectedSigName   = $null    # sig name currently selected for preview/assign
    $script:selectedAccountKey= $null    # RegistryPath of the highlighted account card
    $script:extrasUnlocked    = $false
    $script:outlookView       = 'Signatures'   # which sub-view is active: 'Signatures' or 'Permissions'
    $script:dlgResult      = $false
    $script:windowRef      = $Window
    $script:AppSettings    = Load-Settings

    # Panel refs needed by Set-OutlookView (stored at script scope so the function works)
    $script:popOutlookMenuRef    = $popOutlookMenu
    $script:panelSignaturesRef   = $panelSignaturesView
    $script:panelPermissionsRef  = $panelPermissionsView
    $script:panelDevRef          = $panelDevView
    $script:navMenuDevRef        = $navMenuDev

    # Theme/language selector refs (needed by Show-UnlockDialog + BtnLockExtras before selectors wired)
    $script:lbThemeSelectorRef    = $lbThemeSelector
    $script:cbLanguageSelectorRef = $cbLanguageSelector

    # -- Navigation system --
    $script:navButtons = @($navBtnOutlook, $navBtnSkype, $navBtnMUD, $navBtnExtras)
    $script:contentPanels = @{
        'Outlook'  = $panelOutlookContent
        'Skype'    = $panelSkypeContent
        'MUD'      = $panelMUDContent
        'About'    = $panelAboutContent
        'Extras'   = $panelExtrasContent
    }
    $script:navBtnExtrasRef = $navBtnExtras
    $script:navBtnSkypeRef  = $navBtnSkype
    $script:navBtnMUDRef    = $navBtnMUD
    $script:devModeUnlocked = $false
    $script:activeSection = 'Outlook'

    function script:Set-ActiveSection([string]$section) {
        # Hide all content panels
        foreach ($key in $script:contentPanels.Keys) {
            $script:contentPanels[$key].Visibility = 'Collapsed'
        }
        # Show the requested panel
        $panelKey = switch ($section) {
            'Outlook'  { 'Outlook' }
            'Skype'    { 'Skype' }
            'MUD'      { 'MUD' }
            'About'    { 'About' }
            'Extras'   { 'Extras' }
            default    { 'Outlook' }
        }
        $script:contentPanels[$panelKey].Visibility = 'Visible'

        # Update button styles — deselect all, select the active one
        foreach ($btn in $script:navButtons) {
            if ($null -eq $btn) { continue }
            $btn.SetResourceReference(
                [System.Windows.Controls.Control]::ForegroundProperty, 'TextSecondaryBrush')
            $btn.BorderBrush = [System.Windows.Media.Brushes]::Transparent
        }

        $activeBtn = switch ($section) {
            'Outlook'  { $script:navButtons[0] }
            'Skype'    { $script:navButtons[1] }
            'MUD'      { $script:navButtons[2] }
            'Extras'   { $script:navButtons[3] }
            default    { $null }
        }

        if ($null -ne $activeBtn) {
            $activeBtn.SetResourceReference(
                [System.Windows.Controls.Control]::ForegroundProperty, 'TextPrimaryBrush')
            $activeBtn.SetResourceReference(
                [System.Windows.Controls.Control]::BorderBrushProperty, 'AccentBrush')
        }

        $script:activeSection = $section
    }

    function script:Set-OutlookView([string]$view) {
        $script:panelSignaturesRef.Visibility  = 'Collapsed'
        $script:panelPermissionsRef.Visibility = 'Collapsed'
        $script:panelDevRef.Visibility         = 'Collapsed'
        switch ($view) {
            'Signatures'  { $script:panelSignaturesRef.Visibility  = 'Visible' }
            'Permissions' { $script:panelPermissionsRef.Visibility = 'Visible' }
            'Dev'         { $script:panelDevRef.Visibility         = 'Visible' }
        }
        $script:outlookView = $view
    }

    # -- Helpers --

    function script:Set-Status([string]$msg) {
        $script:txtStatusRef.Text = $msg
        $script:statusBarRef.Visibility = 'Visible'
        if ($null -ne $script:statusHideTimer) {
            $script:statusHideTimer.Stop()
            $script:statusHideTimer.Start()
        }
    }

    function script:Show-InputBox([string]$prompt, [string]$title, [string]$default = '') {
        Add-Type -AssemblyName Microsoft.VisualBasic
        return [Microsoft.VisualBasic.Interaction]::InputBox($prompt, $title, $default)
    }

    function script:Confirm-Action([string]$msg, [string]$title = 'Confirm') {
        return [System.Windows.MessageBox]::Show($msg, $title,
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning) -eq [System.Windows.MessageBoxResult]::Yes
    }

    function script:Show-Error([string]$msg) {
        [System.Windows.MessageBox]::Show($msg, 'Error',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error) | Out-Null
    }

    # -- Status bar setup --
    $script:txtStatusRef = $txtStatus
    $script:statusBarRef = $Window.FindName('StatusBar')

    # Auto-hide timer for status bar (collapses after 2s, content expands to fill)
    $script:statusHideTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:statusHideTimer.Interval = [TimeSpan]::FromSeconds(2)
    $script:statusHideTimer.Add_Tick({
        $script:statusHideTimer.Stop()
        $script:statusBarRef.Visibility = 'Collapsed'
    })

    # -- Navigation button click handlers --

    # Outlook button: hover opens dropdown (no click action — only the menu items are clickable)
    $navBtnOutlook.Add_MouseEnter({
        if ($null -ne $script:olCloseTimer) { $script:olCloseTimer.Stop() }
        $script:popOutlookMenuRef.IsOpen = $true
    })
    # Close dropdown when mouse leaves button+popup area (delay so user can traverse the gap)
    $script:olCloseTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:olCloseTimer.Interval = [TimeSpan]::FromMilliseconds(400)
    $script:olCloseTimer.Add_Tick({
        $script:olCloseTimer.Stop()
        $script:popOutlookMenuRef.IsOpen = $false
    })
    $navBtnOutlook.Add_MouseLeave({
        $script:olCloseTimer.Start()
    })
    # Cancel close when mouse enters the popup content (events on the child Border, not the Popup)
    $popOutlookMenuOuter = $popOutlookMenu.Child   # the outer transparent Border
    $popOutlookMenuOuter.Add_MouseEnter({
        if ($null -ne $script:olCloseTimer) { $script:olCloseTimer.Stop() }
    })
    $popOutlookMenuOuter.Add_MouseLeave({
        $script:olCloseTimer.Start()
    })
    # Dropdown menu items
    $navMenuSignatures.Add_Click({
        $script:popOutlookMenuRef.IsOpen = $false
        Set-ActiveSection 'Outlook'
        Set-OutlookView 'Signatures'
    })
    $navMenuPermissions.Add_Click({
        $script:popOutlookMenuRef.IsOpen = $false
        Set-ActiveSection 'Outlook'
        Set-OutlookView 'Permissions'
    })
    $navMenuDev.Add_Click({
        $script:popOutlookMenuRef.IsOpen = $false
        Set-ActiveSection 'Outlook'
        Set-OutlookView 'Dev'
    })

    $navBtnSkype.Add_Click({ Set-ActiveSection 'Skype' })
    $navBtnMUD.Add_Click({ Set-ActiveSection 'MUD' })
    $navBtnExtras.Add_Click({ Set-ActiveSection 'Extras' })

    # Settings button: hover opens dropdown overlay (does not switch content panels)
    $navBtnSettings.Add_MouseEnter({
        if ($null -ne $script:stCloseTimer) { $script:stCloseTimer.Stop() }
        $script:popSettingsMenuRef.IsOpen = $true
    })
    $script:stCloseTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:stCloseTimer.Interval = [TimeSpan]::FromMilliseconds(400)
    $script:stCloseTimer.Add_Tick({
        $script:stCloseTimer.Stop()
        $script:popSettingsMenuRef.IsOpen = $false
    })
    $navBtnSettings.Add_MouseLeave({ $script:stCloseTimer.Start() })
    $popSettingsMenuOuter = $popSettingsMenu.Child
    $popSettingsMenuOuter.Add_MouseEnter({
        if ($null -ne $script:stCloseTimer) { $script:stCloseTimer.Stop() }
    })
    $popSettingsMenuOuter.Add_MouseLeave({ $script:stCloseTimer.Start() })

    # About button inside settings dropdown
    $btnAbout = $Window.FindName('BtnAbout')
    $btnAbout.Add_MouseLeftButtonUp({
        $script:popSettingsMenuRef.IsOpen = $false
        Set-ActiveSection 'About'
    })

    # Readme link
    $lnkReadme = $Window.FindName('LnkReadme')
    if ($null -ne $lnkReadme) {
        $lnkReadme.Add_MouseLeftButtonUp(({
            $readmePath = Join-Path $ScriptRoot '..\README.md'
            if (Test-Path $readmePath) {
                Start-Process notepad.exe -ArgumentList "`"$readmePath`""
            }
        }).GetNewClosure())
    }

    # Set initial section
    Set-ActiveSection 'Outlook'
    Set-OutlookView 'Signatures'
    # Start with status bar hidden — it appears briefly when Set-Status is called
    $script:statusBarRef.Visibility = 'Collapsed'

    $btnKillAll = $Window.FindName('BtnKillAll')
    $btnKillAll.Add_Click({
        $killPath = Join-Path $ScriptRoot "extras\KillAll.ps1"
        if (Test-Path $killPath) {
            Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$killPath`"" -WindowStyle Hidden
        }
        Set-Status 'All troll processes killed'
    })

    # ── Delete All Signatures (Utilities) ────────────────────────────────────
    $btnDeleteAllSigs = $Window.FindName('BtnDeleteAllSigs')
    if ($null -ne $btnDeleteAllSigs) {
        $btnDeleteAllSigs.Add_Click({
            $allSigs = Get-Signatures
            $count   = $allSigs.Count
            if ($count -eq 0) {
                [System.Windows.MessageBox]::Show(
                    'No local signatures found.',
                    'Delete All Signatures',
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Information) | Out-Null
                return
            }
            $confirm = [System.Windows.MessageBox]::Show(
                "This will permanently delete all $count local signature(s) and clear all mailbox signature assignments.`n`nThis cannot be undone. Continue?",
                'Delete All Signatures',
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Warning)
            if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) { return }

            $errors = @()
            foreach ($name in $allSigs) {
                try { Remove-Signature -Name $name } catch { $errors += "${name}: $_" }
            }

            # Clear all registry assignments
            try {
                $assignments = Get-SignatureAssignments
                foreach ($a in $assignments) {
                    Set-SignatureAssignment -RegistryPath $a.RegistryPath -NewSignature '' -ReplySignature ''
                }
            } catch {}

            # Refresh signatures list
            try { & $script:refreshInboxList } catch {}
            $script:currentSig = $null
            $script:selectedSigName = $null
            Set-SelectedSigLabel '(none selected)'
            $script:sigTxtSignatureInfo.Text = 'Select a mailbox or signature to preview.'
            $script:previewBrowserRef.Visibility = 'Collapsed'
            if ($null -ne $script:previewEmptyHintRef) { $script:previewEmptyHintRef.Visibility = 'Visible' }

            if ($errors.Count -gt 0) {
                [System.Windows.MessageBox]::Show(
                    "Done with $($errors.Count) error(s):`n" + ($errors -join "`n"),
                    'Delete All Signatures',
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Warning) | Out-Null
            } else {
                [System.Windows.MessageBox]::Show(
                    "All $count signature(s) deleted successfully.",
                    'Delete All Signatures',
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Information) | Out-Null
            }
        })
    }

    $btnLockExtras = $Window.FindName('BtnLockExtras')
    $btnLockExtras.Add_Click({
        Stop-RetroHueTimer
        # Hide extras and switch back to Outlook
        $script:extrasUnlocked = $false
        $script:navBtnExtrasRef.Visibility = 'Collapsed'
        Set-ActiveSection 'Outlook'
        # Remove Retro from the theme selector
        for ($i = $script:lbThemeSelectorRef.Items.Count - 1; $i -ge 0; $i--) {
            if ($script:lbThemeSelectorRef.Items[$i].Tag -eq 'Retro') {
                $script:lbThemeSelectorRef.Items.RemoveAt($i); break
            }
        }
        # Reset theme to Dark
        Apply-Theme -Window $script:windowRef -ThemeName 'Dark'
        $script:AppSettings.Theme = 'Dark'
        Save-Settings -Settings $script:AppSettings
        for ($i = 0; $i -lt $script:lbThemeSelectorRef.Items.Count; $i++) {
            if ($script:lbThemeSelectorRef.Items[$i].Tag -eq 'Dark') {
                $script:lbThemeSelectorRef.SelectedIndex = $i; break
            }
        }
        Set-Status (Get-Str 'StatusExtrasLocked')
    })

    # ── Unlock dialog (called from About section rapid-click) ────────────────
    function script:Show-UnlockDialog {
        if ($script:extrasUnlocked -and $script:devModeUnlocked) { return }

        $t = $script:Themes[$script:AppSettings.Theme]
        $cBg       = $t.BgBrush
        $cSurface  = $t.SurfaceBrush
        $cBorder   = $t.BorderBrush
        $cAccent   = $t.AccentBrush
        $cFgPri    = $t.TextPrimaryBrush
        $cFgSec    = $t.TextSecondaryBrush
        $cOnAccent = $t.TextOnAccentBrush

        $dlgXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="" Width="380" SizeToContent="Height"
        Background="$cBg" Foreground="$cFgPri"
        FontFamily="Segoe UI" FontSize="13"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        FocusManager.FocusedElement="{Binding ElementName=TxtInput}">
  <StackPanel Margin="20,16,20,16">
    <Border Background="$cSurface" CornerRadius="6" BorderBrush="$cBorder" BorderThickness="1" Padding="12,10" Margin="0,0,0,12">
      <TextBlock TextWrapping="Wrap" TextAlignment="Center" FontSize="12"
                 Foreground="$cFgSec" LineHeight="18"
                 Text="Are you sure you want to enter debug mode?"/>
    </Border>
    <TextBox Name="TxtInput"
             Background="$cSurface" Foreground="$cFgPri" BorderBrush="$cBorder" BorderThickness="1"
             Padding="8,5" FontSize="13"/>
    <Grid Margin="0,10,0,0">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="8"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>
      <Button Name="BtnCancel" Grid.Column="0" Content="Cancel"
              Background="$cSurface" Foreground="$cFgPri" BorderBrush="$cBorder" BorderThickness="1"
              Padding="0,6" FontSize="12"/>
      <Button Name="BtnUnlock" Grid.Column="2" Content="Proceed"
              Background="$cAccent" Foreground="$cOnAccent" BorderThickness="0"
              Padding="0,6" FontSize="12" FontWeight="SemiBold"/>
    </Grid>
  </StackPanel>
</Window>
"@
        $dlgReader = New-Object System.Xml.XmlNodeReader([xml]$dlgXaml)
        $dlg       = [Windows.Markup.XamlReader]::Load($dlgReader)
        $txtInput  = $dlg.FindName('TxtInput')
        $btnUnlock = $dlg.FindName('BtnUnlock')
        $btnCancel = $dlg.FindName('BtnCancel')

        $script:dlgResult = $false
        $btnUnlock.Add_Click({ $script:dlgResult = $true;  $dlg.Close() })
        $btnCancel.Add_Click({ $script:dlgResult = $false; $dlg.Close() })
        $txtInput.Add_KeyDown({
            param($s, $e)
            if ($e.Key -eq 'Return') { $script:dlgResult = $true; $dlg.Close() }
        })

        $dlg.ShowDialog() | Out-Null
        $entered = $txtInput.Text

        if ($script:dlgResult -and $entered -ne '' -and $entered -ne 'itrustlucas' -and $entered -ne 'yes') {
            # Wrong password — chirps + confetti
            # Resolve WAV path once; chirps fired from the DispatcherTimer (STA thread) via Play()
            $chirpWav = @('C:\Windows\Media\ding.wav','C:\Windows\Media\Windows Ding.wav','C:\Windows\Media\chord.wav') |
                Where-Object { Test-Path $_ } | Select-Object -First 1
            $chirpPlayer = if ($chirpWav) { [System.Media.SoundPlayer]::new($chirpWav) } else { $null }

            # Size and position confetti over the main window
            $mainLeft   = $script:windowRef.Left
            $mainTop    = $script:windowRef.Top
            $mainWidth  = $script:windowRef.ActualWidth
            $mainHeight = $script:windowRef.ActualHeight

            # Confetti window — ShowDialog pumps its own message loop so the timer runs
            $confettiXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="" Width="$mainWidth" Height="$mainHeight"
        WindowStyle="None" AllowsTransparency="True" Background="Transparent"
        WindowStartupLocation="Manual" Topmost="True"
        ShowInTaskbar="False">
  <Canvas Name="ConfettiCanvas"/>
</Window>
"@
            $confettiReader = New-Object System.Xml.XmlNodeReader([xml]$confettiXaml)
            $confettiWin    = [Windows.Markup.XamlReader]::Load($confettiReader)
            $confettiWin.Left = $mainLeft
            $confettiWin.Top  = $mainTop
            $canvas         = $confettiWin.FindName('ConfettiCanvas')

            $colors = @('#E74C3C','#F39C12','#2ECC71','#3498DB','#9B59B6','#1ABC9C','#E91E63','#FF5722')
            $rng    = New-Object System.Random
            $pieces = [System.Collections.Generic.List[hashtable]]::new()

            for ($p = 0; $p -lt 80; $p++) {
                $rect = New-Object System.Windows.Shapes.Rectangle
                $rect.Width  = $rng.Next(7, 16)
                $rect.Height = $rng.Next(4, 10)
                $rect.Fill   = [System.Windows.Media.BrushConverter]::new().ConvertFromString($colors[$rng.Next($colors.Count)])
                $rect.RenderTransform = New-Object System.Windows.Media.RotateTransform($rng.Next(0, 360))
                $x = $rng.NextDouble() * $mainWidth
                [System.Windows.Controls.Canvas]::SetLeft($rect, $x)
                [System.Windows.Controls.Canvas]::SetTop($rect, -20)
                $canvas.Children.Add($rect) | Out-Null
                $pieces.Add(@{
                    Rect = $rect
                    X    = $x
                    Y    = [double]($rng.Next(-60, 0))
                    VX   = ($rng.NextDouble() - 0.5) * 2.5
                    VY   = $rng.NextDouble() * 4 + 2
                    Rot  = [double]($rng.Next(0, 360))
                    RotV = ($rng.NextDouble() - 0.5) * 8
                }) | Out-Null
            }

            $timer          = New-Object System.Windows.Threading.DispatcherTimer
            $timer.Interval = [TimeSpan]::FromMilliseconds(16)
            $elapsed        = [ref] 0
            $chirpsPlayed   = [ref] 0
            $timerRef       = [ref] $timer
            $winRef         = [ref] $confettiWin
            $playerRef      = [ref] $chirpPlayer

            $timer.Add_Tick(({
                $elapsed.Value += 16
                # Fire a chirp at 0 ms, 350 ms, 700 ms
                $chirpAt = @(0, 350, 700)
                if ($chirpsPlayed.Value -lt 3) {
                    $due = $chirpAt[$chirpsPlayed.Value]
                    if ($elapsed.Value -ge $due) {
                        if ($null -ne $playerRef.Value) { $playerRef.Value.Play() }
                        $chirpsPlayed.Value++
                    }
                }
                foreach ($piece in $pieces) {
                    $piece.Y   += $piece.VY
                    $piece.X   += $piece.VX
                    $piece.Rot += $piece.RotV
                    [System.Windows.Controls.Canvas]::SetLeft($piece.Rect, $piece.X)
                    [System.Windows.Controls.Canvas]::SetTop($piece.Rect,  $piece.Y)
                    $piece.Rect.RenderTransform = New-Object System.Windows.Media.RotateTransform($piece.Rot)
                }
                if ($elapsed.Value -ge 2500) {
                    $timerRef.Value.Stop()
                    $winRef.Value.Close()
                }
            }).GetNewClosure())

            $timer.Start()
            $confettiWin.ShowDialog() | Out-Null
        }

        if ($script:dlgResult -and $entered -eq 'yes' -and -not $script:devModeUnlocked) {
            $script:devModeUnlocked = $true
            $script:navBtnSkypeRef.Visibility = 'Visible'
            $script:navBtnMUDRef.Visibility   = 'Visible'
            $script:navMenuDevRef.Visibility  = 'Visible'
            Set-Status (Get-Str 'StatusDevUnlocked')
        }

        if ($script:dlgResult -and $entered -eq 'itrustlucas') {
            try {
                $script:extrasUnlocked = $true
                $script:navBtnExtrasRef.Visibility = 'Visible'
                Set-ActiveSection 'Extras'
                Set-Status (Get-Str 'StatusExtrasUnlocked')
                # Unlock the Retro theme and switch to it
                $alreadyAdded = $false
                for ($i = 0; $i -lt $script:lbThemeSelectorRef.Items.Count; $i++) {
                    if ($script:lbThemeSelectorRef.Items[$i].Tag -eq 'Retro') { $alreadyAdded = $true; break }
                }
                if (-not $alreadyAdded) {
                    $retroBlob            = New-Object System.Windows.Controls.Border
                    $retroBlob.Width      = 26
                    $retroBlob.Height     = 26
                    $retroBlob.CornerRadius = [System.Windows.CornerRadius]::new(6)
                    $retroBlob.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#FF2D78')
                    $retroItem            = New-Object System.Windows.Controls.ListBoxItem
                    $retroItem.Content    = $retroBlob
                    $retroItem.Tag        = 'Retro'
                    $retroItem.ToolTip    = 'Retro'
                    $retroItem.Padding    = [System.Windows.Thickness]::new(0)
                    $retroItem.Margin     = [System.Windows.Thickness]::new(0, 0, 6, 0)
                    $script:lbThemeSelectorRef.Items.Add($retroItem) | Out-Null
                }
                Apply-Theme -Window $script:windowRef -ThemeName 'Retro'
                Start-RetroHueTimer -Window $script:windowRef
                $script:AppSettings.Theme = 'Retro'
                Save-Settings -Settings $script:AppSettings
                for ($i = 0; $i -lt $script:lbThemeSelectorRef.Items.Count; $i++) {
                    if ($script:lbThemeSelectorRef.Items[$i].Tag -eq 'Retro') {
                        $script:lbThemeSelectorRef.SelectedIndex = $i; break
                    }
                }
            } catch {
                Set-Status (Get-Str 'StatusUnlockError' $_)
            }
        }
    }  # end Show-UnlockDialog
    $script:showUnlockDialog = Get-Item Function:\Show-UnlockDialog

    # Arrow sequence triggers the unlock dialog
    # ↑ ↑ ↓ ↓ ← → ← →
    $script:konamiSequence = @('Up','Up','Down','Down','Left','Right','Left','Right')
    $script:konamiNavKeys  = @('Up','Down','Left','Right')
    $script:konamiProgress = 0
    $Window.Add_PreviewKeyDown({
        param($s, $e)
        $key      = $e.Key.ToString()
        $expected = $script:konamiSequence[$script:konamiProgress]
        if ($key -eq $expected) {
            $script:konamiProgress++
            if ($script:konamiNavKeys -contains $key) { $e.Handled = $true }
            if ($script:konamiProgress -eq $script:konamiSequence.Length) {
                $script:konamiProgress = 0
                & $script:showUnlockDialog
            }
        } else {
            $script:konamiProgress = if ($key -eq $script:konamiSequence[0]) { 1 } else { 0 }
            if ($script:konamiNavKeys -contains $key -and $script:konamiProgress -gt 0) { $e.Handled = $true }
        }
    })

    # -- Mailbox preload (shared data for Permissions + Wizard tabs) --
    $script:permAllAccounts = @()
    try { $script:permAllAccounts = @(Get-SignedInAccounts) } catch {}

    # -- Initialize tab-specific UI --
    Initialize-SignaturesTab  -Window $Window -ScriptRoot $ScriptRoot
    Initialize-PermissionsTab -Window $Window -ScriptRoot $ScriptRoot
    Initialize-WizardTab      -Window $Window -ScriptRoot $ScriptRoot
    Initialize-ExtrasTab      -Window $Window -ScriptRoot $ScriptRoot

    # -- Settings: apply saved theme and pre-select controls --

    # Restore unlocked state if Retro was previously unlocked
    if ($script:AppSettings.Theme -eq 'Retro') {
        $script:extrasUnlocked = $true
        $script:navBtnExtrasRef.Visibility = 'Visible'
        $retroBlob            = New-Object System.Windows.Controls.Border
        $retroBlob.Width      = 26
        $retroBlob.Height     = 26
        $retroBlob.CornerRadius = [System.Windows.CornerRadius]::new(6)
        $retroBlob.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#FF2D78')
        $retroItem            = New-Object System.Windows.Controls.ListBoxItem
        $retroItem.Content    = $retroBlob
        $retroItem.Tag        = 'Retro'
        $retroItem.ToolTip    = 'Retro'
        $retroItem.Padding    = [System.Windows.Thickness]::new(0)
        $retroItem.Margin     = [System.Windows.Thickness]::new(0, 0, 6, 0)
        $lbThemeSelector.Items.Add($retroItem) | Out-Null
    }

    Apply-Theme -Window $Window -ThemeName $script:AppSettings.Theme
    if ($script:AppSettings.Theme -eq 'Retro') { Start-RetroHueTimer -Window $Window }

    # Pre-select saved theme
    for ($i = 0; $i -lt $lbThemeSelector.Items.Count; $i++) {
        if ($lbThemeSelector.Items[$i].Tag -eq $script:AppSettings.Theme) {
            $lbThemeSelector.SelectedIndex = $i; break
        }
    }

    # Pre-select saved language in ComboBox
    for ($i = 0; $i -lt $cbLanguageSelector.Items.Count; $i++) {
        if ($cbLanguageSelector.Items[$i].Tag -eq $script:AppSettings.Language) {
            $cbLanguageSelector.SelectedIndex = $i; break
        }
    }

    Apply-Lang -Window $Window

    $lbThemeSelector.Add_SelectionChanged({
        param($s, $e)
        if ($e.OriginalSource -ne $script:lbThemeSelectorRef) { return }
        $sel = $script:lbThemeSelectorRef.SelectedItem
        if ($null -eq $sel) { return }
        $themeName = $sel.Tag
        Stop-RetroHueTimer
        Apply-Theme -Window $script:windowRef -ThemeName $themeName
        if ($themeName -eq 'Retro') { Start-RetroHueTimer -Window $script:windowRef }
        $script:AppSettings.Theme = $themeName
        Save-Settings -Settings $script:AppSettings
    })

    $cbLanguageSelector.Add_SelectionChanged({
        param($s, $e)
        $sel = $script:cbLanguageSelectorRef.SelectedItem
        if ($null -eq $sel) { return }
        $script:AppSettings.Language = $sel.Tag
        Save-Settings -Settings $script:AppSettings
        Apply-Lang -Window $script:windowRef
    })
}
