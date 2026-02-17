$script:extrasProcs = @{}

# ═══ Theme Definitions ═══════════════════════════════════════════════════════
# To add a theme: add an entry here and a matching ListBoxItem in MainWindow.xaml.
$script:Themes = [ordered]@{
    'Dark' = @{
        BgBrush            = '#1E1E2E'
        SurfaceBrush       = '#2A2A3C'
        SurfaceHoverBrush  = '#33334A'
        BorderBrush        = '#3E3E55'
        AccentBrush        = '#7C6FE0'
        AccentHoverBrush   = '#9589E8'
        AccentSubtleBrush  = '#2E2B4A'
        TextPrimaryBrush   = '#EAEAEF'
        TextSecondaryBrush = '#9898A8'
        TextOnAccentBrush  = '#FFFFFF'
    }
    'Light' = @{
        BgBrush            = '#F0F0F5'
        SurfaceBrush       = '#FFFFFF'
        SurfaceHoverBrush  = '#E4E4EE'
        BorderBrush        = '#C8C8D8'
        AccentBrush        = '#5B50C8'
        AccentHoverBrush   = '#6A5FD8'
        AccentSubtleBrush  = '#DDDAF5'
        TextPrimaryBrush   = '#111120'
        TextSecondaryBrush = '#55556A'
        TextOnAccentBrush  = '#FFFFFF'
    }
    'Crimson' = @{
        BgBrush            = '#1A1A1A'
        SurfaceBrush       = '#2B2B2B'
        SurfaceHoverBrush  = '#363636'
        BorderBrush        = '#4A4A4A'
        AccentBrush        = '#C0392B'
        AccentHoverBrush   = '#E74C3C'
        AccentSubtleBrush  = '#3A1A1A'
        TextPrimaryBrush   = '#E8E8E8'
        TextSecondaryBrush = '#999999'
        TextOnAccentBrush  = '#FFFFFF'
    }
}

# ═══ Settings Persistence ════════════════════════════════════════════════════
$script:SettingsPath = Join-Path $env:APPDATA 'outlookmAnAger\settings.json'

function Load-Settings {
    $defaults = [PSCustomObject]@{ Theme = 'Dark'; Language = 'en' }
    if (-not (Test-Path $script:SettingsPath)) { return $defaults }
    try {
        $obj = Get-Content -Path $script:SettingsPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if (-not $script:Themes.Contains($obj.Theme)) { $obj.Theme = 'Dark' }
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

function Apply-Theme {
    param($Window, [string]$ThemeName)
    if (-not $script:Themes.Contains($ThemeName)) { return }
    $colors = $script:Themes[$ThemeName]
    foreach ($key in $colors.Keys) {
        $color = [System.Windows.Media.ColorConverter]::ConvertFromString($colors[$key])
        $brush = $Window.Resources[$key]
        if ($null -ne $brush -and $brush -is [System.Windows.Media.SolidColorBrush] -and -not $brush.IsFrozen) {
            $brush.Color = [System.Windows.Media.Color] $color
        } else {
            $nb = [System.Windows.Media.SolidColorBrush](New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color] $color))
            $Window.Resources[$key] = $nb
        }
    }
    # When no signature is selected, keep browser hidden so PanelPreview's
    # DynamicResource background shows through instead of the browser's white default.
    # When a sig with a missing file is selected, re-render the error page with new colors.
    if ($null -ne $script:previewBrowserRef -and $null -ne $script:currentSig) {
        $htmlPath = Get-SignatureHtmlPath -Name $script:currentSig
        if (-not (Test-Path $htmlPath)) {
            $bgHex = $script:Themes[$ThemeName].BgBrush
            $fgHex = $script:Themes[$ThemeName].TextSecondaryBrush
            $script:previewBrowserRef.NavigateToString(
                "<!DOCTYPE html><html style=`"background:$bgHex`"><body style=`"background:$bgHex;color:$fgHex;font-family:Segoe UI;padding:16px`">No HTML file found.</body></html>")
        }
    }
}

function Initialize-Ui {
    param(
        [Parameter(Mandatory=$true)] $Window,
        [Parameter(Mandatory=$true)] [string] $ScriptRoot
    )

    # -- Control references --
    # Signatures tab
    $btnNew              = $Window.FindName('BtnNewSignature')
    $btnRename           = $Window.FindName('BtnRenameSignature')
    $btnDelete           = $Window.FindName('BtnDeleteSignature')
    $previewBrowser      = $Window.FindName('PreviewBrowser')
    $script:previewBrowserRef = $previewBrowser
    $txtSignatureInfo    = $Window.FindName('TxtSignatureInfo')
    $panelInboxList      = $Window.FindName('PanelInboxList')
    $panelCopyTargets    = $Window.FindName('PanelCopyTargets')
    $txtSelectedSig      = $Window.FindName('TxtSelectedSig')
    $btnAssignSig        = $Window.FindName('BtnAssignSig')

    # Permissions tab
    $tbMailboxSearch  = $Window.FindName('TbMailboxSearch')
    $lbMailboxes      = $Window.FindName('LbMailboxes')
    $lbFolders        = $Window.FindName('LbFolders')
    $dgCurrentPerms   = $Window.FindName('DgCurrentPerms')
    $txtAddUser       = $Window.FindName('TxtAddUser')
    $popAddUser       = $Window.FindName('PopAddUser')
    $lbAdResults      = $Window.FindName('LbAdResults')
    $cbPermLevel      = $Window.FindName('CbPermLevel')
    $btnSavePerm      = $Window.FindName('BtnSavePerm')
    $btnRemovePerm    = $Window.FindName('BtnRemovePerm')
    $txtPermStatus    = $Window.FindName('TxtPermStatus')
    $txtFolderHint    = $Window.FindName('TxtFolderHint')
    $txtFoldersHint   = $Window.FindName('TxtFoldersHint')
    $panelPermRight   = $Window.FindName('PanelPermRight')
    $btnRefreshPerm   = $Window.FindName('BtnRefreshPerm')

    # Extras tab
    $panelExtras         = $Window.FindName('PanelExtras')
    $tabExtras           = $Window.FindName('TabExtras')
    $tabSignatures       = $Window.FindName('TabSignatures')
    $tabPermissions      = $Window.FindName('TabPermissions')

    # Settings tab
    $lbThemeSelector     = $Window.FindName('LbThemeSelector')
    $lbLanguageSelector  = $Window.FindName('LbLanguageSelector')

    # Status bar
    $txtStatus           = $Window.FindName('TxtStatus')

    # -- State --
    $script:currentSig        = $null
    $script:selectedSigName   = $null    # sig name clicked for assign/preview
    $script:selectedAccountKey= $null    # RegistryPath of selected account card
    $script:tabVisitCounts = @{ 0 = 0; 1 = 0 }   # index 0=Signatures, 1=Permissions
    $script:dlgResult      = $false
    $script:windowRef      = $Window
    $script:AppSettings    = Load-Settings

    # Script-scope refs for controls used inside script:-scoped functions
    $script:sigTxtSignatureInfo  = $txtSignatureInfo
    $script:sigPanelInboxList    = $panelInboxList
    $script:sigPanelCopyTargets  = $panelCopyTargets
    $script:sigTxtSelectedSig    = $txtSelectedSig

    # -- Helpers --

    function script:Set-Status([string]$msg) {
        $script:txtStatusRef.Text = $msg
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

    # Load a signature into the preview pane
    function script:Load-SignaturePreview([string]$name) {
        $script:currentSig = $name
        $htmlPath = Get-SignatureHtmlPath -Name $name
        $status   = Get-SignatureStatus -Name $name

        $script:previewBrowserRef.Visibility = 'Visible'
        if (Test-Path $htmlPath) {
            try {
                $script:previewBrowserRef.Navigate((New-Object System.Uri($htmlPath)))
                $warn = if ($status.Warning) { " [!] $($status.Warning)" } else { '' }
                $script:sigTxtSignatureInfo.Text = "$name$warn"
            } catch {
                $script:sigTxtSignatureInfo.Text = "Preview failed: $_"
            }
        } else {
            $bgHex2 = $script:Themes[$script:AppSettings.Theme].BgBrush
            $fgHex2 = $script:Themes[$script:AppSettings.Theme].TextSecondaryBrush
            $script:previewBrowserRef.NavigateToString("<!DOCTYPE html><html style=`"background:$bgHex2`"><body style=`"background:$bgHex2;color:$fgHex2;font-family:Segoe UI;padding:16px`">No HTML file found.</body></html>")
            $script:sigTxtSignatureInfo.Text = "[!] Missing .htm file"
        }
    }

    # Build a map of sig name -> colour based on file content hash
    function script:Build-SigColourMap {
        $palette = @('#D96C6C','#4FA89B','#C49A2A','#7B68C8','#4A8CC1','#6B9E3F','#B85C8A')
        $hashToColour = @{}
        $nameToColour = @{}
        $idx = 0
        foreach ($name in (Get-Signatures)) {
            $path = Get-SignatureHtmlPath -Name $name
            if (Test-Path $path) {
                $hash = (Get-FileHash $path -Algorithm MD5).Hash
                if (-not $hashToColour.ContainsKey($hash)) {
                    $hashToColour[$hash] = $palette[$idx % $palette.Count]
                    $idx++
                }
                $nameToColour[$name] = $hashToColour[$hash]
            }
        }
        return $nameToColour
    }

    # Select an account card visually (highlight border) and pre-tick its checkbox
    function script:Select-AccountCard($card, [string]$regPath) {
        # Deselect previous
        if ($null -ne $script:selectedAccountKey) {
            foreach ($child in $script:sigPanelInboxList.Children) {
                if ($child.Tag -eq $script:selectedAccountKey) {
                    $child.BorderThickness = [System.Windows.Thickness]::new(1)
                    $child.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'BorderBrush')
                    break
                }
            }
        }
        $script:selectedAccountKey = $regPath
        $card.BorderThickness = [System.Windows.Thickness]::new(2)
        $card.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'AccentBrush')

        # Pre-tick the matching checkbox in the bottom bar
        foreach ($cb in $script:sigPanelCopyTargets.Children) {
            if ($cb -is [System.Windows.Controls.CheckBox]) {
                $cb.IsChecked = ($cb.Tag -eq $regPath)
            }
        }
    }

    # Refresh the inbox list and bottom-bar checkboxes
    $script:refreshInboxList = {
        $colourMap   = Build-SigColourMap
        $assignments = Get-SignatureAssignments

        $script:sigPanelInboxList.Children.Clear()
        $script:sigPanelCopyTargets.Children.Clear()

        if ($assignments.Count -eq 0) {
            $lbl = New-Object System.Windows.Controls.TextBlock
            $lbl.Text = 'No Outlook accounts found.'
            $lbl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextSecondaryBrush')
            $lbl.FontSize = 12
            $lbl.Margin = [System.Windows.Thickness]::new(0, 8, 0, 0)
            $script:sigPanelInboxList.Children.Add($lbl) | Out-Null
            Set-Status 'No Outlook accounts found'
            return
        }

        foreach ($a in $assignments) {
            $displayName = if ($a.SmtpAddress) { $a.SmtpAddress } else { $a.AccountName }
            $regPath     = $a.RegistryPath

            # ── Account card ──
            $card = New-Object System.Windows.Controls.Border
            $card.CornerRadius    = [System.Windows.CornerRadius]::new(8)
            $card.BorderThickness = [System.Windows.Thickness]::new(1)
            $card.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty,   'SurfaceBrush')
            $card.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty,  'BorderBrush')
            $card.Padding = [System.Windows.Thickness]::new(10, 8, 10, 8)
            $card.Margin  = [System.Windows.Thickness]::new(0, 0, 0, 6)
            $card.Cursor  = [System.Windows.Input.Cursors]::Hand
            $card.Tag     = $regPath

            $sp = New-Object System.Windows.Controls.StackPanel

            # Account name
            $lblName = New-Object System.Windows.Controls.TextBlock
            $lblName.Text       = $displayName
            $lblName.FontSize   = 12
            $lblName.FontWeight = 'SemiBold'
            $lblName.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextPrimaryBrush')
            $lblName.TextTrimming = 'CharacterEllipsis'
            $lblName.ToolTip = "$($a.AccountName)`n$regPath"
            $sp.Children.Add($lblName) | Out-Null

            # Build New-mail sig label
            $lblNew = New-Object System.Windows.Controls.TextBlock
            $lblNew.FontSize = 11
            $lblNew.Margin   = [System.Windows.Thickness]::new(0, 2, 0, 0)
            $lblNew.Cursor   = [System.Windows.Input.Cursors]::Hand
            if ([string]::IsNullOrEmpty($a.NewSignature)) {
                $lblNew.Text = 'New: (None)'
                $lblNew.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextSecondaryBrush')
            } else {
                $cNew = if ($colourMap.ContainsKey($a.NewSignature)) { $colourMap[$a.NewSignature] } else { '#9898A8' }
                $lblNew.Text       = "New: $($a.NewSignature)"
                $lblNew.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($cNew)
                $lblNew.Tag        = $a.NewSignature
            }

            # Build Reply sig label
            $lblReply = New-Object System.Windows.Controls.TextBlock
            $lblReply.FontSize = 11
            $lblReply.Margin   = [System.Windows.Thickness]::new(0, 2, 0, 0)
            $lblReply.Cursor   = [System.Windows.Input.Cursors]::Hand
            if ([string]::IsNullOrEmpty($a.ReplySignature)) {
                $lblReply.Text = 'Reply: (None)'
                $lblReply.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextSecondaryBrush')
            } else {
                $cRep = if ($colourMap.ContainsKey($a.ReplySignature)) { $colourMap[$a.ReplySignature] } else { '#9898A8' }
                $lblReply.Text       = "Reply: $($a.ReplySignature)"
                $lblReply.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($cRep)
                $lblReply.Tag        = $a.ReplySignature
            }
            $sp.Children.Add($lblNew)   | Out-Null
            $sp.Children.Add($lblReply) | Out-Null

            $card.Child = $sp
            $script:sigPanelInboxList.Children.Add($card) | Out-Null

            # Card click → select card, preview New sig
            $cardRef    = $card
            $regPathRef = $regPath
            $newSigRef  = $a.NewSignature
            $card.add_MouseLeftButtonUp(({
                param($s2, $e2)
                Select-AccountCard $cardRef $regPathRef
                if (-not [string]::IsNullOrEmpty($newSigRef)) {
                    $script:selectedSigName = $newSigRef
                    $script:sigTxtSelectedSig.Text = $newSigRef
                    Load-SignaturePreview $newSigRef
                }
            }).GetNewClosure())

            # Sig-name label click → preview that specific sig + set as selected
            foreach ($child in @($lblNew, $lblReply)) {
                if (-not [string]::IsNullOrEmpty($child.Tag)) {
                    $sigRef   = $child.Tag
                    $cardRef2 = $card
                    $regRef2  = $regPath
                    $child.add_MouseLeftButtonUp(({
                        param($s3, $e3)
                        $e3.Handled = $true
                        Select-AccountCard $cardRef2 $regRef2
                        $script:selectedSigName = $sigRef
                        $script:sigTxtSelectedSig.Text = $sigRef
                        Load-SignaturePreview $sigRef
                    }).GetNewClosure())
                }
            }

            # ── Bottom-bar checkbox ──
            $cb = New-Object System.Windows.Controls.CheckBox
            $cb.Content  = $displayName
            $cb.Tag      = $regPath
            $cb.FontSize = 11
            $cb.Margin   = [System.Windows.Thickness]::new(0, 0, 12, 0)
            $cb.VerticalAlignment = 'Center'
            $script:sigPanelCopyTargets.Children.Add($cb) | Out-Null
        }

        Set-Status "Loaded $($assignments.Count) mailbox(es)"
    }

    # Reusable: refresh mailbox list (Permissions tab)
    $script:permAllAccounts = @()
    $refreshPermMailboxes = {
        $script:permAllAccounts = Get-SignedInAccounts
        $script:permLbMailboxes.ItemsSource = $null
        $script:permLbMailboxes.ItemsSource = $script:permAllAccounts
    }

    # Shared state for permissions tab
    $script:permSelectedFolder  = $null   # PSCustomObject with EntryID, StoreID, Name
    $script:permAdTimer         = $null   # DispatcherTimer for debounced AD search

    # -- WebBrowser: suppress native IE scrollbar via DOM injection on every load --
    # (ClipToBounds cannot clip HwndHost children; CSS overflow:hidden is the reliable fix)
    $previewBrowser.add_LoadCompleted({
        try {
            $doc = $script:previewBrowserRef.Document
            if ($null -ne $doc) {
                $body = $doc.Body
                if ($null -ne $body) {
                    $existing = $body.Style
                    if ($existing -notmatch 'overflow') {
                        $body.Style = "$existing; overflow: hidden !important;"
                    }
                }
            }
        } catch { }
    })

    # -- Extras tab unlock (visit Signatures AND Permissions each 3 times) --

    $mainTabs = $Window.FindName('MainTabs')

    $script:mainTabsRef  = $mainTabs
    $script:tabExtrasRef = $tabExtras
    $script:txtStatusRef = $txtStatus
    $script:statusBarRef = $Window.FindName('StatusBar')

    # Show status bar only on the Signatures tab (index 0)
    $mainTabs.add_SelectionChanged({
        param($s, $e)
        if ($e.OriginalSource -ne $script:mainTabsRef) { return }
        if ($script:mainTabsRef.SelectedIndex -eq 0) {
            $script:statusBarRef.Visibility = 'Visible'
            $script:txtStatusRef.Text = "Loaded $($script:sigPanelInboxList.Children.Count) mailbox(es)"
        } else {
            $script:statusBarRef.Visibility = 'Collapsed'
        }
    })

    $btnKillAll = $Window.FindName('BtnKillAll')
    $btnKillAll.Add_Click({
        $killPath = Join-Path $ScriptRoot "extras\KillAll.ps1"
        if (Test-Path $killPath) {
            Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$killPath`"" -WindowStyle Hidden
        }
        $script:txtStatusRef.Text = 'All troll processes killed'
    })

    $btnLockExtras = $Window.FindName('BtnLockExtras')
    $btnLockExtras.Add_Click({
        # Hide extras tab and switch back to Signatures
        $script:tabExtrasRef.Visibility = 'Collapsed'
        $script:mainTabsRef.SelectedIndex = 0
        # Remove Crimson from the theme selector
        for ($i = $script:lbThemeSelectorRef.Items.Count - 1; $i -ge 0; $i--) {
            if ($script:lbThemeSelectorRef.Items[$i].Tag -eq 'Crimson') {
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
        $script:txtStatusRef.Text = 'Extras locked'
    })

    # ── Unlock dialog (called from About section rapid-click) ────────────────
    function Show-UnlockDialog {
        if ($script:tabExtrasRef.Visibility -eq 'Visible') { return }

        $dlgXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Restricted Area" Width="420" SizeToContent="Height"
        Background="#2B2B2B" Foreground="#E8E8E8"
        FontFamily="Segoe UI" FontSize="13"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize">
  <StackPanel Margin="28,24,28,24">
    <Border Background="#3A1A1A" CornerRadius="8" BorderBrush="#C0392B" BorderThickness="1" Padding="16,12" Margin="0,0,0,20">
      <StackPanel>
        <TextBlock Text="&#x26A0;  RESTRICTED AREA" FontSize="15" FontWeight="Bold"
                   Foreground="#E74C3C" HorizontalAlignment="Center" Margin="0,0,0,8"/>
        <TextBlock TextWrapping="Wrap" HorizontalAlignment="Center" FontSize="11"
                   Foreground="#C0785A" TextAlignment="Center"
                   Text="You are attempting to access features outside the intended scope of this application. Proceed only if you know what you are doing."/>
      </StackPanel>
    </Border>
    <TextBlock Text="Passphrase" Foreground="#999999" FontSize="11" Margin="0,0,0,6"/>
    <PasswordBox Name="PwdInput"
                 Background="#3A3A3A" Foreground="#E8E8E8" BorderBrush="#555555" BorderThickness="1"
                 Padding="8,6" FontSize="13"/>
    <Grid Margin="0,16,0,0">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="8"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>
      <Button Name="BtnCancel" Grid.Column="0" Content="Cancel"
              Background="#3A3A3A" Foreground="#E8E8E8" BorderBrush="#555555" BorderThickness="1"
              Padding="0,8" FontSize="12"/>
      <Button Name="BtnUnlock" Grid.Column="2" Content="&#x1F513;  Unlock"
              Background="#C0392B" Foreground="#FFFFFF" BorderThickness="0"
              Padding="0,8" FontSize="12" FontWeight="SemiBold"/>
    </Grid>
  </StackPanel>
</Window>
"@
        $dlgReader = New-Object System.Xml.XmlNodeReader([xml]$dlgXaml)
        $dlg       = [Windows.Markup.XamlReader]::Load($dlgReader)
        $pwdInput  = $dlg.FindName('PwdInput')
        $btnUnlock = $dlg.FindName('BtnUnlock')
        $btnCancel = $dlg.FindName('BtnCancel')

        $script:dlgResult = $false
        $btnUnlock.Add_Click({ $script:dlgResult = $true;  $dlg.Close() })
        $btnCancel.Add_Click({ $script:dlgResult = $false; $dlg.Close() })
        $pwdInput.Add_KeyDown({
            param($s, $e)
            if ($e.Key -eq 'Return') { $script:dlgResult = $true; $dlg.Close() }
        })

        $dlg.ShowDialog() | Out-Null
        $entered = $pwdInput.Password

        if ($script:dlgResult -and $entered -ne '' -and $entered -ne 'itrustlucas') {
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

        if ($script:dlgResult -and $entered -eq 'itrustlucas') {
            $script:tabExtrasRef.Visibility = 'Visible'
            $script:txtStatusRef.Text = 'Extras unlocked'
            # Unlock the Crimson theme and switch to it
            $alreadyAdded = $false
            for ($i = 0; $i -lt $script:lbThemeSelectorRef.Items.Count; $i++) {
                if ($script:lbThemeSelectorRef.Items[$i].Tag -eq 'Crimson') { $alreadyAdded = $true; break }
            }
            if (-not $alreadyAdded) {
                $crimsonItem = New-Object System.Windows.Controls.ListBoxItem
                $crimsonItem.Content = 'Crimson'
                $crimsonItem.Tag     = 'Crimson'
                $crimsonItem.Padding = [System.Windows.Thickness]::new(18, 8, 18, 8)
                $crimsonItem.Margin  = [System.Windows.Thickness]::new(0, 0, 6, 0)
                $script:lbThemeSelectorRef.Items.Add($crimsonItem) | Out-Null
            }
            Apply-Theme -Window $script:windowRef -ThemeName 'Crimson'
            $script:AppSettings.Theme = 'Crimson'
            Save-Settings -Settings $script:AppSettings
            for ($i = 0; $i -lt $script:lbThemeSelectorRef.Items.Count; $i++) {
                if ($script:lbThemeSelectorRef.Items[$i].Tag -eq 'Crimson') {
                    $script:lbThemeSelectorRef.SelectedIndex = $i; break
                }
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

    # -- Initial load --

    $previewBrowser.Visibility = 'Collapsed'

    try { & $script:refreshInboxList } catch { Set-Status "Error loading mailboxes: $_" }
    # Note: refreshPermMailboxes is called after script-scope perm refs are stored (below)

    # -- Settings: apply saved theme and pre-select controls --

    # Restore unlocked state if Crimson was previously unlocked
    if ($script:AppSettings.Theme -eq 'Crimson') {
        $script:tabExtrasRef.Visibility = 'Visible'
        $crimsonItem = New-Object System.Windows.Controls.ListBoxItem
        $crimsonItem.Content = 'Crimson'
        $crimsonItem.Tag     = 'Crimson'
        $crimsonItem.Padding = [System.Windows.Thickness]::new(18, 8, 18, 8)
        $crimsonItem.Margin  = [System.Windows.Thickness]::new(0, 0, 6, 0)
        $lbThemeSelector.Items.Add($crimsonItem) | Out-Null
    }

    Apply-Theme -Window $Window -ThemeName $script:AppSettings.Theme

    # Pre-select saved theme
    for ($i = 0; $i -lt $lbThemeSelector.Items.Count; $i++) {
        if ($lbThemeSelector.Items[$i].Tag -eq $script:AppSettings.Theme) {
            $lbThemeSelector.SelectedIndex = $i; break
        }
    }

    # Pre-select saved language
    for ($i = 0; $i -lt $lbLanguageSelector.Items.Count; $i++) {
        if ($lbLanguageSelector.Items[$i].Tag -eq $script:AppSettings.Language) {
            $lbLanguageSelector.SelectedIndex = $i; break
        }
    }

    $script:lbThemeSelectorRef    = $lbThemeSelector
    $script:lbLanguageSelectorRef = $lbLanguageSelector

    $lbThemeSelector.Add_SelectionChanged({
        param($s, $e)
        if ($e.OriginalSource -ne $script:lbThemeSelectorRef) { return }
        $sel = $script:lbThemeSelectorRef.SelectedItem
        if ($null -eq $sel) { return }
        $themeName = $sel.Tag
        Apply-Theme -Window $script:windowRef -ThemeName $themeName
        $script:AppSettings.Theme = $themeName
        Save-Settings -Settings $script:AppSettings
    })

    $lbLanguageSelector.Add_SelectionChanged({
        param($s, $e)
        if ($e.OriginalSource -ne $script:lbLanguageSelectorRef) { return }
        $sel = $script:lbLanguageSelectorRef.SelectedItem
        if ($null -eq $sel) { return }
        $script:AppSettings.Language = $sel.Tag
        Save-Settings -Settings $script:AppSettings
    })

    # -- New --

    $btnNew.Add_Click({
        $name = Show-InputBox 'Enter a name for the new signature:' 'New Signature'
        if ([string]::IsNullOrWhiteSpace($name)) { return }
        try {
            New-Signature -Name $name
            & $script:refreshInboxList
            Set-Status "Created '$name'"
        } catch { Show-Error "$_"; Set-Status "Create failed" }
    })

    # -- Rename --

    $btnRename.Add_Click({
        $old = $script:selectedSigName
        if ([string]::IsNullOrWhiteSpace($old)) { Show-Error 'Click a signature name in the left panel first.'; return }
        $new = Show-InputBox "Rename '$old' to:" 'Rename Signature' $old
        if ([string]::IsNullOrWhiteSpace($new) -or $new -eq $old) { return }
        try {
            Rename-Signature -OldName $old -NewName $new
            $script:selectedSigName = $new
            $script:sigTxtSelectedSig.Text = $new
            & $script:refreshInboxList
            Set-Status "Renamed '$old' -> '$new'"
        } catch { Show-Error "$_"; Set-Status "Rename failed" }
    })

    # -- Delete --

    $btnDelete.Add_Click({
        $name = $script:selectedSigName
        if ([string]::IsNullOrWhiteSpace($name)) { Show-Error 'Click a signature name in the left panel first.'; return }
        if (-not (Confirm-Action "Delete signature '$name'? This cannot be undone." 'Delete Signature')) { return }
        try {
            Remove-Signature -Name $name
            $script:currentSig      = $null
            $script:selectedSigName = $null
            $script:sigTxtSelectedSig.Text = '(none selected)'
            $script:sigTxtSignatureInfo.Text = 'Select a mailbox or signature to preview.'
            $script:previewBrowserRef.Visibility = 'Collapsed'
            & $script:refreshInboxList
            Set-Status "Deleted '$name'"
        } catch { Show-Error "$_"; Set-Status "Delete failed" }
    })

    # -- Assign to checked mailboxes --

    $btnAssignSig.Add_Click({
        if ([string]::IsNullOrWhiteSpace($script:selectedSigName)) {
            Show-Error 'Click a signature name in the left panel first.'
            return
        }
        $ticked = @($script:sigPanelCopyTargets.Children |
                    Where-Object { $_ -is [System.Windows.Controls.CheckBox] -and $_.IsChecked })
        if ($ticked.Count -eq 0) { Show-Error 'Tick at least one mailbox below.'; return }
        try {
            foreach ($cb in $ticked) {
                Set-SignatureAssignment -RegistryPath $cb.Tag `
                    -NewSignature $script:selectedSigName `
                    -ReplySignature $script:selectedSigName
            }
            & $script:refreshInboxList
            $outlookRunning = Test-OutlookRunning
            Set-Status "Assigned '$($script:selectedSigName)' to $($ticked.Count) mailbox(es)"
            if ($outlookRunning) {
                [System.Windows.MessageBox]::Show(
                    "Assignment saved.`n`nOutlook is currently running - restart it for changes to take effect.",
                    'Assigned', [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Information) | Out-Null
            }
        } catch { Show-Error "$_"; Set-Status "Assign failed" }
    })

    # ── Permissions tab ─────────────────────────────────────────────────────────

    # Populate permission level ComboBox from module
    foreach ($lvl in (Get-PermissionLevels)) { $cbPermLevel.Items.Add($lvl) | Out-Null }
    $cbPermLevel.SelectedIndex = 1  # default: "Can view"

    # Store ALL perm controls at script scope so WPF event handler closures can reach them
    $script:permTxtStatus    = $txtPermStatus
    $script:permDgPerms      = $dgCurrentPerms
    $script:permPanelRight   = $panelPermRight
    $script:permLbMailboxes  = $lbMailboxes
    $script:permLbFolders    = $lbFolders
    $script:permTxtFolderHint  = $txtFolderHint
    $script:permTxtFoldersHint = $txtFoldersHint
    $script:permCbPermLevel       = $cbPermLevel
    $script:permTbMailboxSearch   = $tbMailboxSearch

    # Initial mailbox load now that all script-scope refs are set
    try { & $refreshPermMailboxes } catch { Set-Status "Mailboxes unavailable: $_" }

    function script:Set-PermStatus {
        param([string]$Msg, [string]$Colour = '')
        $script:permTxtStatus.Text = $Msg
        if ($Colour) {
            $script:permTxtStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($Colour)
        } else {
            $script:permTxtStatus.SetResourceReference(
                [System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextSecondaryBrush')
        }
    }

    function script:Refresh-PermGrid {
        if ($null -eq $script:permSelectedFolder) { return }
        try {
            $rows = Get-FolderPermissions -EntryID $script:permSelectedFolder.EntryID `
                                          -StoreID  $script:permSelectedFolder.StoreID
            $script:permDgPerms.ItemsSource = $null
            $script:permDgPerms.ItemsSource = $rows
        } catch {
            Set-PermStatus "Could not read permissions: $_" '#E74C3C'
        }
    }

    # ── Mailbox search filter ────────────────────────────────────────────────
    $tbMailboxSearch.Add_TextChanged({
        $q = $script:permTbMailboxSearch.Text.Trim()
        if ([string]::IsNullOrEmpty($q)) {
            $script:permLbMailboxes.ItemsSource = $script:permAllAccounts
        } else {
            $script:permLbMailboxes.ItemsSource = @($script:permAllAccounts | Where-Object {
                $_.Name -like "*$q*" -or $_.SmtpAddress -like "*$q*"
            })
        }
    })

    # Script-scope helpers called from event handlers
    function script:Set-PermRightEnabled {
        param([bool]$Enabled)
        $script:permPanelRight.IsEnabled = $Enabled
        $script:permPanelRight.Opacity   = if ($Enabled) { 1.0 } else { 0.45 }
    }

    function script:Get-FolderIcon {
        param([string]$Name, [int]$Depth)
        if ($Depth -eq 0) { return '[M]' }
        switch -Wildcard ($Name) {
            'Inbox'    { return '[I]' }
            'Sent*'    { return '[S]' }
            'Deleted*' { return '[X]' }
            'Drafts'   { return '[D]' }
            'Calendar' { return '[C]' }
            'Contacts' { return '[P]' }
            'Tasks'    { return '[T]' }
            'Junk*'    { return '[J]' }
            'Archive*' { return '[A]' }
            'Outbox'   { return '[O]' }
            default    { if ($Depth -eq 1) { return '[F]' } else { return '  [F]' } }
        }
    }

    # Initial state: right panel dimmed
    Set-PermRightEnabled $false

    # ── Mailbox selection → populate folder list ─────────────────────────────
    $lbMailboxes.Add_SelectionChanged({
        $sel = $script:permLbMailboxes.SelectedItem
        if ($null -eq $sel) { return }
        $script:permLbFolders.ItemsSource   = $null
        $script:permDgPerms.ItemsSource     = $null
        $script:permSelectedFolder          = $null
        $script:permTxtFolderHint.Visibility  = 'Visible'
        $script:permTxtFoldersHint.Text       = 'Loading folders...'
        Set-PermRightEnabled $false
        Set-PermStatus ''
        try {
            $raw = Get-MailboxFolders -SmtpAddress $sel.SmtpAddress
            if ($raw.Count -eq 0) {
                $script:permTxtFoldersHint.Text = 'No folders found'
            } else {
                $script:permTxtFoldersHint.Text = ''
                $items = $raw | ForEach-Object {
                    [PSCustomObject]@{
                        Name       = $_.Name
                        Icon       = Get-FolderIcon -Name $_.Name -Depth $_.Depth
                        FolderPath = $_.FolderPath
                        EntryID    = $_.EntryID
                        StoreID    = $_.StoreID
                        Depth      = $_.Depth
                        IndentPad  = [System.Windows.Thickness]::new([int]($_.Depth * 12), 1, 4, 1)
                    }
                }
                $script:permLbFolders.ItemsSource = $items
            }
            Set-Status "Loaded folders for $($sel.Name)"
        } catch {
            $script:permTxtFoldersHint.Text = 'Could not load folders'
            Set-PermStatus "Could not load folders: $_" '#E74C3C'
            Set-Status "Error loading folders"
        }
    })

    # ── Folder selection → show permissions ──────────────────────────────────
    $lbFolders.Add_SelectionChanged({
        $sel = $script:permLbFolders.SelectedItem
        if ($null -eq $sel) { return }
        $script:permSelectedFolder          = $sel
        $script:permTxtFolderHint.Visibility  = 'Collapsed'
        $script:permDgPerms.ItemsSource     = $null
        $script:permRemovePending           = $null
        Set-PermStatus ''
        Set-PermRightEnabled $true
        try {
            Refresh-PermGrid
            Set-Status "Permissions for: $($sel.Name)"
        } catch {
            Set-PermStatus "Error reading permissions: $_" '#E74C3C'
        }
    })

    # ── AD search: debounced via DispatcherTimer ─────────────────────────────
    # The timer tick and the search logic are broken out as named functions
    # so the parser never sees a scriptblock literal inside an event handler.
    $script:permAdLastQuery = ''
    $script:permAdTimer     = $null

    # Store controls needed by AD search functions at script scope
    $script:permLbAdResults = $lbAdResults
    $script:permPopAddUser  = $popAddUser
    $script:permTxtAddUser  = $txtAddUser

    function script:Invoke-AdSearchTick {
        if ($null -ne $script:permAdTimer) { $script:permAdTimer.Stop() }
        $script:permAdTimer = $null
        try {
            $hits = Search-ADUsers -Query $script:permAdLastQuery
            $script:permLbAdResults.ItemsSource = $null
            if ($hits.Count -gt 0) {
                $script:permLbAdResults.ItemsSource = $hits
                $script:permPopAddUser.IsOpen = $true
            } else {
                $script:permPopAddUser.IsOpen = $false
                Set-PermStatus "No AD matches - type an email address directly." ''
            }
        } catch {
            $script:permPopAddUser.IsOpen = $false
        }
    }

    function script:Start-AdSearchDebounce {
        param([string]$Query)
        if ($null -ne $script:permAdTimer) {
            $script:permAdTimer.Stop()
            $script:permAdTimer = $null
        }
        if ($Query.Length -lt 2) {
            $script:permPopAddUser.IsOpen = $false
            return
        }
        $script:permAdLastQuery = $Query
        $t = New-Object System.Windows.Threading.DispatcherTimer
        $t.Interval = [System.TimeSpan]::FromMilliseconds(350)
        $t.Add_Tick({ Invoke-AdSearchTick })
        $script:permAdTimer = $t
        $t.Start()
    }

    $txtAddUser.Add_TextChanged({
        Start-AdSearchDebounce -Query $script:permTxtAddUser.Text.Trim()
    })

    # ── AD result selected → fill search box, close popup ───────────────────
    $lbAdResults.Add_SelectionChanged({
        $hit = $script:permLbAdResults.SelectedItem
        if ($null -eq $hit) { return }
        $value = if ($hit.Mail) { $hit.Mail } else { $hit.DisplayName }
        $script:permTxtAddUser.Text      = $value
        $script:permTxtAddUser.CaretIndex = $value.Length
        $script:permPopAddUser.IsOpen    = $false
        Set-PermStatus ''
    })

    # ── Save permission ──────────────────────────────────────────────────────
    $btnSavePerm.Add_Click({
        if ($null -eq $script:permSelectedFolder) {
            Set-PermStatus 'Please select a folder first.' '#E74C3C'; return
        }
        $user  = $script:permTxtAddUser.Text.Trim()
        $level = $script:permCbPermLevel.SelectedItem
        if ([string]::IsNullOrWhiteSpace($user)) {
            Set-PermStatus 'Please enter a name or email address.' '#E74C3C'; return
        }
        if ($null -eq $level) {
            Set-PermStatus 'Please select an access level.' '#E74C3C'; return
        }
        try {
            Set-FolderPermission -EntryID $script:permSelectedFolder.EntryID `
                                 -StoreID $script:permSelectedFolder.StoreID `
                                 -User    $user `
                                 -Level   $level
            Refresh-PermGrid
            Set-PermStatus "Saved: $user - $level" '#27AE60'
            Set-Status "Permission saved for $user"
        } catch {
            Set-PermStatus "Could not save: $_" '#E74C3C'
            Set-Status "Permission save failed"
        }
    })

    # ── Remove permission (two-click confirm via flag) ───────────────────────
    $script:permRemovePending = $null   # user name pending confirmation
    $btnRemovePerm.Add_Click({
        if ($null -eq $script:permSelectedFolder) {
            $script:permRemovePending = $null
            Set-PermStatus 'Please select a folder first.' '#E74C3C'; return
        }
        $row = $script:permDgPerms.SelectedItem
        if ($null -eq $row) {
            $script:permRemovePending = $null
            Set-PermStatus 'Select a person in the list above to remove them.' '#E74C3C'; return
        }
        if ($row.User -eq 'Default' -or $row.User -eq 'Anonymous') {
            $script:permRemovePending = $null
            Set-PermStatus "The '$($row.User)' entry cannot be removed - change its level instead." '#E74C3C'
            return
        }
        if ($script:permRemovePending -ne $row.User) {
            # First click — ask for confirmation
            $script:permRemovePending = $row.User
            Set-PermStatus "Remove access for $($row.User)? Click Remove again to confirm." ''
            return
        }
        # Second click — confirmed
        $script:permRemovePending = $null
        try {
            Remove-FolderPermission -EntryID $script:permSelectedFolder.EntryID `
                                    -StoreID $script:permSelectedFolder.StoreID `
                                    -User    $row.User
            Refresh-PermGrid
            Set-PermStatus "Removed: $($row.User)" '#27AE60'
            Set-Status "Removed $($row.User)"
        } catch {
            Set-PermStatus "Could not remove: $_" '#E74C3C'
            Set-Status "Remove failed"
        }
    })
    # Clear pending state when the selected row changes
    $dgCurrentPerms.Add_SelectionChanged({ $script:permRemovePending = $null; Set-PermStatus '' })

    # ── Refresh mailboxes button ─────────────────────────────────────────────
    $btnRefreshPerm.Add_Click({
        try {
            & $refreshPermMailboxes
            $script:permLbFolders.ItemsSource      = $null
            $script:permDgPerms.ItemsSource        = $null
            $script:permSelectedFolder             = $null
            $script:permTxtFolderHint.Visibility   = 'Visible'
            $script:permTxtFoldersHint.Text        = 'Select a mailbox to see its folders'
            Set-PermRightEnabled $false
            Set-PermStatus ''
            Set-Status "Mailboxes refreshed"
        } catch {
            Set-Status "Error refreshing mailboxes: $_"
        }
    })

    # ─── Extras Tab ───────────────────────────────────────────────────────────
    # Each entry: display name, script filename, description, whether it is a
    # one-shot (runs once and exits) or a loop (needs its own process).
    $extrasDefinitions = @(
        [PSCustomObject]@{ Name='Mouse Confusion';       File='MouseConfusion.ps1';   Desc='Nudges the mouse cursor a few pixels every 5 s.';                          Loop=$true;  Hidden=$true  }
        [PSCustomObject]@{ Name='Caps Lock Troll';       File='CapsLockTroll.ps1';    Desc='Toggles Caps Lock every 8 s.';                                             Loop=$true;  Hidden=$true  }
        [PSCustomObject]@{ Name='Infinite Popup';        File='InfinitePopup.ps1';    Desc='Pops up a message box every 10 s.';                                        Loop=$true;  Hidden=$true  }
        [PSCustomObject]@{ Name='Reverse Keyboard';      File='ReverseKeyboard.ps1';  Desc='Sends a LEFT arrow keystroke every 10 s while typing.';                    Loop=$true;  Hidden=$true  }
        [PSCustomObject]@{ Name='Annoying Beeps';        File='AnnoyingBeeps.ps1';    Desc='Beeps every 500 ms.';                                                       Loop=$true;  Hidden=$true  }
        [PSCustomObject]@{ Name='Invert Colors';         File='InvertColors.ps1';     Desc='Toggles High Contrast mode (run again to restore).';                       Loop=$false; Hidden=$true  }
        [PSCustomObject]@{ Name='Fake Progress Bar';     File='FakeProgressBar.ps1';  Desc='Shows a fake "Installing Update..." progress bar in a console window.';    Loop=$false; Hidden=$false }
    )

    # Track running background processes: script name -> Process object
    function Stop-ExtraProcess([string]$file) {
        if ($script:extrasProcs.ContainsKey($file)) {
            $proc = $script:extrasProcs[$file]
            if ($proc -and -not $proc.HasExited) {
                try { $proc.Kill() } catch {}
            }
            $script:extrasProcs.Remove($file)
        }
    }

    foreach ($def in $extrasDefinitions) {
        $scriptPath = Join-Path $ScriptRoot "extras\$($def.File)"
        $defCopy    = $def   # capture for closure

        # Outer card border
        $card = New-Object System.Windows.Controls.Border
        $card.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'SurfaceBrush')
        $card.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'BorderBrush')
        $card.CornerRadius     = New-Object System.Windows.CornerRadius 6
        $card.BorderThickness  = New-Object System.Windows.Thickness 1
        $card.Padding          = New-Object System.Windows.Thickness 12,10,12,10
        $card.Margin           = New-Object System.Windows.Thickness 0,0,0,8

        $row = New-Object System.Windows.Controls.Grid
        $col0 = New-Object System.Windows.Controls.ColumnDefinition; $col0.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        $col1 = New-Object System.Windows.Controls.ColumnDefinition; $col1.Width = [System.Windows.GridLength]::Auto
        $row.ColumnDefinitions.Add($col0)
        $row.ColumnDefinitions.Add($col1)

        # Text block (name + description)
        $txtPanel = New-Object System.Windows.Controls.StackPanel
        $txtPanel.Orientation = [System.Windows.Controls.Orientation]::Vertical

        $lblName = New-Object System.Windows.Controls.TextBlock
        $lblName.Text       = $defCopy.Name
        $lblName.FontWeight = [System.Windows.FontWeights]::SemiBold
        $lblName.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextPrimaryBrush')
        $lblName.FontSize   = 13

        $lblDesc = New-Object System.Windows.Controls.TextBlock
        $lblDesc.Text       = $defCopy.Desc
        $lblDesc.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextSecondaryBrush')
        $lblDesc.FontSize   = 11
        $lblDesc.TextWrapping = [System.Windows.TextWrapping]::Wrap

        $txtPanel.Children.Add($lblName) | Out-Null
        $txtPanel.Children.Add($lblDesc) | Out-Null

        [System.Windows.Controls.Grid]::SetColumn($txtPanel, 0)
        $row.Children.Add($txtPanel) | Out-Null

        # Toggle button
        $btn = New-Object System.Windows.Controls.Button
        $btn.Content  = 'Run'
        $btn.Width    = 72
        $btn.Height   = 30
        $btn.Margin   = New-Object System.Windows.Thickness 12,0,0,0
        $btn.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        [System.Windows.Controls.Grid]::SetColumn($btn, 1)
        $row.Children.Add($btn) | Out-Null

        $card.Child = $row
        $panelExtras.Children.Add($card) | Out-Null

        # Button click handler — use a single-element array as a mutable process ref
        # that survives across Run/Stop clicks within the same closure.
        $scriptPathCopy = $scriptPath
        $isLoop         = $defCopy.Loop
        $isHidden       = $defCopy.Hidden
        $nameCopy       = $defCopy.Name
        $btnCopy        = $btn   # capture for closure — $btn is reused each iteration
        $procBox        = [ref] $null  # mutable slot for the running process

        $btn.Add_Click(({
            if ($btnCopy.Content -eq 'Run') {
                if (-not (Test-Path $scriptPathCopy)) {
                    [System.Windows.MessageBox]::Show("Script not found: $scriptPathCopy", 'Error',
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Error) | Out-Null
                    return
                }
                if ($isLoop) {
                    $psi = New-Object System.Diagnostics.ProcessStartInfo
                    $psi.FileName               = 'powershell.exe'
                    $psi.Arguments              = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPathCopy`""
                    $psi.WindowStyle            = [System.Diagnostics.ProcessWindowStyle]::Hidden
                    $psi.CreateNoWindow         = $true
                    $procBox.Value              = [System.Diagnostics.Process]::Start($psi)
                    $btnCopy.Content    = 'Stop'
                    $btnCopy.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#C0392B')
                    $txtStatus.Text = "$nameCopy started"
                } else {
                    $winStyle = if ($isHidden) { 'Hidden' } else { 'Normal' }
                    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPathCopy`"" -WindowStyle $winStyle
                    $txtStatus.Text = "$nameCopy executed"
                }
            } else {
                $p = $procBox.Value
                if ($p -and -not $p.HasExited) { try { $p.Kill() } catch {} }
                $procBox.Value = $null
                $btnCopy.Content    = 'Run'
                $btnCopy.SetResourceReference([System.Windows.Controls.Control]::BackgroundProperty, 'AccentBrush')
                $txtStatus.Text = "$nameCopy stopped"
            }
        }).GetNewClosure())
    }
    # ─── End Extras Tab ───────────────────────────────────────────────────────
}
