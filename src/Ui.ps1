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
    $lbSignatures        = $Window.FindName('LbSignatures')
    $script:lbSignaturesRef = $lbSignatures
    $btnNew              = $Window.FindName('BtnNewSignature')
    $btnDuplicate        = $Window.FindName('BtnDuplicateSignature')
    $btnRename           = $Window.FindName('BtnRenameSignature')
    $btnDelete           = $Window.FindName('BtnDeleteSignature')
    $btnRefreshSigs      = $Window.FindName('BtnRefreshSignatures')
    $txtEditorTitle      = $Window.FindName('TxtEditorTitle')
    $btnToggleEditor     = $Window.FindName('BtnToggleEditor')
    $btnSave             = $Window.FindName('BtnSaveSignature')
    $previewBrowser      = $Window.FindName('PreviewBrowser')
    $script:previewBrowserRef = $previewBrowser
    $panelPreview        = $Window.FindName('PanelPreview')
    $txtHtmlEditor       = $Window.FindName('TxtHtmlEditor')
    $txtSignatureInfo    = $Window.FindName('TxtSignatureInfo')
    $btnExport           = $Window.FindName('BtnExportSignature')
    $btnImport           = $Window.FindName('BtnImportSignature')
    $btnRefreshAssign    = $Window.FindName('BtnRefreshAssignments')
    $panelRows           = $Window.FindName('PanelAssignmentRows')
    $btnApply            = $Window.FindName('BtnApplyAssignments')
    $txtOutlookWarning   = $Window.FindName('TxtOutlookWarning')

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
    $script:editorMode   = 'preview'   # 'preview' | 'html'
    $script:currentSig   = $null
    $script:assignRows   = @{}         # RegistryPath -> @{ NewCb; ReplyCb }
    $script:tabVisitCounts = @{ 0 = 0; 1 = 0 }   # index 0=Signatures, 1=Permissions
    $script:dlgResult      = $false
    $script:windowRef      = $Window
    $script:AppSettings    = Load-Settings

    # -- Helpers --

    function Set-Status([string]$msg) {
        $txtStatus.Text = $msg
    }

    function Show-InputBox([string]$prompt, [string]$title, [string]$default = '') {
        Add-Type -AssemblyName Microsoft.VisualBasic
        return [Microsoft.VisualBasic.Interaction]::InputBox($prompt, $title, $default)
    }

    function Confirm-Action([string]$msg, [string]$title = 'Confirm') {
        return [System.Windows.MessageBox]::Show($msg, $title,
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning) -eq [System.Windows.MessageBoxResult]::Yes
    }

    function Show-Error([string]$msg) {
        [System.Windows.MessageBox]::Show($msg, 'Error',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error) | Out-Null
    }

    # Reusable: refresh signature list, preserving selection by name
    $refreshSignatures = {
        param([string]$SelectName = $null)
        $lbSignatures.Items.Clear()
        $statuses = Get-SignatureStatusList
        foreach ($s in $statuses) {
            $label = if ($s.Warning) { "[!] $($s.Name)" } else { $s.Name }
            $item = New-Object System.Windows.Controls.ListBoxItem
            $item.Content = $label
            $item.Tag     = $s.Name
            $item.ToolTip = if ($s.Warning) { "Missing files: $($s.Warning)" } else { $null }
            $lbSignatures.Items.Add($item) | Out-Null
        }
        if ($SelectName) {
            for ($i = 0; $i -lt $lbSignatures.Items.Count; $i++) {
                if ($lbSignatures.Items[$i].Tag -eq $SelectName) {
                    $lbSignatures.SelectedIndex = $i
                    break
                }
            }
        }
        Set-Status "Loaded $($lbSignatures.Items.Count) signatures"
    }

    # Get clean sig name from selected ListBoxItem
    function Get-SelectedSigName {
        $sel = $lbSignatures.SelectedItem
        if ($null -eq $sel) { return $null }
        return $sel.Tag
    }

    # Load signature into the editor/preview area
    function Load-Signature([string]$name) {
        $script:currentSig = $name
        $txtEditorTitle.Text = $name
        $btnSave.IsEnabled = $false

        # Always switch back to preview mode when loading a new sig
        if ($script:editorMode -eq 'html') {
            $script:editorMode = 'preview'
            $btnToggleEditor.Content = 'Edit HTML'
            $panelPreview.Visibility = 'Visible'
            $txtHtmlEditor.Visibility = 'Collapsed'
        }

        $htmlPath = Get-SignatureHtmlPath -Name $name
        $status   = Get-SignatureStatus -Name $name

        $previewBrowser.Visibility = 'Visible'
        if (Test-Path $htmlPath) {
            try {
                $previewBrowser.Navigate((New-Object System.Uri($htmlPath)))
                $warn = if ($status.Warning) { " [!] $($status.Warning)" } else { '' }
                $txtSignatureInfo.Text = "$name$warn"
            } catch {
                $txtSignatureInfo.Text = "Preview failed: $_"
            }
        } else {
            $bgHex2 = $script:Themes[$script:AppSettings.Theme].BgBrush
            $fgHex2 = $script:Themes[$script:AppSettings.Theme].TextSecondaryBrush
            $previewBrowser.NavigateToString("<!DOCTYPE html><html style=`"background:$bgHex2`"><body style=`"background:$bgHex2;color:$fgHex2;font-family:Segoe UI;padding:16px`">No HTML file found.</body></html>")
            $txtSignatureInfo.Text = "[!] Missing .htm file"
        }
    }

    # Build one assignment row in the right panel
    function Add-AssignmentRow($assignment, [string[]]$sigNames) {
        $row = New-Object System.Windows.Controls.Grid
        $col0 = New-Object System.Windows.Controls.ColumnDefinition
        $col0.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        $col1 = New-Object System.Windows.Controls.ColumnDefinition
        $col1.Width = [System.Windows.GridLength]::new(6)
        $col2 = New-Object System.Windows.Controls.ColumnDefinition
        $col2.Width = [System.Windows.GridLength]::new(90)
        $col3 = New-Object System.Windows.Controls.ColumnDefinition
        $col3.Width = [System.Windows.GridLength]::new(6)
        $col4 = New-Object System.Windows.Controls.ColumnDefinition
        $col4.Width = [System.Windows.GridLength]::new(90)
        $row.ColumnDefinitions.Add($col0)
        $row.ColumnDefinitions.Add($col1)
        $row.ColumnDefinitions.Add($col2)
        $row.ColumnDefinitions.Add($col3)
        $row.ColumnDefinitions.Add($col4)
        $row.Margin = [System.Windows.Thickness]::new(0, 0, 0, 6)

        # Account label
        $lblAccount = New-Object System.Windows.Controls.TextBlock
        $lblAccount.Text = if ($assignment.SmtpAddress) { $assignment.SmtpAddress } else { $assignment.AccountName }
        $lblAccount.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextPrimaryBrush')
        $lblAccount.FontSize = 11
        $lblAccount.VerticalAlignment = 'Center'
        $lblAccount.TextTrimming = 'CharacterEllipsis'
        $lblAccount.ToolTip = "$($assignment.AccountName)`n$($assignment.RegistryPath)"
        [System.Windows.Controls.Grid]::SetColumn($lblAccount, 0)

        # Helper: build a small sig picker ListBox
        function New-SigPicker([string]$currentValue) {
            $lb = New-Object System.Windows.Controls.ListBox
            $lb.SetResourceReference([System.Windows.Controls.ListBox]::BackgroundProperty, 'BgBrush')
            $lb.SetResourceReference([System.Windows.Controls.ListBox]::BorderBrushProperty, 'BorderBrush')
            $lb.BorderThickness = [System.Windows.Thickness]::new(1)
            $lb.Padding = [System.Windows.Thickness]::new(2)
            $lb.Tag = $assignment.RegistryPath

            $noneItem = New-Object System.Windows.Controls.ListBoxItem
            $noneItem.Content = '(None)'
            $noneItem.Tag = ''
            $noneItem.Padding = [System.Windows.Thickness]::new(4, 3, 4, 3)
            $noneItem.FontSize = 11
            $lb.Items.Add($noneItem) | Out-Null

            foreach ($s in $sigNames) {
                $item = New-Object System.Windows.Controls.ListBoxItem
                $item.Content = $s
                $item.Tag = $s
                $item.Padding = [System.Windows.Thickness]::new(4, 3, 4, 3)
                $item.FontSize = 11
                if ($s -eq $currentValue -and -not (Test-Path (Get-SignatureHtmlPath -Name $s))) {
                    $item.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#E06F6F')
                    $item.ToolTip = 'Signature file not found on disk'
                }
                $lb.Items.Add($item) | Out-Null
            }

            for ($i = 0; $i -lt $lb.Items.Count; $i++) {
                if ($lb.Items[$i].Tag -eq $currentValue) { $lb.SelectedIndex = $i; break }
            }
            if ($lb.SelectedIndex -lt 0) { $lb.SelectedIndex = 0 }
            return $lb
        }

        $cbNew   = New-SigPicker $assignment.NewSignature
        $cbReply = New-SigPicker $assignment.ReplySignature
        [System.Windows.Controls.Grid]::SetColumn($cbNew,   2)
        [System.Windows.Controls.Grid]::SetColumn($cbReply, 4)

        $row.Children.Add($lblAccount) | Out-Null
        $row.Children.Add($cbNew)      | Out-Null
        $row.Children.Add($cbReply)    | Out-Null
        $panelRows.Children.Add($row)  | Out-Null

        $script:assignRows[$assignment.RegistryPath] = @{ NewCb = $cbNew; ReplyCb = $cbReply }
    }

    # Refresh the assignment panel
    $refreshAssignments = {
        $panelRows.Children.Clear()
        $script:assignRows = @{}
        $sigNames = Get-Signatures
        $assignments = Get-SignatureAssignments
        if ($assignments.Count -eq 0) {
            $lbl = New-Object System.Windows.Controls.TextBlock
            $lbl.Text = 'No Outlook accounts found.'
            $lbl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextSecondaryBrush')
            $lbl.FontSize = 12
            $lbl.Margin = [System.Windows.Thickness]::new(4, 8, 0, 0)
            $panelRows.Children.Add($lbl) | Out-Null
        } else {
            foreach ($a in $assignments) { Add-AssignmentRow $a $sigNames }
        }
        $txtOutlookWarning.Visibility = if (Test-OutlookRunning) { 'Visible' } else { 'Collapsed' }
    }

    # Reusable: refresh mailbox list (Permissions tab)
    $script:permAllAccounts = @()
    $refreshPermMailboxes = {
        $script:permAllAccounts = Get-SignedInAccounts
        $lbMailboxes.ItemsSource = $null
        $lbMailboxes.ItemsSource = $script:permAllAccounts
    }

    # Shared state for permissions tab
    $script:permSelectedFolder  = $null   # PSCustomObject with EntryID, StoreID, Name
    $script:permAdTimer         = $null   # DispatcherTimer for debounced AD search

    # -- WebBrowser: suppress native IE scrollbar via DOM injection on every load --
    # (ClipToBounds cannot clip HwndHost children; CSS overflow:hidden is the reliable fix)
    $previewBrowser.add_LoadCompleted({
        try {
            $doc = $previewBrowser.Document
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
            $script:txtStatusRef.Text = "Loaded $($script:lbSignaturesRef.Items.Count) signatures"
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

    try { & $refreshSignatures } catch { Set-Status "Error loading signatures: $_" }
    try { & $refreshAssignments } catch { Set-Status "Error loading assignments: $_" }
    try { & $refreshPermMailboxes } catch { }  # silent - Outlook COM may not be available

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

    # -- Signature list selection --

    $lbSignatures.add_SelectionChanged({
        $name = Get-SelectedSigName
        if ($null -ne $name) { Load-Signature $name }
    })

    # -- New --

    $btnNew.Add_Click({
        $name = Show-InputBox 'Enter a name for the new signature:' 'New Signature'
        if ([string]::IsNullOrWhiteSpace($name)) { return }
        try {
            New-Signature -Name $name
            & $refreshSignatures -SelectName $name
            & $refreshAssignments
            Set-Status "Created '$name'"
        } catch { Show-Error "$_"; Set-Status "Create failed" }
    })

    # -- Duplicate --

    $btnDuplicate.Add_Click({
        $src = Get-SelectedSigName
        if ($null -eq $src) { Show-Error 'Select a signature first.'; return }
        $target = Show-InputBox "Enter a name for the copy of '$src':" 'Duplicate Signature' "$src - Copy"
        if ([string]::IsNullOrWhiteSpace($target)) { return }
        try {
            Copy-Signature -SourceName $src -TargetName $target
            & $refreshSignatures -SelectName $target
            & $refreshAssignments
            Set-Status "Duplicated '$src' -> '$target'"
        } catch { Show-Error "$_"; Set-Status "Duplicate failed" }
    })

    # -- Rename --

    $btnRename.Add_Click({
        $old = Get-SelectedSigName
        if ($null -eq $old) { Show-Error 'Select a signature first.'; return }
        $new = Show-InputBox "Rename '$old' to:" 'Rename Signature' $old
        if ([string]::IsNullOrWhiteSpace($new) -or $new -eq $old) { return }
        try {
            Rename-Signature -OldName $old -NewName $new
            & $refreshSignatures -SelectName $new
            & $refreshAssignments
            Set-Status "Renamed '$old' -> '$new'"
        } catch { Show-Error "$_"; Set-Status "Rename failed" }
    })

    # -- Delete --

    $btnDelete.Add_Click({
        $name = Get-SelectedSigName
        if ($null -eq $name) { Show-Error 'Select a signature first.'; return }
        if (-not (Confirm-Action "Delete signature '$name'? This cannot be undone." 'Delete Signature')) { return }
        try {
            Remove-Signature -Name $name
            $script:currentSig = $null
            $txtEditorTitle.Text = 'Select a signature'
            $txtSignatureInfo.Text = 'Select a signature to edit or preview.'
            $previewBrowser.Visibility = 'Collapsed'
            & $refreshSignatures
            & $refreshAssignments
            Set-Status "Deleted '$name'"
        } catch { Show-Error "$_"; Set-Status "Delete failed" }
    })

    # -- Refresh signatures --

    $btnRefreshSigs.Add_Click({
        $cur = Get-SelectedSigName
        & $refreshSignatures -SelectName $cur
    })

    # -- Toggle WYSIWYG / HTML editor --

    $btnToggleEditor.Add_Click({
        if ($null -eq $script:currentSig) { Show-Error 'Select a signature first.'; return }

        if ($script:editorMode -eq 'preview') {
            $htmlPath = Get-SignatureHtmlPath -Name $script:currentSig
            if (Test-Path $htmlPath) {
                $txtHtmlEditor.Text = [System.IO.File]::ReadAllText($htmlPath, [System.Text.Encoding]::UTF8)
            } else {
                $txtHtmlEditor.Text = ''
            }
            $panelPreview.Visibility  = 'Collapsed'
            $txtHtmlEditor.Visibility = 'Visible'
            $btnToggleEditor.Content  = 'Preview'
            $btnSave.IsEnabled        = $true
            $script:editorMode        = 'html'
        } else {
            $panelPreview.Visibility  = 'Visible'
            $txtHtmlEditor.Visibility = 'Collapsed'
            $btnToggleEditor.Content  = 'Edit HTML'
            $btnSave.IsEnabled        = $false
            $script:editorMode        = 'preview'
            Load-Signature $script:currentSig
        }
    })

    # -- Save HTML --

    $btnSave.Add_Click({
        if ($null -eq $script:currentSig) { return }
        try {
            Save-SignatureHtml -Name $script:currentSig -HtmlContent $txtHtmlEditor.Text
            Set-Status "Saved '$($script:currentSig)'"
            $panelPreview.Visibility  = 'Visible'
            $txtHtmlEditor.Visibility = 'Collapsed'
            $btnToggleEditor.Content  = 'Edit HTML'
            $btnSave.IsEnabled        = $false
            $script:editorMode        = 'preview'
            Load-Signature $script:currentSig
        } catch { Show-Error "$_"; Set-Status "Save failed" }
    })

    # -- Export --

    $btnExport.Add_Click({
        $name = Get-SelectedSigName
        if ($null -eq $name) { Show-Error 'Select a signature first.'; return }
        Add-Type -AssemblyName System.Windows.Forms
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.FileName  = "$name.zip"
        $sfd.Filter    = 'Zip files (*.zip)|*.zip'
        $sfd.Title     = "Export '$name'"
        if ($sfd.ShowDialog() -eq 'OK') {
            try {
                Export-Signature -Name $name -Destination $sfd.FileName
                Set-Status "Exported '$name' to $($sfd.FileName)"
            } catch { Show-Error "$_"; Set-Status "Export failed" }
        }
    })

    # -- Import --

    $btnImport.Add_Click({
        Add-Type -AssemblyName System.Windows.Forms
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = 'Zip files (*.zip)|*.zip|All files (*.*)|*.*'
        $ofd.Title  = 'Import Signature'
        if ($ofd.ShowDialog() -eq 'OK') {
            try {
                Import-Signature -ZipPath $ofd.FileName
                & $refreshSignatures
                & $refreshAssignments
                Set-Status "Imported from $($ofd.FileName)"
            } catch { Show-Error "$_"; Set-Status "Import failed" }
        }
    })

    # -- Assignment: Refresh --

    $btnRefreshAssign.Add_Click({
        try { & $refreshAssignments; Set-Status "Refreshed assignments" }
        catch { Show-Error "$_"; Set-Status "Error refreshing assignments" }
    })

    # -- Assignment: Apply --

    $btnApply.Add_Click({
        if ($script:assignRows.Count -eq 0) { return }
        $outlookWasRunning = Test-OutlookRunning
        try {
            foreach ($regPath in $script:assignRows.Keys) {
                $row    = $script:assignRows[$regPath]
                $newSig = if ($row.NewCb.SelectedItem)   { $row.NewCb.SelectedItem.Tag }   else { '' }
                $repSig = if ($row.ReplyCb.SelectedItem) { $row.ReplyCb.SelectedItem.Tag } else { '' }
                Set-SignatureAssignment -RegistryPath $regPath -NewSignature $newSig -ReplySignature $repSig
            }
            Set-Status "Assignments applied to registry"
            if ($outlookWasRunning) {
                [System.Windows.MessageBox]::Show(
                    "Assignments written.`n`nOutlook is currently running - restart it for changes to take effect.",
                    'Applied', [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Information) | Out-Null
            } else {
                [System.Windows.MessageBox]::Show("Assignments applied successfully.", 'Applied',
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Information) | Out-Null
            }
        } catch { Show-Error "$_"; Set-Status "Apply failed" }
    })

    # ── Permissions tab ─────────────────────────────────────────────────────────

    # Populate permission level ComboBox from module
    foreach ($lvl in (Get-PermissionLevels)) { $cbPermLevel.Items.Add($lvl) | Out-Null }
    $cbPermLevel.SelectedIndex = 1  # default: "Can view"

    # Helper: set inline status with optional colour
    function Set-PermStatus {
        param([string]$Msg, [string]$Colour = '')
        $txtPermStatus.Text = $Msg
        if ($Colour) {
            $txtPermStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($Colour)
        } else {
            # Use theme secondary colour
            $txtPermStatus.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextSecondaryBrush')
        }
    }

    # Helper: reload the permissions DataGrid for the currently selected folder
    function Refresh-PermGrid {
        if ($null -eq $script:permSelectedFolder) { return }
        try {
            $rows = Get-FolderPermissions -EntryID $script:permSelectedFolder.EntryID `
                                          -StoreID  $script:permSelectedFolder.StoreID
            $dgCurrentPerms.ItemsSource = $null
            $dgCurrentPerms.ItemsSource = $rows
        } catch {
            Set-PermStatus "Could not read permissions: $_" '#E74C3C'
        }
    }

    # ── Mailbox search filter ────────────────────────────────────────────────
    $tbMailboxSearch.Add_TextChanged({
        $q = $tbMailboxSearch.Text.Trim()
        if ([string]::IsNullOrEmpty($q)) {
            $lbMailboxes.ItemsSource = $script:permAllAccounts
        } else {
            $lbMailboxes.ItemsSource = @($script:permAllAccounts | Where-Object {
                $_.Name -like "*$q*" -or $_.SmtpAddress -like "*$q*"
            })
        }
    })

    # ── Mailbox selection → populate folder list ─────────────────────────────
    $lbMailboxes.Add_SelectionChanged({
        $sel = $lbMailboxes.SelectedItem
        if ($null -eq $sel) { return }
        $lbFolders.ItemsSource   = $null
        $dgCurrentPerms.ItemsSource = $null
        $script:permSelectedFolder  = $null
        $txtFolderHint.Visibility   = 'Visible'
        Set-PermStatus ''
        try {
            $raw = Get-MailboxFolders -SmtpAddress $sel.SmtpAddress
            # Build display items with indentation prefix for depth
            $items = $raw | ForEach-Object {
                $indent = if ($_.Depth -gt 0) { ('    ' * ($_.Depth - 1)) + '  └─ ' } else { '' }
                [PSCustomObject]@{
                    Name        = $_.Name
                    Indent      = $indent
                    FolderPath  = $_.FolderPath
                    EntryID     = $_.EntryID
                    StoreID     = $_.StoreID
                    Depth       = $_.Depth
                }
            }
            $lbFolders.ItemsSource = $items
            Set-Status "Loaded folders for $($sel.Name)"
        } catch {
            Set-PermStatus "Could not load folders: $_" '#E74C3C'
            Set-Status "Error loading folders"
        }
    })

    # ── Folder selection → show permissions ──────────────────────────────────
    $lbFolders.Add_SelectionChanged({
        $sel = $lbFolders.SelectedItem
        if ($null -eq $sel) { return }
        $script:permSelectedFolder = $sel
        $txtFolderHint.Visibility  = 'Collapsed'
        $dgCurrentPerms.ItemsSource = $null
        Set-PermStatus ''
        try {
            Refresh-PermGrid
            Set-Status "Showing permissions for: $($sel.Name)"
        } catch {
            Set-PermStatus "Error reading permissions: $_" '#E74C3C'
        }
    })

    # ── AD search: debounced via DispatcherTimer ─────────────────────────────
    # Tick handler is defined as a named scriptblock so the parser doesn't
    # struggle with deeply-nested closures inside event handlers.
    $script:permAdTickHandler = $null

    $txtAddUser.Add_TextChanged({
        $q = $txtAddUser.Text.Trim()
        if ($null -ne $script:permAdTimer) {
            $script:permAdTimer.Stop()
            $script:permAdTimer = $null
        }
        if ($q.Length -lt 2) {
            $popAddUser.IsOpen = $false
            return
        }
        # Build a fresh tick handler that closes over the current query string
        $script:permAdTickHandler = {
            $script:permAdTimer.Stop()
            $script:permAdTimer = $null
            try {
                $hits = Search-ADUsers -Query $script:permAdLastQuery
                $lbAdResults.ItemsSource = $null
                if ($hits.Count -gt 0) {
                    $lbAdResults.ItemsSource = $hits
                    $popAddUser.IsOpen = $true
                } else {
                    $popAddUser.IsOpen = $false
                    Set-PermStatus "No AD matches — you can still type an email address directly." ''
                }
            } catch {
                $popAddUser.IsOpen = $false
            }
        }
        $script:permAdLastQuery = $q
        $timer = New-Object System.Windows.Threading.DispatcherTimer
        $timer.Interval = [System.TimeSpan]::FromMilliseconds(350)
        $timer.Add_Tick($script:permAdTickHandler)
        $script:permAdTimer = $timer
        $timer.Start()
    })

    # ── AD result selected → fill search box, close popup ───────────────────
    $lbAdResults.Add_SelectionChanged({
        $hit = $lbAdResults.SelectedItem
        if ($null -eq $hit) { return }
        # Prefer email, fall back to display name
        $value = if ($hit.Mail) { $hit.Mail } else { $hit.DisplayName }
        $txtAddUser.Text     = $value
        $txtAddUser.CaretIndex = $value.Length
        $popAddUser.IsOpen   = $false
        Set-PermStatus ''
    })

    # ── Save permission ──────────────────────────────────────────────────────
    $btnSavePerm.Add_Click({
        if ($null -eq $script:permSelectedFolder) {
            Set-PermStatus 'Please select a folder first.' '#E74C3C'; return
        }
        $user  = $txtAddUser.Text.Trim()
        $level = $cbPermLevel.SelectedItem
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
            Set-PermStatus "Saved: $user — $level" '#27AE60'
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
        $row = $dgCurrentPerms.SelectedItem
        if ($null -eq $row) {
            $script:permRemovePending = $null
            Set-PermStatus 'Select a person in the list above to remove them.' '#E74C3C'; return
        }
        if ($row.User -eq 'Default' -or $row.User -eq 'Anonymous') {
            $script:permRemovePending = $null
            Set-PermStatus "The '$($row.User)' entry cannot be removed — you can change its level instead." '#E74C3C'
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
            $lbFolders.ItemsSource      = $null
            $dgCurrentPerms.ItemsSource = $null
            $script:permSelectedFolder  = $null
            $txtFolderHint.Visibility   = 'Visible'
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
