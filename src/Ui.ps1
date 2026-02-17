$script:extrasProcs = @{}
# Theme definitions  -> Theme.ps1
# Translations       -> Language.ps1

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
    $script:previewBrowserRef     = $previewBrowser
    $script:previewEmptyHintRef   = $Window.FindName('TxtPreviewEmptyHint')
    $txtSignatureInfo    = $Window.FindName('TxtSignatureInfo')
    $panelInboxList      = $Window.FindName('PanelInboxList')
    $panelCopyTargets    = $Window.FindName('PanelCopyTargets')
    $txtSelectedSig      = $Window.FindName('TxtSelectedSig')
    # Permissions tab (Advanced)
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
    # Permissions tab (Easy — overview + wizard)
    $btnGenerateOverview = $Window.FindName('BtnGenerateOverview')
    $panelPermReport     = $Window.FindName('PanelPermReport')
    $panelPermReportCnt  = $Window.FindName('PanelPermReportContent')
    $lbWizMailboxes      = $Window.FindName('LbWizMailboxes')
    $txtWizUserSearch    = $Window.FindName('TxtWizUserSearch')
    $popWizAdResults     = $Window.FindName('PopWizAdResults')
    $lbWizAdResults      = $Window.FindName('LbWizAdResults')
    $panelWizSelectedUsers = $Window.FindName('PanelWizSelectedUsers')
    $panelWizFolders     = $Window.FindName('PanelWizFolders')
    $panelWizPermLevels  = $Window.FindName('PanelWizPermLevels')
    $panelWizSummary     = $Window.FindName('PanelWizSummary')
    $txtWizResult        = $Window.FindName('TxtWizResult')

    # Extras tab
    $panelExtras         = $Window.FindName('PanelExtras')
    $tabExtras           = $Window.FindName('TabExtras')
    $tabSignatures       = $Window.FindName('TabSignatures')
    $tabPermissions      = $Window.FindName('TabPermissions')

    # Settings tab
    $lbThemeSelector     = $Window.FindName('LbThemeSelector')
    $cbLanguageSelector  = $Window.FindName('CbLanguageSelector')

    # Status bar
    $txtStatus           = $Window.FindName('TxtStatus')

    # -- State --
    $script:currentSig        = $null
    $script:selectedSigName   = $null    # sig name currently selected for preview/assign
    $script:selectedAccountKey= $null    # RegistryPath of the highlighted account card
    $script:tabVisitCounts = @{ 0 = 0; 1 = 0 }   # index 0=Signatures, 1=Permissions
    $script:dlgResult      = $false
    $script:windowRef      = $Window
    $script:AppSettings    = Load-Settings

    # Script-scope refs for controls used inside script:-scoped functions
    # NOTE: TxtSelectedSig is looked up via $script:windowRef.FindName at call-time to
    # avoid a $null reference when closures fire before the window fully renders.
    $script:sigTxtSignatureInfo  = $txtSignatureInfo
    $script:sigPanelInboxList    = $panelInboxList
    $script:sigPanelCopyTargets  = $panelCopyTargets
    $script:sigBtnEdit           = $Window.FindName('BtnEditSignature')
    $script:sigBtnReload         = $Window.FindName('BtnReloadPreview')

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

    # Helper: safely update TxtSelectedSig text (guards against $null reference)
    # Also annotates bottom-bar pills with "New", "Reply", or "New + Reply" badges
    # and auto-highlights pills whose mailbox already has the selected signature.
    function script:Set-SelectedSigLabel([string]$text) {
        $ctrl = $script:windowRef.FindName('TxtSelectedSig')
        if ($null -ne $ctrl) { $ctrl.Text = $text }
        $editBtn = $script:windowRef.FindName('BtnEditSignature')
        $reloadBtn = $script:windowRef.FindName('BtnReloadPreview')
        $isEmpty = [string]::IsNullOrEmpty($text) -or $text -eq '(none selected)'
        if ($null -ne $editBtn)   { $editBtn.Visibility   = if ($isEmpty) { 'Collapsed' } else { 'Visible' } }
        if ($null -ne $reloadBtn) { $reloadBtn.Visibility = if ($isEmpty) { 'Collapsed' } else { 'Visible' } }

        # Colour-match TxtSelectedSig to the signature's colour from the left panel
        if ($null -ne $ctrl) {
            if (-not $isEmpty -and $null -ne $script:colourMap -and $script:colourMap.ContainsKey($text)) {
                $ctrl.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($script:colourMap[$text])
            } else {
                $ctrl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'AccentBrush')
            }
        }

        # Highlight matching signature rows in the left panel
        if ($null -ne $script:sigPanelInboxList) {
            foreach ($child in $script:sigPanelInboxList.Children) {
                if (-not ($child -is [System.Windows.Controls.Border])) { continue }
                $tag = $child.Tag
                if ($null -eq $tag -or $tag -isnot [string]) { continue }
                # Mailbox cards have registry paths as Tag (contain backslash); sig rows have plain names
                if ($tag.Contains('\')) { continue }
                if (-not $isEmpty -and $tag -eq $text) {
                    $child.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'AccentSubtleBrush')
                } else {
                    $child.Background = $null
                }
            }
        }

        # Annotate and highlight pills based on assignment state
        if ($null -ne $script:sigPanelCopyTargets) {
            foreach ($pill in $script:sigPanelCopyTargets.Children) {
                if (-not ($pill -is [System.Windows.Controls.Border])) { continue }
                $sp = $pill.Child
                if (-not ($sp -is [System.Windows.Controls.StackPanel] -and $sp.Children.Count -ge 2)) { continue }
                $pillTxt = $sp.Children[0]
                $badge   = $sp.Children[1]
                if ($isEmpty -or $null -eq $pill.Tag) {
                    $badge.Visibility = 'Collapsed'
                    $badge.Text = ''
                    $pill.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'SurfaceHoverBrush')
                    $pillTxt.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextPrimaryBrush')
                    continue
                }
                $isNew   = ($pill.Tag.NewSignature   -eq $text)
                $isReply = ($pill.Tag.ReplySignature -eq $text)
                $isAssigned = $isNew -or $isReply
                if ($isNew -and $isReply) {
                    $badge.Text       = 'New + Reply'
                    $badge.Visibility = 'Visible'
                } elseif ($isNew) {
                    $badge.Text       = 'New'
                    $badge.Visibility = 'Visible'
                } elseif ($isReply) {
                    $badge.Text       = 'Reply'
                    $badge.Visibility = 'Visible'
                } else {
                    $badge.Visibility = 'Collapsed'
                    $badge.Text = ''
                }
                # Highlight pill if this mailbox has the selected signature assigned
                if ($isAssigned) {
                    $pill.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'AccentBrush')
                    $pillTxt.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextOnAccentBrush')
                } else {
                    $pill.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'SurfaceHoverBrush')
                    $pillTxt.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextPrimaryBrush')
                }
            }
        }
    }

    # Load a signature into the preview pane
    function script:Load-SignaturePreview([string]$name) {
        $script:currentSig = $name
        $htmlPath = Get-SignatureHtmlPath -Name $name
        $status   = Get-SignatureStatus -Name $name

        $script:previewBrowserRef.Visibility = 'Visible'
        if ($null -ne $script:previewEmptyHintRef) { $script:previewEmptyHintRef.Visibility = 'Collapsed' }
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

    # Select an account card visually (accent border)
    function script:Select-AccountCard($card, [string]$regPath) {
        # Deselect previous card
        if ($null -ne $script:selectedAccountKey) {
            foreach ($child in $script:sigPanelInboxList.Children) {
                if ($child -is [System.Windows.Controls.Border] -and $child.Tag -eq $script:selectedAccountKey) {
                    $child.BorderThickness = [System.Windows.Thickness]::new(1)
                    $child.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'BorderBrush')
                    break
                }
            }
        }
        $script:selectedAccountKey = $regPath
        $card.BorderThickness = [System.Windows.Thickness]::new(2)
        $card.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'AccentBrush')
    }

    # Toggle-assign: clicking a pill assigns or unassigns the currently selected signature
    function script:Toggle-AssignPill($pill) {
        # Read from the UI label as authoritative source (closures may bind $script: to a different scope)
        $ctrl = $script:windowRef.FindName('TxtSelectedSig')
        $sigName = if ($null -ne $ctrl) { $ctrl.Text } else { $script:selectedSigName }
        $noneLabels = @('(none selected)', '(select a signature above)', '(Signatur oben auswählen)')
        if ([string]::IsNullOrWhiteSpace($sigName) -or $sigName -in $noneLabels) {
            Show-Error 'Click a signature name in the left panel first.'
            return
        }
        $assignHtmlPath = Get-SignatureHtmlPath -Name $sigName
        if (-not (Test-Path $assignHtmlPath)) {
            Show-Error "Signature '$sigName' no longer exists."
            Set-SelectedSigLabel '(none selected)'
            $script:selectedSigName = $null
            return
        }
        $regPath = $pill.Tag.RegPath
        $isAssigned = ($pill.Tag.NewSignature -eq $sigName -or $pill.Tag.ReplySignature -eq $sigName)
        try {
            if ($isAssigned) {
                Set-SignatureAssignment -RegistryPath $regPath -NewSignature '' -ReplySignature ''
                $accountName = $pill.Tag.AccountName
                Set-Status (Get-Str 'StatusUnassigned' $sigName $accountName)
            } else {
                Set-SignatureAssignment -RegistryPath $regPath -NewSignature $sigName -ReplySignature $sigName
                $accountName = $pill.Tag.AccountName
                Set-Status (Get-Str 'StatusAssigned' $sigName $accountName)
            }
            # Refresh UI and re-select the same signature
            & $script:refreshInboxList
            $script:selectedSigName = $sigName
            Set-SelectedSigLabel $sigName
            Load-SignaturePreview $sigName
        } catch { Show-Error "$_"; Set-Status (Get-Str 'StatusAssignFailed') }
    }

    # Refresh the inbox list and bottom-bar pills
    $script:refreshInboxList = {
        $script:colourMap = Build-SigColourMap
        $assignments = Get-SignatureAssignments
        # Capture TxtSelectedSig control for use in closures (avoids $script:windowRef scope issues)
        $selSigCtrl = $script:windowRef.FindName('TxtSelectedSig')
        if ($null -eq $selSigCtrl) { $selSigCtrl = $Window.FindName('TxtSelectedSig') }

        $script:sigPanelInboxList.Children.Clear()
        $script:sigPanelCopyTargets.Children.Clear()
        $script:selectedAccountKey  = $null   # prevent ghost-highlight after refresh

        if ($assignments.Count -eq 0) {
            $lbl = New-Object System.Windows.Controls.TextBlock
            $lbl.Text = 'No Outlook accounts found.'
            $lbl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextSecondaryBrush')
            $lbl.FontSize = 12
            $lbl.Margin = [System.Windows.Thickness]::new(0, 8, 0, 0)
            $script:sigPanelInboxList.Children.Add($lbl) | Out-Null
            Set-Status (Get-Str 'StatusNoAccounts')
            return
        }

        # ── "Local Signatures" header (all sigs not yet assigned) ──
        $allSigNames = Get-Signatures
        $assignedSigNames = @()
        foreach ($a in $assignments) {
            if (-not [string]::IsNullOrEmpty($a.NewSignature))   { $assignedSigNames += $a.NewSignature }
            if (-not [string]::IsNullOrEmpty($a.ReplySignature)) { $assignedSigNames += $a.ReplySignature }
        }
        $assignedSigNames = @($assignedSigNames | Sort-Object -Unique)
        $localOnlySigs = @($allSigNames | Where-Object { $assignedSigNames -notcontains $_ })

        $isFirstAccount = $true
        foreach ($a in $assignments) {
            $accountDisplay = $a.AccountName
            $smtpDisplay    = $a.SmtpAddress
            $regPath        = $a.RegistryPath

            # ── Mailbox row (flat, not a card) ──
            $mbBorder = New-Object System.Windows.Controls.Border
            $mbBorder.CornerRadius    = [System.Windows.CornerRadius]::new(6)
            $mbBorder.BorderThickness = [System.Windows.Thickness]::new(1)
            $mbBorder.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty,  'SurfaceBrush')
            $mbBorder.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'BorderBrush')
            $mbBorder.Padding = [System.Windows.Thickness]::new(10, 6, 10, 6)
            # Extra top margin between mailbox groups (not before the first one)
            $topMargin = if ($isFirstAccount) { 0 } else { 8 }
            $mbBorder.Margin  = [System.Windows.Thickness]::new(0, $topMargin, 0, 2)
            $mbBorder.Cursor  = [System.Windows.Input.Cursors]::Hand
            $mbBorder.Tag     = $regPath
            $isFirstAccount   = $false

            $mbSp = New-Object System.Windows.Controls.StackPanel

            $lblMbName = New-Object System.Windows.Controls.TextBlock
            $lblMbName.Text         = $accountDisplay
            $lblMbName.FontSize     = 12
            $lblMbName.FontWeight   = 'SemiBold'
            $lblMbName.TextTrimming = 'CharacterEllipsis'
            $lblMbName.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextPrimaryBrush')
            if (-not [string]::IsNullOrEmpty($smtpDisplay)) { $lblMbName.ToolTip = $smtpDisplay }
            $mbSp.Children.Add($lblMbName) | Out-Null

            if (-not [string]::IsNullOrEmpty($smtpDisplay) -and $smtpDisplay -ne $accountDisplay) {
                $lblSmtp = New-Object System.Windows.Controls.TextBlock
                $lblSmtp.Text         = $smtpDisplay
                $lblSmtp.FontSize     = 10
                $lblSmtp.TextTrimming = 'CharacterEllipsis'
                $lblSmtp.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextSecondaryBrush')
                $mbSp.Children.Add($lblSmtp) | Out-Null
            }

            $mbBorder.Child = $mbSp
            $script:sigPanelInboxList.Children.Add($mbBorder) | Out-Null

            # Collect unique sig names for this account
            $sigNames = @()
            if (-not [string]::IsNullOrEmpty($a.NewSignature))   { $sigNames += $a.NewSignature }
            if (-not [string]::IsNullOrEmpty($a.ReplySignature) -and $a.ReplySignature -ne $a.NewSignature) {
                $sigNames += $a.ReplySignature
            }

            $sigLabels = @()
            if ($sigNames.Count -eq 0) {
                $noSig = New-Object System.Windows.Controls.TextBlock
                $noSig.Text     = '  (no signature assigned)'
                $noSig.FontSize = 10
                $noSig.FontStyle = 'Italic'
                $noSig.Margin   = [System.Windows.Thickness]::new(14, 2, 0, 4)
                $noSig.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextSecondaryBrush')
                $script:sigPanelInboxList.Children.Add($noSig) | Out-Null
            } else {
                foreach ($sn in $sigNames) {
                    $colour = if ($colourMap.ContainsKey($sn)) { $colourMap[$sn] } else { '#9898A8' }
                    $sigRow = New-Object System.Windows.Controls.Border
                    $sigRow.CornerRadius    = [System.Windows.CornerRadius]::new(4)
                    $sigRow.BorderThickness = [System.Windows.Thickness]::new(0)
                    $sigRow.Padding         = [System.Windows.Thickness]::new(6, 3, 6, 3)
                    $sigRow.Margin          = [System.Windows.Thickness]::new(14, 1, 4, 1)
                    $sigRow.Cursor          = [System.Windows.Input.Cursors]::Hand
                    $sigRow.Tag             = $sn

                    # Horizontal layout: bullet + name
                    $sigSp = New-Object System.Windows.Controls.StackPanel
                    $sigSp.Orientation = 'Horizontal'

                    $sigBullet = New-Object System.Windows.Controls.TextBlock
                    $sigBullet.Text     = [char]0x25CF   # ●
                    $sigBullet.FontSize = 7
                    $sigBullet.Margin   = [System.Windows.Thickness]::new(0, 3, 6, 0)
                    $sigBullet.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($colour)
                    $sigSp.Children.Add($sigBullet) | Out-Null

                    $sigTb = New-Object System.Windows.Controls.TextBlock
                    $sigTb.Text         = $sn
                    $sigTb.FontSize     = 11
                    $sigTb.TextTrimming = 'CharacterEllipsis'
                    $sigTb.Foreground   = [System.Windows.Media.BrushConverter]::new().ConvertFrom($colour)
                    $sigSp.Children.Add($sigTb) | Out-Null
                    $sigRow.Child = $sigSp

                    $sigRow.add_MouseEnter(({
                        param($sr,$er)
                        $sr.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'SurfaceHoverBrush')
                    }).GetNewClosure())
                    $sigRow.add_MouseLeave(({
                        param($sr,$er)
                        $selName = if ($null -ne $selSigCtrl) { $selSigCtrl.Text } else { '' }
                        if ($sr.Tag -eq $selName) {
                            $sr.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'AccentSubtleBrush')
                        } else {
                            $sr.Background = $null
                        }
                    }).GetNewClosure())

                    $script:sigPanelInboxList.Children.Add($sigRow) | Out-Null
                    $sigLabels += $sigRow
                }
            }

            # Mailbox row click → select card, preview first sig
            $mbBorderRef = $mbBorder
            $regPathRef  = $regPath
            $firstSig    = if ($sigNames.Count -gt 0) { $sigNames[0] } else { '' }
            $mbBorder.add_MouseLeftButtonUp(({
                param($s2, $e2)
                Select-AccountCard $mbBorderRef $regPathRef
                if (-not [string]::IsNullOrEmpty($firstSig)) {
                    $script:selectedSigName = $firstSig
                    Set-SelectedSigLabel $firstSig
                    Load-SignaturePreview $firstSig
                } else {
                    $script:selectedSigName = $null
                    Set-SelectedSigLabel '(none selected)'
                }
            }).GetNewClosure())

            # Sig row click → preview that sig
            foreach ($sigRow in $sigLabels) {
                $sigRef      = $sigRow.Tag
                $mbBorderRef2 = $mbBorder
                $regRef2      = $regPath
                $sigRow.add_MouseLeftButtonUp(({
                    param($s3, $e3)
                    $e3.Handled = $true
                    Select-AccountCard $mbBorderRef2 $regRef2
                    $script:selectedSigName = $sigRef
                    Set-SelectedSigLabel $sigRef
                    Load-SignaturePreview $sigRef
                }).GetNewClosure())
            }

            # ── Bottom-bar pill ──
            $pill = New-Object System.Windows.Controls.Border
            $pill.CornerRadius    = [System.Windows.CornerRadius]::new(12)
            $pill.BorderThickness = [System.Windows.Thickness]::new(1)
            $pill.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty,  'SurfaceHoverBrush')
            $pill.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'BorderBrush')
            $pill.Padding         = [System.Windows.Thickness]::new(10, 4, 10, 4)
            $pill.Margin          = [System.Windows.Thickness]::new(0, 0, 6, 0)
            $pill.Cursor          = [System.Windows.Input.Cursors]::Hand
            # Store regPath + assignment info so Set-SelectedSigLabel can annotate the pill
            $pill.Tag             = [PSCustomObject]@{
                RegPath        = $regPath
                AccountName    = $accountDisplay
                NewSignature   = $a.NewSignature
                ReplySignature = $a.ReplySignature
            }

            $pillSp = New-Object System.Windows.Controls.StackPanel
            $pillSp.Orientation = 'Vertical'

            $pillTxt = New-Object System.Windows.Controls.TextBlock
            $pillTxt.Text     = $accountDisplay
            $pillTxt.FontSize = 11
            $pillTxt.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextPrimaryBrush')
            $pillSp.Children.Add($pillTxt) | Out-Null

            # Badge label — shows "New", "Reply", or "New + Reply" when that sig is active here
            $pillBadge = New-Object System.Windows.Controls.TextBlock
            $pillBadge.FontSize   = 9
            $pillBadge.Visibility = 'Collapsed'
            $pillBadge.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextOnAccentBrush')
            $pillSp.Children.Add($pillBadge) | Out-Null

            $pill.Child = $pillSp

            $pillRef = $pill
            $pill.add_MouseLeftButtonUp(({
                param($sp2, $ep2)
                Toggle-AssignPill $pillRef
            }).GetNewClosure())

            $script:sigPanelCopyTargets.Children.Add($pill) | Out-Null
        }

        # ── Local Signatures section (unassigned sigs) ──
        if ($localOnlySigs.Count -gt 0) {
            $localHdr = New-Object System.Windows.Controls.TextBlock
            $localHdr.Text       = 'LOCAL SIGNATURES'
            $localHdr.FontSize   = 10
            $localHdr.FontWeight = 'SemiBold'
            $localHdr.Margin     = [System.Windows.Thickness]::new(2, 14, 0, 4)
            $localHdr.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextSecondaryBrush')
            $script:sigPanelInboxList.Children.Add($localHdr) | Out-Null

            foreach ($sn in $localOnlySigs) {
                $colour = if ($colourMap.ContainsKey($sn)) { $colourMap[$sn] } else { '#9898A8' }
                $localRow = New-Object System.Windows.Controls.Border
                $localRow.CornerRadius    = [System.Windows.CornerRadius]::new(4)
                $localRow.BorderThickness = [System.Windows.Thickness]::new(0)
                $localRow.Padding         = [System.Windows.Thickness]::new(6, 3, 6, 3)
                $localRow.Margin          = [System.Windows.Thickness]::new(0, 1, 4, 1)
                $localRow.Cursor          = [System.Windows.Input.Cursors]::Hand
                $localRow.Tag             = $sn

                $localSp = New-Object System.Windows.Controls.StackPanel
                $localSp.Orientation = 'Horizontal'

                $localBullet = New-Object System.Windows.Controls.TextBlock
                $localBullet.Text     = [char]0x25CF   # ●
                $localBullet.FontSize = 7
                $localBullet.Margin   = [System.Windows.Thickness]::new(0, 3, 6, 0)
                $localBullet.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($colour)
                $localSp.Children.Add($localBullet) | Out-Null

                $localTb = New-Object System.Windows.Controls.TextBlock
                $localTb.Text         = $sn
                $localTb.FontSize     = 11
                $localTb.TextTrimming = 'CharacterEllipsis'
                $localTb.Foreground   = [System.Windows.Media.BrushConverter]::new().ConvertFrom($colour)
                $localSp.Children.Add($localTb) | Out-Null
                $localRow.Child = $localSp

                $localRow.add_MouseEnter(({
                    param($sr,$er)
                    $sr.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'SurfaceHoverBrush')
                }).GetNewClosure())
                $localRow.add_MouseLeave(({
                    param($sr,$er)
                    $selName = if ($null -ne $selSigCtrl) { $selSigCtrl.Text } else { '' }
                    if ($sr.Tag -eq $selName) {
                        $sr.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'AccentSubtleBrush')
                    } else {
                        $sr.Background = $null
                    }
                }).GetNewClosure())

                $snRef = $sn
                $localRow.add_MouseLeftButtonUp(({
                    param($s4, $e4)
                    $script:selectedAccountKey = $null
                    $script:selectedSigName = $snRef
                    Set-SelectedSigLabel $snRef
                    Load-SignaturePreview $snRef
                }).GetNewClosure())

                $script:sigPanelInboxList.Children.Add($localRow) | Out-Null
            }
        }

        Set-Status (Get-Str 'StatusLoadedMailboxes' $assignments.Count)
    }

    # Reusable: refresh mailbox list (Permissions tab)
    $script:permAllAccounts = @()
    $refreshPermMailboxes = {
        $script:permAllAccounts = @(Get-SignedInAccounts)
        $script:permLbMailboxes.ItemsSource = $null
        $script:permLbMailboxes.ItemsSource = $script:permAllAccounts
        # Also update the wizard mailbox list
        if ($null -ne $script:wizLbMailboxes) {
            $script:wizLbMailboxes.ItemsSource = $null
            $script:wizLbMailboxes.ItemsSource = $script:permAllAccounts
        }
        $isEmpty = $script:permAllAccounts.Count -eq 0
        if ($null -ne $script:permTxtMailboxEmptyHint) {
            $script:permTxtMailboxEmptyHint.Visibility = if ($isEmpty) { 'Visible' } else { 'Collapsed' }
        }
        if ($isEmpty) {
            Set-Status (Get-Str 'StatusNoAccountsOl')
        } else {
            Set-Status (Get-Str 'StatusLoadedMailboxes' $script:permAllAccounts.Count)
        }
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

    # Status bar is always visible — no toggle needed
    $mainTabs.add_SelectionChanged({
        param($s, $e)
        if ($e.OriginalSource -ne $script:mainTabsRef) { return }
        # Reset Konami progress if Extras tab is selected while still locked
        # (guards against partial-sequence keyboard navigation reaching the tab)
        $selected = $script:mainTabsRef.SelectedItem
        if ($selected -eq $script:tabExtrasRef -and $script:tabExtrasRef.IsEnabled -eq $false) {
            $script:konamiProgress = 0
            $script:mainTabsRef.SelectedIndex = 0
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
        # Hide extras tab and switch back to Signatures
        $script:tabExtrasRef.Visibility = 'Collapsed'
        $script:tabExtrasRef.IsEnabled  = $false
        $script:mainTabsRef.SelectedIndex = 0
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
        $script:txtStatusRef.Text = Get-Str 'StatusExtrasLocked'
    })

    # ── Unlock dialog (called from About section rapid-click) ────────────────
    function Show-UnlockDialog {
        if ($script:tabExtrasRef.Visibility -eq 'Visible') { return }

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
        Title="" Width="400" SizeToContent="Height"
        Background="$cBg" Foreground="$cFgPri"
        FontFamily="Segoe UI" FontSize="13"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize">
  <StackPanel Margin="28,24,28,24">
    <Border Background="$cSurface" CornerRadius="8" BorderBrush="$cBorder" BorderThickness="1" Padding="16,14" Margin="0,0,0,20">
      <TextBlock TextWrapping="Wrap" TextAlignment="Center" FontSize="12"
                 Foreground="$cFgSec" LineHeight="20"
                 Text="Only proceed if you REALLY know what you are doing."/>
    </Border>
    <TextBlock Text="Passphrase" Foreground="$cFgSec" FontSize="11" Margin="0,0,0,6"/>
    <PasswordBox Name="PwdInput"
                 Background="$cSurface" Foreground="$cFgPri" BorderBrush="$cBorder" BorderThickness="1"
                 Padding="8,6" FontSize="13"/>
    <Grid Margin="0,16,0,0">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="8"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>
      <Button Name="BtnCancel" Grid.Column="0" Content="Cancel"
              Background="$cSurface" Foreground="$cFgPri" BorderBrush="$cBorder" BorderThickness="1"
              Padding="0,8" FontSize="12"/>
      <Button Name="BtnUnlock" Grid.Column="2" Content="Proceed"
              Background="$cAccent" Foreground="$cOnAccent" BorderThickness="0"
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
            try {
                $script:tabExtrasRef.Visibility = 'Visible'
                $script:tabExtrasRef.IsEnabled  = $true
                $script:txtStatusRef.Text = Get-Str 'StatusExtrasUnlocked'
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
                $script:txtStatusRef.Text = Get-Str 'StatusUnlockError' $_
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
    if ($null -ne $script:previewEmptyHintRef) { $script:previewEmptyHintRef.Visibility = 'Visible' }

    try { & $script:refreshInboxList } catch { Set-Status (Get-Str 'StatusErrorMailboxes' $_) }
    # Note: refreshPermMailboxes is called after script-scope perm refs are stored (below)

    # -- Settings: apply saved theme and pre-select controls --

    # Restore unlocked state if Retro was previously unlocked
    if ($script:AppSettings.Theme -eq 'Retro') {
        $script:tabExtrasRef.Visibility = 'Visible'
        $script:tabExtrasRef.IsEnabled  = $true
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

    $script:lbThemeSelectorRef    = $lbThemeSelector
    $script:cbLanguageSelectorRef = $cbLanguageSelector

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

    # -- New --

    $btnNew.Add_Click({
        $name = Show-InputBox (Get-Str 'DlgNewSigPrompt') (Get-Str 'DlgNewSigTitle')
        if ([string]::IsNullOrWhiteSpace($name)) { return }
        try {
            New-Signature -Name $name
            & $script:refreshInboxList
            Set-Status (Get-Str 'StatusCreated' $name)
        } catch { Show-Error "$_"; Set-Status (Get-Str 'StatusCreateFailed') }
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
            Set-SelectedSigLabel $new
            & $script:refreshInboxList
            Set-Status (Get-Str 'StatusRenamed' $old $new)
        } catch { Show-Error "$_"; Set-Status (Get-Str 'StatusRenameFailed') }
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
            Set-SelectedSigLabel '(none selected)'
            $script:sigTxtSignatureInfo.Text = 'Select a mailbox or signature to preview.'
            $script:previewBrowserRef.Visibility = 'Collapsed'
            if ($null -ne $script:previewEmptyHintRef) { $script:previewEmptyHintRef.Visibility = 'Visible' }
            & $script:refreshInboxList
            Set-Status (Get-Str 'StatusDeleted' $name)
        } catch { Show-Error "$_"; Set-Status (Get-Str 'StatusDeleteFailed') }
    })

    # -- Edit signature in Outlook editor --

    $btnEditSig = $Window.FindName('BtnEditSignature')
    if ($null -ne $btnEditSig) {
        $btnEditSig.Add_Click({
            if ([string]::IsNullOrWhiteSpace($script:currentSig)) { return }
            $htmlPath = Get-SignatureHtmlPath -Name $script:currentSig
            if (-not (Test-Path $htmlPath)) {
                Show-Error "Cannot find file for '$($script:currentSig)'"
                return
            }
            # Locate outlook.exe
            $outlookExe = $null
            $regPaths = @(
                'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE',
                'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE'
            )
            foreach ($rp in $regPaths) {
                if (Test-Path $rp) {
                    $v = (Get-ItemProperty -Path $rp -ErrorAction SilentlyContinue).'(default)'
                    if ($v -and (Test-Path $v)) { $outlookExe = $v; break }
                }
            }
            if (-not $outlookExe) {
                # Fallback: search common Office install paths
                $candidates = @(
                    "$env:ProgramFiles\Microsoft Office\root\Office16\OUTLOOK.EXE",
                    "$env:ProgramFiles\Microsoft Office\Office16\OUTLOOK.EXE",
                    "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\OUTLOOK.EXE",
                    "${env:ProgramFiles(x86)}\Microsoft Office\Office16\OUTLOOK.EXE"
                )
                $outlookExe = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
            }
            if (-not $outlookExe) {
                Show-Error "Could not locate Outlook. Please open Outlook manually and edit the signature from there."
                return
            }
            $isRunning = Test-OutlookRunning
            if (-not $isRunning) {
                $answer = [System.Windows.MessageBox]::Show(
                    "Outlook is not running. Start Outlook to open the signature editor?",
                    'Start Outlook?',
                    [System.Windows.MessageBoxButton]::YesNo,
                    [System.Windows.MessageBoxImage]::Question)
                if ($answer -ne [System.Windows.MessageBoxResult]::Yes) { return }
            }
            # Open Outlook's signature editor for this signature
            Start-Process -FilePath $outlookExe -ArgumentList "/sig `"$($script:currentSig)`""
            Set-Status (Get-Str 'StatusOpened' $script:currentSig)
        })
    }

    # -- Reload preview (after external edit) --

    $btnReloadPrev = $Window.FindName('BtnReloadPreview')
    if ($null -ne $btnReloadPrev) {
        $btnReloadPrev.Add_Click({
            if (-not [string]::IsNullOrWhiteSpace($script:currentSig)) {
                Load-SignaturePreview $script:currentSig
                Set-Status (Get-Str 'StatusPreviewRefreshed')
            }
        })
    }

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
    $script:permCbPermLevel             = $cbPermLevel
    $script:permTxtMailboxEmptyHint     = $Window.FindName('TxtMailboxEmptyHint')
    # Easy tab script-scope refs
    $script:wizBtnGenerateOverview   = $btnGenerateOverview
    $script:wizPanelPermReport       = $panelPermReport
    $script:wizPanelPermReportCnt    = $panelPermReportCnt
    $script:wizLbMailboxes           = $lbWizMailboxes
    $script:wizTxtUserSearch         = $txtWizUserSearch
    $script:wizPopAdResults          = $popWizAdResults
    $script:wizLbAdResults           = $lbWizAdResults
    $script:wizPanelSelectedUsers    = $panelWizSelectedUsers
    $script:wizPanelFolders          = $panelWizFolders
    $script:wizPanelPermLevels       = $panelWizPermLevels
    $script:wizPanelSummary          = $panelWizSummary
    $script:wizTxtResult             = $txtWizResult

    # Initial mailbox load now that all script-scope refs are set
    try { & $refreshPermMailboxes } catch { Set-Status (Get-Str 'StatusMailboxesUnavailable' $_) }

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
            $rows = @(Get-FolderPermissions -EntryID $script:permSelectedFolder.EntryID `
                                            -StoreID  $script:permSelectedFolder.StoreID)
            $script:permDgPerms.ItemsSource = $null
            if ($rows.Count -gt 0) {
                $script:permDgPerms.ItemsSource = $rows
            }
            Set-PermStatus (Get-Str 'PermShowing' $rows.Count) ''
        } catch {
            Set-PermStatus (Get-Str 'PermCouldNotRead' $_) '#E74C3C'
        }
    }

    $script:permLastSelectedSmtp = $null   # track last selected mailbox across ItemsSource resets

    # Script-scope helpers called from event handlers
    function script:Set-PermRightEnabled {
        param([bool]$Enabled)
        $script:permPanelRight.IsEnabled = $Enabled
        $script:permPanelRight.Opacity   = if ($Enabled) { 1.0 } else { 0.45 }
    }

    function script:Get-FolderIcon {
        param([string]$Name, [int]$Depth)
        # Unicode glyphs from Segoe UI Symbol (well-supported on Windows 10/11)
        if ($Depth -eq 0) { return [char]0x2709 }  # ✉ mailbox root
        switch -Wildcard ($Name.ToLower()) {
            'inbox'          { return [char]0x25BC }   # ▼ incoming
            'sent*'          { return [char]0x25B2 }   # ▲ outgoing
            'deleted*'       { return [char]0x2715 }   # ✕ deleted
            'trash'          { return [char]0x2715 }   # ✕
            'drafts'         { return [char]0x270F }   # ✏ drafts
            'calendar*'      { return [char]0x25A6 }   # ▦ calendar
            'contacts*'      { return [char]0x263A }   # ☺ contacts
            'tasks'          { return [char]0x2611 }   # ☑ tasks
            'junk*'          { return [char]0x26A0 }   # ⚠ junk
            'spam'           { return [char]0x26A0 }   # ⚠
            'archive*'       { return [char]0x25A4 }   # ▤ archive
            'outbox'         { return [char]0x2191 }   # ↑ outbox
            'notes'          { return [char]0x270E }   # ✎ notes
            'journal'        { return [char]0x25D4 }   # ◔ journal
            default          { return [char]0x25B7 }   # ▷ generic folder
        }
    }

    # Initial state: right panel dimmed
    Set-PermRightEnabled $false

    # ── Mailbox selection → populate folder list ─────────────────────────────
    $lbMailboxes.Add_SelectionChanged({
        $sel = $script:permLbMailboxes.SelectedItem
        if ($null -eq $sel) { return }
        $script:permLastSelectedSmtp = $sel.SmtpAddress   # persist for search-filter reselect
        $script:permLbFolders.ItemsSource   = $null
        $script:permDgPerms.ItemsSource     = $null
        $script:permSelectedFolder          = $null
        $script:permTxtFolderHint.Visibility  = 'Visible'
        $script:permTxtFoldersHint.Text       = 'Loading folders...'
        Set-PermRightEnabled $false
        Set-PermStatus ''
        try {
            $raw = @(Get-MailboxFolders -SmtpAddress $sel.SmtpAddress)
            if ($raw.Count -eq 0) {
                $script:permTxtFoldersHint.Text = 'No folders found (is Outlook running?)'
            } else {
                $script:permTxtFoldersHint.Text = ''
                $items = [System.Collections.Generic.List[PSCustomObject]]::new()
                foreach ($f in $raw) {
                    $items.Add([PSCustomObject]@{
                        Name       = $f.Name
                        Icon       = Get-FolderIcon -Name $f.Name -Depth $f.Depth
                        FolderPath = $f.FolderPath
                        EntryID    = $f.EntryID
                        StoreID    = $f.StoreID
                        Depth      = $f.Depth
                        IndentPad  = [System.Windows.Thickness]::new([int]($f.Depth * 12), 1, 4, 1)
                    })
                }
                $script:permLbFolders.ItemsSource = $null
                $script:permLbFolders.ItemsSource = $items
            }
            Set-Status (Get-Str 'StatusLoadedFolders' $raw.Count $sel.Name)
        } catch {
            $script:permTxtFoldersHint.Text = Get-Str 'PermErrorFolders'
            Set-PermStatus (Get-Str 'PermCouldNotLoad' $_) '#E74C3C'
            Set-Status (Get-Str 'PermErrorFolders')
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
            Set-Status (Get-Str 'StatusPermissions' $sel.Name)
        } catch {
            Set-PermStatus (Get-Str 'PermErrorReading' $_) '#E74C3C'
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
                Set-PermStatus (Get-Str 'PermNoAdMatches') ''
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
            Set-PermStatus (Get-Str 'PermSelectFolder') '#E74C3C'; return
        }
        $user  = $script:permTxtAddUser.Text.Trim()
        $level = $script:permCbPermLevel.SelectedItem
        if ([string]::IsNullOrWhiteSpace($user)) {
            Set-PermStatus (Get-Str 'PermEnterName') '#E74C3C'; return
        }
        if ($null -eq $level) {
            Set-PermStatus (Get-Str 'PermSelectLevel') '#E74C3C'; return
        }
        try {
            $result = Set-FolderPermissionWithAncestors `
                                 -EntryID $script:permSelectedFolder.EntryID `
                                 -StoreID $script:permSelectedFolder.StoreID `
                                 -User    $user `
                                 -Level   $level
            Refresh-PermGrid
            $msg = Get-Str 'PermSaved' $user $level
            if ($result.AutoGranted.Count -gt 0) {
                $folderNames = ($result.AutoGranted | ForEach-Object { $_.FolderName }) -join ', '
                $msg += ' | ' + (Get-Str 'PermAutoGranted' $folderNames $user)
            }
            Set-PermStatus $msg '#27AE60'
            Set-Status $msg
        } catch {
            Set-PermStatus (Get-Str 'PermCouldNotSave' $_) '#E74C3C'
            Set-Status (Get-Str 'PermCouldNotSave' $_)
        }
    })

    # ── Remove permission (two-click confirm via flag) ───────────────────────
    $script:permRemovePending = $null   # user name pending confirmation
    $btnRemovePerm.Add_Click({
        if ($null -eq $script:permSelectedFolder) {
            $script:permRemovePending = $null
            Set-PermStatus (Get-Str 'PermSelectFolder') '#E74C3C'; return
        }
        $row = $script:permDgPerms.SelectedItem
        if ($null -eq $row) {
            $script:permRemovePending = $null
            Set-PermStatus (Get-Str 'PermSelectPerson') '#E74C3C'; return
        }
        if ($row.User -eq 'Default' -or $row.User -eq 'Anonymous') {
            $script:permRemovePending = $null
            Set-PermStatus (Get-Str 'PermCannotRemove' $row.User) '#E74C3C'
            return
        }
        if ($script:permRemovePending -ne $row.User) {
            # First click — ask for confirmation
            $script:permRemovePending = $row.User
            Set-PermStatus (Get-Str 'PermConfirmRemove' $row.User) ''
            return
        }
        # Second click — confirmed
        $script:permRemovePending = $null
        try {
            Remove-FolderPermission -EntryID $script:permSelectedFolder.EntryID `
                                    -StoreID $script:permSelectedFolder.StoreID `
                                    -User    $row.User
            Refresh-PermGrid
            Set-PermStatus (Get-Str 'PermRemoved' $row.User) '#27AE60'
            Set-Status (Get-Str 'PermRemoved' $row.User)
        } catch {
            Set-PermStatus (Get-Str 'PermCouldNotRemove' $_) '#E74C3C'
            Set-Status (Get-Str 'PermCouldNotRemove' $_)
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
            $script:permTxtFoldersHint.Text        = Get-Str 'TxtFoldersHint'
            Set-PermRightEnabled $false
            Set-PermStatus ''
            Set-Status (Get-Str 'StatusMailboxesRefreshed')
        } catch {
            Set-Status (Get-Str 'StatusErrorRefresh' $_)
        }
    })

    # ─── Easy tab: Overview + Wizard ────────────────────────────────────────────

    # ── Wizard state ───────────────────────────────────────────────────────────
    $script:wizCurrentStep     = 1
    $script:wizSelectedMailbox = $null
    $script:wizSelectedUsers   = [System.Collections.Generic.List[string]]::new()
    $script:wizSelectedFolders = [System.Collections.Generic.List[PSCustomObject]]::new()
    $script:wizSelectedLevel   = $null  # PSCustomObject with Label, Description, OlLevel
    $script:wizFolderCheckboxes = @()   # CheckBox controls for folder step

    function script:Show-WizStep([int]$step) {
        $script:wizCurrentStep = $step
        for ($i = 1; $i -le 5; $i++) {
            $panel = $script:windowRef.FindName("WizStep$i")
            if ($null -ne $panel) {
                $panel.Visibility = if ($i -eq $step) { 'Visible' } else { 'Collapsed' }
            }
        }
    }

    function script:Set-WizStatus([int]$step, [string]$msg, [string]$colour = '') {
        $ctrl = $script:windowRef.FindName("TxtWizStatus$step")
        if ($null -eq $ctrl) { return }
        $ctrl.Text = $msg
        if ($colour) {
            $ctrl.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($colour)
        } else {
            $ctrl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextSecondaryBrush')
        }
    }

    function script:Reset-Wizard {
        $script:wizSelectedMailbox = $null
        $script:wizSelectedUsers.Clear()
        $script:wizSelectedFolders.Clear()
        $script:wizSelectedLevel   = $null
        $script:wizFolderCheckboxes = @()
        if ($null -ne $script:wizLbMailboxes) { $script:wizLbMailboxes.SelectedIndex = -1 }
        if ($null -ne $script:wizTxtUserSearch) { $script:wizTxtUserSearch.Text = '' }
        if ($null -ne $script:wizPanelSelectedUsers) { $script:wizPanelSelectedUsers.Children.Clear() }
        if ($null -ne $script:wizPanelFolders) { $script:wizPanelFolders.Children.Clear() }
        if ($null -ne $script:wizPanelPermLevels) { $script:wizPanelPermLevels.Children.Clear() }
        if ($null -ne $script:wizPanelSummary) { $script:wizPanelSummary.Children.Clear() }
        if ($null -ne $script:wizTxtResult) { $script:wizTxtResult.Text = '' }
        for ($i = 1; $i -le 5; $i++) { Set-WizStatus $i '' }
        Show-WizStep 1
    }

    # ── Overview: Generate button ──────────────────────────────────────────────
    $btnGenerateOverview.Add_Click({
        $script:wizPanelPermReport.Visibility = 'Visible'
        $script:wizPanelPermReportCnt.Children.Clear()
        try {
            $overview = Get-PermissionsOverview
            $hasAny = $false
            foreach ($acct in $overview) {
                if ($acct.Entries.Count -eq 0) { continue }
                $hasAny = $true
                # Mailbox header
                $hdr = New-Object System.Windows.Controls.TextBlock
                $hdr.Text       = "$($acct.Mailbox) ($($acct.SmtpAddress))"
                $hdr.FontWeight = [System.Windows.FontWeights]::SemiBold
                $hdr.FontSize   = 13
                $hdr.Margin     = [System.Windows.Thickness]::new(0, 8, 0, 4)
                $hdr.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'AccentBrush')
                $script:wizPanelPermReportCnt.Children.Add($hdr) | Out-Null

                foreach ($entry in $acct.Entries) {
                    # User name
                    $userLbl = New-Object System.Windows.Controls.TextBlock
                    $userLbl.Text       = $entry.User
                    $userLbl.FontWeight = [System.Windows.FontWeights]::SemiBold
                    $userLbl.FontSize   = 12
                    $userLbl.Margin     = [System.Windows.Thickness]::new(12, 4, 0, 2)
                    $userLbl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextPrimaryBrush')
                    $script:wizPanelPermReportCnt.Children.Add($userLbl) | Out-Null

                    foreach ($f in $entry.Folders) {
                        # Row: folder name, level, remove button
                        $row = New-Object System.Windows.Controls.StackPanel
                        $row.Orientation = [System.Windows.Controls.Orientation]::Horizontal
                        $row.Margin      = [System.Windows.Thickness]::new(24, 1, 0, 1)

                        $fLabel = New-Object System.Windows.Controls.TextBlock
                        $fLabel.Text  = $f.FolderName
                        $fLabel.Width = 180
                        $fLabel.FontSize = 11
                        $fLabel.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextPrimaryBrush')
                        $row.Children.Add($fLabel) | Out-Null

                        $lLabel = New-Object System.Windows.Controls.TextBlock
                        $lLabel.Text  = $f.Level
                        $lLabel.Width = 160
                        $lLabel.FontSize = 11
                        $lLabel.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextSecondaryBrush')
                        $row.Children.Add($lLabel) | Out-Null

                        # Remove button with two-click confirm
                        $rmBtn = New-Object System.Windows.Controls.Button
                        $rmBtn.Content  = [char]0x2715  # ✕
                        $rmBtn.FontSize = 10
                        $rmBtn.Padding  = [System.Windows.Thickness]::new(4, 1, 4, 1)
                        $rmBtn.Tag      = [PSCustomObject]@{
                            EntryID  = $f.EntryID
                            StoreID  = $f.StoreID
                            User     = $entry.User
                            Row      = $row
                            Pending  = $false
                        }
                        $rmBtn.SetResourceReference([System.Windows.Controls.Control]::StyleProperty, 'SecondaryButton')
                        $rmBtn.Add_Click({
                            $tag = $this.Tag
                            if (-not $tag.Pending) {
                                # First click — confirm
                                $tag.Pending = $true
                                $this.Content = Get-Str 'OverviewRemoveConfirm'
                                $this.SetResourceReference([System.Windows.Controls.Control]::ForegroundProperty, 'AccentBrush')
                                return
                            }
                            # Second click — remove
                            try {
                                Remove-FolderPermission -EntryID $tag.EntryID -StoreID $tag.StoreID -User $tag.User
                                $parent = $tag.Row.Parent
                                if ($null -ne $parent) { $parent.Children.Remove($tag.Row) }
                                Set-Status (Get-Str 'OverviewRemoved' $tag.User $tag.EntryID)
                            } catch {
                                Set-Status (Get-Str 'OverviewRemoveError' $_)
                            }
                        })
                        $row.Children.Add($rmBtn) | Out-Null
                        $script:wizPanelPermReportCnt.Children.Add($row) | Out-Null
                    }
                }
            }
            if (-not $hasAny) {
                $empty = New-Object System.Windows.Controls.TextBlock
                $empty.Text       = Get-Str 'OverviewEmpty'
                $empty.FontStyle  = [System.Windows.FontStyles]::Italic
                $empty.FontSize   = 12
                $empty.Margin     = [System.Windows.Thickness]::new(0, 4, 0, 0)
                $empty.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextSecondaryBrush')
                $script:wizPanelPermReportCnt.Children.Add($empty) | Out-Null
            }
            Set-Status (Get-Str 'OverviewGenerated' $overview.Count)
        } catch {
            $err = New-Object System.Windows.Controls.TextBlock
            $err.Text = "Error: $_"
            $err.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#E74C3C')
            $script:wizPanelPermReportCnt.Children.Add($err) | Out-Null
        }
    })

    # ── Wizard Step 1: Next → validate mailbox selection ───────────────────────
    $Window.FindName('BtnWizNext1').Add_Click({
        $sel = $script:wizLbMailboxes.SelectedItem
        if ($null -eq $sel) {
            Set-WizStatus 1 (Get-Str 'WizSelectMailbox') '#E74C3C'
            return
        }
        $script:wizSelectedMailbox = $sel
        Set-WizStatus 1 ''
        Show-WizStep 2
    })

    # ── Wizard Step 2: AD search (own debounce timer) ──────────────────────────
    $script:wizAdLastQuery = ''
    $script:wizAdTimer     = $null

    function script:Invoke-WizAdSearchTick {
        if ($null -ne $script:wizAdTimer) { $script:wizAdTimer.Stop() }
        $script:wizAdTimer = $null
        try {
            $hits = Search-ADUsers -Query $script:wizAdLastQuery
            $script:wizLbAdResults.ItemsSource = $null
            if ($hits.Count -gt 0) {
                $script:wizLbAdResults.ItemsSource = $hits
                $script:wizPopAdResults.IsOpen = $true
            } else {
                $script:wizPopAdResults.IsOpen = $false
                Set-WizStatus 2 (Get-Str 'PermNoAdMatches') ''
            }
        } catch {
            $script:wizPopAdResults.IsOpen = $false
        }
    }

    $txtWizUserSearch.Add_TextChanged({
        $q = $script:wizTxtUserSearch.Text.Trim()
        if ($null -ne $script:wizAdTimer) { $script:wizAdTimer.Stop(); $script:wizAdTimer = $null }
        if ($q.Length -lt 2) { $script:wizPopAdResults.IsOpen = $false; return }
        $script:wizAdLastQuery = $q
        $t = New-Object System.Windows.Threading.DispatcherTimer
        $t.Interval = [System.TimeSpan]::FromMilliseconds(350)
        $t.Add_Tick({ Invoke-WizAdSearchTick })
        $script:wizAdTimer = $t
        $t.Start()
    })

    # AD result selected → add to selected users list
    $lbWizAdResults.Add_SelectionChanged({
        $hit = $script:wizLbAdResults.SelectedItem
        if ($null -eq $hit) { return }
        $value = if ($hit.Mail) { $hit.Mail } else { $hit.DisplayName }
        if ($script:wizSelectedUsers.Contains($value)) { return }
        $script:wizSelectedUsers.Add($value)
        $script:wizTxtUserSearch.Text = ''
        $script:wizPopAdResults.IsOpen = $false
        Set-WizStatus 2 ''
        # Build removable user pill
        $pill = New-Object System.Windows.Controls.StackPanel
        $pill.Orientation = [System.Windows.Controls.Orientation]::Horizontal
        $pill.Margin      = [System.Windows.Thickness]::new(0, 2, 0, 2)
        $lbl = New-Object System.Windows.Controls.TextBlock
        $lbl.Text     = $value
        $lbl.FontSize = 12
        $lbl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextPrimaryBrush')
        $lbl.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        $pill.Children.Add($lbl) | Out-Null
        $rmBtn = New-Object System.Windows.Controls.Button
        $rmBtn.Content  = [char]0x2715
        $rmBtn.FontSize = 9
        $rmBtn.Padding  = [System.Windows.Thickness]::new(4, 0, 4, 0)
        $rmBtn.Margin   = [System.Windows.Thickness]::new(6, 0, 0, 0)
        $rmBtn.Tag      = $value
        $rmBtn.SetResourceReference([System.Windows.Controls.Control]::StyleProperty, 'SecondaryButton')
        $rmBtn.Add_Click({
            $script:wizSelectedUsers.Remove($this.Tag)
            $parent = $this.Parent
            if ($null -ne $parent -and $null -ne $parent.Parent) { $parent.Parent.Children.Remove($parent) }
        })
        $pill.Children.Add($rmBtn) | Out-Null
        $script:wizPanelSelectedUsers.Children.Add($pill) | Out-Null
    })

    $Window.FindName('BtnWizBack2').Add_Click({ Show-WizStep 1 })
    $Window.FindName('BtnWizNext2').Add_Click({
        if ($script:wizSelectedUsers.Count -eq 0) {
            # Allow typing an email directly without AD search
            $typed = $script:wizTxtUserSearch.Text.Trim()
            if (-not [string]::IsNullOrWhiteSpace($typed)) {
                $script:wizSelectedUsers.Add($typed)
                $script:wizTxtUserSearch.Text = ''
                $pill = New-Object System.Windows.Controls.StackPanel
                $pill.Orientation = [System.Windows.Controls.Orientation]::Horizontal
                $pill.Margin      = [System.Windows.Thickness]::new(0, 2, 0, 2)
                $lbl = New-Object System.Windows.Controls.TextBlock
                $lbl.Text     = $typed
                $lbl.FontSize = 12
                $lbl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextPrimaryBrush')
                $pill.Children.Add($lbl) | Out-Null
                $script:wizPanelSelectedUsers.Children.Add($pill) | Out-Null
            }
        }
        if ($script:wizSelectedUsers.Count -eq 0) {
            Set-WizStatus 2 (Get-Str 'WizAddUser') '#E74C3C'
            return
        }
        Set-WizStatus 2 ''
        # Populate folder checkboxes for Step 3
        $script:wizPanelFolders.Children.Clear()
        $script:wizFolderCheckboxes = @()
        try {
            $raw = @(Get-MailboxFolders -SmtpAddress $script:wizSelectedMailbox.SmtpAddress)
            foreach ($f in $raw) {
                if ($f.Depth -eq 0) { continue }  # skip root folder
                $sp = New-Object System.Windows.Controls.StackPanel
                $sp.Orientation = [System.Windows.Controls.Orientation]::Horizontal
                $sp.Margin      = [System.Windows.Thickness]::new([int]($f.Depth * 12), 2, 0, 2)
                $cb = New-Object System.Windows.Controls.CheckBox
                $cb.Margin          = [System.Windows.Thickness]::new(0, 0, 6, 0)
                $cb.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
                $cb.Tag = $f
                # Auto-cascade: checking a parent checks all children
                $cb.Add_Checked({
                    $myDepth = $this.Tag.Depth
                    $found   = $false
                    foreach ($chk in $script:wizFolderCheckboxes) {
                        if ([object]::ReferenceEquals($chk, $this)) { $found = $true; continue }
                        if (-not $found) { continue }
                        if ($chk.Tag.Depth -le $myDepth) { break }  # sibling or parent level
                        $chk.IsChecked = $true
                    }
                })
                $cb.Add_Unchecked({
                    $myDepth = $this.Tag.Depth
                    $found   = $false
                    foreach ($chk in $script:wizFolderCheckboxes) {
                        if ([object]::ReferenceEquals($chk, $this)) { $found = $true; continue }
                        if (-not $found) { continue }
                        if ($chk.Tag.Depth -le $myDepth) { break }
                        $chk.IsChecked = $false
                    }
                })
                $sp.Children.Add($cb) | Out-Null
                $icon = New-Object System.Windows.Controls.TextBlock
                $icon.Text       = Get-FolderIcon -Name $f.Name -Depth $f.Depth
                $icon.FontFamily = New-Object System.Windows.Media.FontFamily('Segoe UI Symbol')
                $icon.FontSize   = 12
                $icon.Width      = 16
                $icon.Margin     = [System.Windows.Thickness]::new(0, 0, 5, 0)
                $icon.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextSecondaryBrush')
                $sp.Children.Add($icon) | Out-Null
                $name = New-Object System.Windows.Controls.TextBlock
                $name.Text     = $f.Name
                $name.FontSize = 11
                $name.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextPrimaryBrush')
                $sp.Children.Add($name) | Out-Null
                $script:wizPanelFolders.Children.Add($sp) | Out-Null
                $script:wizFolderCheckboxes += $cb
            }
        } catch {
            Set-WizStatus 3 "Error loading folders: $_" '#E74C3C'
        }
        Show-WizStep 3
    })

    # ── Wizard Step 3: Next → validate folder selection ────────────────────────
    $Window.FindName('BtnWizBack3').Add_Click({ Show-WizStep 2 })
    $Window.FindName('BtnWizNext3').Add_Click({
        $script:wizSelectedFolders.Clear()
        foreach ($cb in $script:wizFolderCheckboxes) {
            if ($cb.IsChecked) {
                $script:wizSelectedFolders.Add($cb.Tag)
            }
        }
        if ($script:wizSelectedFolders.Count -eq 0) {
            Set-WizStatus 3 (Get-Str 'WizSelectFolder') '#E74C3C'
            return
        }
        Set-WizStatus 3 ''
        # Populate permission levels for Step 4
        $script:wizPanelPermLevels.Children.Clear()
        $levels = Get-SimplifiedPermissionLevels
        foreach ($lvl in $levels) {
            $rb = New-Object System.Windows.Controls.RadioButton
            $rb.GroupName = 'WizPermLevel'
            $rb.Margin    = [System.Windows.Thickness]::new(0, 4, 0, 4)
            $rb.Tag       = $lvl
            $rbContent = New-Object System.Windows.Controls.StackPanel
            $lblMain = New-Object System.Windows.Controls.TextBlock
            $lblMain.Text       = $lvl.Label
            $lblMain.FontWeight = [System.Windows.FontWeights]::SemiBold
            $lblMain.FontSize   = 12
            $lblMain.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextPrimaryBrush')
            $rbContent.Children.Add($lblMain) | Out-Null
            $lblDesc = New-Object System.Windows.Controls.TextBlock
            $lblDesc.Text     = $lvl.Description
            $lblDesc.FontSize = 11
            $lblDesc.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextSecondaryBrush')
            $rbContent.Children.Add($lblDesc) | Out-Null
            $rb.Content = $rbContent
            $script:wizPanelPermLevels.Children.Add($rb) | Out-Null
        }
        Show-WizStep 4
    })

    # ── Wizard Step 4: Next → validate level selection, build summary ──────────
    $Window.FindName('BtnWizBack4').Add_Click({ Show-WizStep 3 })
    $Window.FindName('BtnWizNext4').Add_Click({
        $script:wizSelectedLevel = $null
        foreach ($child in $script:wizPanelPermLevels.Children) {
            if ($child -is [System.Windows.Controls.RadioButton] -and $child.IsChecked) {
                $script:wizSelectedLevel = $child.Tag
                break
            }
        }
        if ($null -eq $script:wizSelectedLevel) {
            Set-WizStatus 4 (Get-Str 'WizSelectLevel') '#E74C3C'
            return
        }
        Set-WizStatus 4 ''
        # Build summary for Step 5
        $script:wizPanelSummary.Children.Clear()
        $script:wizTxtResult.Text = ''
        function Add-SummaryRow([string]$label, [string]$value) {
            $row = New-Object System.Windows.Controls.StackPanel
            $row.Orientation = [System.Windows.Controls.Orientation]::Horizontal
            $row.Margin      = [System.Windows.Thickness]::new(0, 2, 0, 2)
            $lbl = New-Object System.Windows.Controls.TextBlock
            $lbl.Text       = "${label}:"
            $lbl.FontWeight = [System.Windows.FontWeights]::SemiBold
            $lbl.FontSize   = 12
            $lbl.Width      = 100
            $lbl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextSecondaryBrush')
            $row.Children.Add($lbl) | Out-Null
            $val = New-Object System.Windows.Controls.TextBlock
            $val.Text       = $value
            $val.FontSize   = 12
            $val.TextWrapping = [System.Windows.TextWrapping]::Wrap
            $val.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextPrimaryBrush')
            $row.Children.Add($val) | Out-Null
            $script:wizPanelSummary.Children.Add($row) | Out-Null
        }
        Add-SummaryRow (Get-Str 'WizSummaryMailbox') "$($script:wizSelectedMailbox.Name) ($($script:wizSelectedMailbox.SmtpAddress))"
        Add-SummaryRow (Get-Str 'WizSummaryUsers')   ($script:wizSelectedUsers -join ', ')
        Add-SummaryRow (Get-Str 'WizSummaryFolders') (($script:wizSelectedFolders | ForEach-Object { $_.Name }) -join ', ')
        Add-SummaryRow (Get-Str 'WizSummaryLevel')   $script:wizSelectedLevel.Label
        Show-WizStep 5
    })

    # ── Wizard Step 5: Apply ───────────────────────────────────────────────────
    $Window.FindName('BtnWizBack5').Add_Click({ Show-WizStep 4 })
    $Window.FindName('BtnWizApply').Add_Click({
        $errors   = 0
        $success  = 0
        $details  = @()
        $olLevel  = $script:wizSelectedLevel.OlLevel
        foreach ($user in $script:wizSelectedUsers) {
            foreach ($folder in $script:wizSelectedFolders) {
                try {
                    $result = Set-FolderPermissionWithAncestors `
                        -EntryID $folder.EntryID -StoreID $folder.StoreID `
                        -User $user -Level $olLevel
                    $success++
                    if ($result.AutoGranted.Count -gt 0) {
                        $folderNames = ($result.AutoGranted | ForEach-Object { $_.FolderName }) -join ', '
                        $details += "Also granted 'Can view' on $folderNames for $user"
                    }
                } catch {
                    $errors++
                    $details += "Error: $user on $($folder.Name): $_"
                }
            }
        }
        $msg = ''
        if ($errors -eq 0) {
            $msg = Get-Str 'WizApplySuccess' $script:wizSelectedUsers.Count $script:wizSelectedFolders.Count
        } else {
            $msg = Get-Str 'WizApplyPartial' $errors
        }
        if ($details.Count -gt 0) {
            $msg += "`n" + ($details -join "`n")
        }
        $script:wizTxtResult.Text = $msg
        $script:wizTxtResult.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString(
            $(if ($errors -eq 0) { '#27AE60' } else { '#E74C3C' })
        )
        Set-Status $msg
    })

    $Window.FindName('BtnWizReset').Add_Click({ Reset-Wizard })

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
