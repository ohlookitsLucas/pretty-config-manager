function Initialize-SignaturesTab {
    param(
        [Parameter(Mandatory=$true)] $Window,
        [Parameter(Mandatory=$true)] [string] $ScriptRoot
    )

    # -- Control references --
    $btnRefreshMailboxes = $Window.FindName('BtnRefreshMailboxes')
    $btnNew              = $Window.FindName('BtnNewSignature')
    $btnRename           = $Window.FindName('BtnRenameSignature')
    $btnDelete           = $Window.FindName('BtnDeleteSignature')
    $previewBrowser      = $Window.FindName('PreviewBrowser')
    $script:previewBrowserRef     = $previewBrowser
    $script:previewEmptyHintRef   = $Window.FindName('TxtPreviewEmptyHint')
    $txtSignatureInfo    = $Window.FindName('TxtSignatureInfo')
    $panelInboxList      = $Window.FindName('PanelInboxList')
    $panelCopyTargets    = $Window.FindName('PanelCopyTargets')

    # Script-scope refs for controls used inside script:-scoped functions
    # NOTE: TxtSelectedSig is looked up via $script:windowRef.FindName at call-time to
    # avoid a $null reference when closures fire before the window fully renders.
    $script:sigTxtSignatureInfo  = $txtSignatureInfo
    $script:sigPanelInboxList    = $panelInboxList
    $script:sigPanelCopyTargets  = $panelCopyTargets
    $script:sigBtnEdit           = $Window.FindName('BtnEditSignature')
    $script:sigBtnReload         = $Window.FindName('BtnReloadPreview')

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

        # Colour TxtSelectedSig with MatchBrush when a signature is selected
        if ($null -ne $ctrl) {
            if (-not $isEmpty) {
                $ctrl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'MatchBrush')
            } else {
                $ctrl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'AccentBrush')
            }
        }

        # Highlight matching signature rows in the left panel (background + text colour)
        if ($null -ne $script:sigPanelInboxList) {
            foreach ($child in $script:sigPanelInboxList.Children) {
                if (-not ($child -is [System.Windows.Controls.Border])) { continue }
                $tag = $child.Tag
                if ($null -eq $tag -or $tag -isnot [string]) { continue }
                # Mailbox headers have registry paths as Tag (contain backslash); sig rows have plain names
                if ($tag.Contains('\')) { continue }
                $sp = $child.Child
                if ($null -eq $sp -or -not ($sp -is [System.Windows.Controls.StackPanel]) -or $sp.Children.Count -lt 2) { continue }
                if (-not $isEmpty -and $tag -eq $text) {
                    $child.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'AccentSubtleBrush')
                    $sp.Children[0].SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'MatchBrush')
                    $sp.Children[1].SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'MatchBrush')
                } else {
                    $child.Background = $null
                    $sp.Children[0].SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextSecondaryBrush')
                    $sp.Children[1].SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextPrimaryBrush')
                }
            }
        }

        # Highlight pills based on assignment state (no badge — hover popup shows details)
        if ($null -ne $script:sigPanelCopyTargets) {
            foreach ($pill in $script:sigPanelCopyTargets.Children) {
                if (-not ($pill -is [System.Windows.Controls.Border])) { continue }
                $sp = $pill.Child
                if (-not ($sp -is [System.Windows.Controls.StackPanel])) { continue }
                $pillTxt = $sp.Children[0]
                if ($isEmpty -or $null -eq $pill.Tag) {
                    $pill.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'SurfaceHoverBrush')
                    $pillTxt.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextPrimaryBrush')
                    continue
                }
                $isNew   = ($pill.Tag.NewSignature   -eq $text)
                $isReply = ($pill.Tag.ReplySignature -eq $text)
                $isAssigned = $isNew -or $isReply
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
        $bgHex = $script:Themes[$script:AppSettings.Theme].BgBrush
        if (Test-Path $htmlPath) {
            try {
                $rawHtml = Get-Content -Path $htmlPath -Raw -Encoding UTF8
                $wrappedHtml = "<!DOCTYPE html><html style=`"background:$bgHex;margin:0;padding:12px;overflow:hidden`"><body style=`"margin:0;padding:0;background:$bgHex;overflow:hidden`"><div style=`"background:white;border-radius:8px;padding:16px;overflow:hidden`">$rawHtml</div></body></html>"
                $script:previewBrowserRef.NavigateToString($wrappedHtml)
                $warn = if ($status.Warning) { " [!] $($status.Warning)" } else { '' }
                $script:sigTxtSignatureInfo.Text = "$name$warn"
            } catch {
                $script:sigTxtSignatureInfo.Text = Get-Str 'SigPreviewFailed' $_
            }
        } else {
            $fgHex2 = $script:Themes[$script:AppSettings.Theme].TextSecondaryBrush
            $script:previewBrowserRef.NavigateToString("<!DOCTYPE html><html style=`"background:$bgHex;overflow:hidden`"><body style=`"background:$bgHex;color:$fgHex2;font-family:Segoe UI;padding:16px;overflow:hidden`">No HTML file found.</body></html>")
            $script:sigTxtSignatureInfo.Text = "[!] Missing .htm file"
        }
    }

    # Select an account card (track selection state, no color change on mailbox name)
    function script:Select-AccountCard($card, [string]$regPath) {
        $script:selectedAccountKey = $regPath
    }

    # Toggle signature assignment for a pill (sets both New and Reply together)
    function script:Toggle-AssignPill($pill) {
        if ($null -eq $pill) { return }
        $ctrl = $script:windowRef.FindName('TxtSelectedSig')
        $sigName = if ($null -ne $ctrl) { $ctrl.Text } else { $script:selectedSigName }
        $noneLabels = @('(none selected)', '(select a signature above)', '(Signatur oben auswählen)')
        if ([string]::IsNullOrWhiteSpace($sigName) -or $sigName -in $noneLabels) {
            Show-Error (Get-Str 'SigSelectFirst')
            return
        }
        $assignHtmlPath = Get-SignatureHtmlPath -Name $sigName
        if (-not (Test-Path $assignHtmlPath)) {
            Show-Error "Signature '$sigName' no longer exists."
            Set-SelectedSigLabel '(none selected)'
            $script:selectedSigName = $null
            return
        }
        $regPath  = $pill.Tag.RegPath
        $curNew   = $pill.Tag.NewSignature
        $curReply = $pill.Tag.ReplySignature
        # Toggle: if either slot already has this sig, clear both; otherwise assign both
        $isAssigned = ($curNew -eq $sigName) -or ($curReply -eq $sigName)
        $newVal = if ($isAssigned) { '' } else { $sigName }
        try {
            Set-SignatureAssignment -RegistryPath $regPath -NewSignature $newVal -ReplySignature $newVal
            $accountName = $pill.Tag.AccountName
            Set-Status (Get-Str 'StatusAssigned' $sigName $accountName)
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
            $lbl.Text = Get-Str 'SigNoAccounts'
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

            # ── Separator line between mailboxes ──
            if (-not $isFirstAccount) {
                $sep = New-Object System.Windows.Controls.Border
                $sep.Height = 1
                $sep.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'BorderBrush')
                $sep.Margin = [System.Windows.Thickness]::new(4, 8, 4, 0)
                $sep.HorizontalAlignment = 'Stretch'
                $script:sigPanelInboxList.Children.Add($sep) | Out-Null
            }

            # ── Mailbox header (transparent border for click target + Tag) ──
            $mbBorder = New-Object System.Windows.Controls.Border
            $mbBorder.BorderThickness = [System.Windows.Thickness]::new(0)
            $mbBorder.Padding = [System.Windows.Thickness]::new(2, 0, 2, 0)
            $topMargin = if ($isFirstAccount) { 0 } else { 6 }
            $mbBorder.Margin  = [System.Windows.Thickness]::new(0, $topMargin, 0, 4)
            $mbBorder.Cursor  = [System.Windows.Input.Cursors]::Hand
            $mbBorder.Tag     = $regPath
            $isFirstAccount   = $false

            $mbSp = New-Object System.Windows.Controls.StackPanel

            $lblMbName = New-Object System.Windows.Controls.TextBlock
            $lblMbName.Text         = $accountDisplay
            $lblMbName.FontSize     = 16
            $lblMbName.FontWeight   = 'Bold'
            $lblMbName.TextTrimming = 'CharacterEllipsis'
            $lblMbName.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextPrimaryBrush')
            if (-not [string]::IsNullOrEmpty($smtpDisplay)) { $lblMbName.ToolTip = $smtpDisplay }
            $mbSp.Children.Add($lblMbName) | Out-Null

            if (-not [string]::IsNullOrEmpty($smtpDisplay) -and $smtpDisplay -ne $accountDisplay) {
                $lblSmtp = New-Object System.Windows.Controls.TextBlock
                $lblSmtp.Text         = $smtpDisplay
                $lblSmtp.FontSize     = 11
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
                $noSig.Text      = Get-Str 'SigNoneAssigned'
                $noSig.FontSize  = 10
                $noSig.FontStyle = 'Italic'
                $noSig.Margin    = [System.Windows.Thickness]::new(14, 2, 0, 4)
                $noSig.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextSecondaryBrush')
                $script:sigPanelInboxList.Children.Add($noSig) | Out-Null
            } else {
                foreach ($sn in $sigNames) {
                    $sigRow = New-Object System.Windows.Controls.Border
                    $sigRow.CornerRadius    = [System.Windows.CornerRadius]::new(4)
                    $sigRow.BorderThickness = [System.Windows.Thickness]::new(0)
                    $sigRow.Padding         = [System.Windows.Thickness]::new(8, 4, 8, 4)
                    $sigRow.Margin          = [System.Windows.Thickness]::new(14, 2, 4, 2)
                    $sigRow.Cursor          = [System.Windows.Input.Cursors]::Hand
                    $sigRow.Tag             = $sn

                    # Horizontal layout: bullet + name
                    $sigSp = New-Object System.Windows.Controls.StackPanel
                    $sigSp.Orientation = 'Horizontal'

                    $sigBullet = New-Object System.Windows.Controls.TextBlock
                    $sigBullet.Text     = [char]0x25CF   # ●
                    $sigBullet.FontSize = 8
                    $sigBullet.Margin   = [System.Windows.Thickness]::new(0, 3, 6, 0)
                    $sigBullet.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextSecondaryBrush')
                    $sigSp.Children.Add($sigBullet) | Out-Null

                    $sigTb = New-Object System.Windows.Controls.TextBlock
                    $sigTb.Text         = $sn
                    $sigTb.FontSize     = 12
                    $sigTb.TextTrimming = 'CharacterEllipsis'
                    $sigTb.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextPrimaryBrush')
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

            $pill.Child = $pillSp

            $pillRef = $pill
            $pill.Add_MouseLeftButtonUp(({
                param($s, $e)
                Toggle-AssignPill $pillRef
            }).GetNewClosure())

            $script:sigPanelCopyTargets.Children.Add($pill) | Out-Null
        }

        # ── Local Signatures section (unassigned sigs) ──
        if ($localOnlySigs.Count -gt 0) {
            $localHdr = New-Object System.Windows.Controls.TextBlock
            $localHdr.Text       = Get-Str 'LblLocalSignatures'
            $localHdr.FontSize   = 10
            $localHdr.FontWeight = 'SemiBold'
            $localHdr.Margin     = [System.Windows.Thickness]::new(2, 14, 0, 4)
            $localHdr.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextSecondaryBrush')
            $script:sigPanelInboxList.Children.Add($localHdr) | Out-Null

            foreach ($sn in $localOnlySigs) {
                $localRow = New-Object System.Windows.Controls.Border
                $localRow.CornerRadius    = [System.Windows.CornerRadius]::new(4)
                $localRow.BorderThickness = [System.Windows.Thickness]::new(0)
                $localRow.Padding         = [System.Windows.Thickness]::new(8, 4, 8, 4)
                $localRow.Margin          = [System.Windows.Thickness]::new(0, 2, 4, 2)
                $localRow.Cursor          = [System.Windows.Input.Cursors]::Hand
                $localRow.Tag             = $sn

                $localSp = New-Object System.Windows.Controls.StackPanel
                $localSp.Orientation = 'Horizontal'

                $localBullet = New-Object System.Windows.Controls.TextBlock
                $localBullet.Text     = [char]0x25CF   # ●
                $localBullet.FontSize = 8
                $localBullet.Margin   = [System.Windows.Thickness]::new(0, 3, 6, 0)
                $localBullet.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextSecondaryBrush')
                $localSp.Children.Add($localBullet) | Out-Null

                $localTb = New-Object System.Windows.Controls.TextBlock
                $localTb.Text         = $sn
                $localTb.FontSize     = 12
                $localTb.TextTrimming = 'CharacterEllipsis'
                $localTb.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextPrimaryBrush')
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

    # -- Initial load --
    $previewBrowser.Visibility = 'Collapsed'
    if ($null -ne $script:previewEmptyHintRef) { $script:previewEmptyHintRef.Visibility = 'Visible' }
    try { & $script:refreshInboxList } catch { Set-Status (Get-Str 'StatusErrorMailboxes' $_) }

    # -- Refresh mailboxes --
    $btnRefreshMailboxes.Add_Click({
        try {
            & $script:refreshInboxList
            Set-Status (Get-Str 'StatusMailboxesRefreshed')
        } catch { Set-Status (Get-Str 'StatusErrorMailboxes' $_) }
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
        if ([string]::IsNullOrWhiteSpace($old)) { Show-Error (Get-Str 'SigSelectFirst'); return }
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
        if ([string]::IsNullOrWhiteSpace($name)) { Show-Error (Get-Str 'SigSelectFirst'); return }
        if (-not (Confirm-Action "Delete signature '$name'? This cannot be undone." 'Delete Signature')) { return }
        try {
            Remove-Signature -Name $name
            $script:currentSig      = $null
            $script:selectedSigName = $null
            Set-SelectedSigLabel '(none selected)'
            $script:sigTxtSignatureInfo.Text = Get-Str 'SigSelectToPreview'
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
}
