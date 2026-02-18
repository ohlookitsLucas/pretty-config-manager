function Initialize-WizardTab {
    param(
        [Parameter(Mandatory=$true)] $Window,
        [Parameter(Mandatory=$true)] [string] $ScriptRoot
    )

    # -- Control references --
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

    # Bind mailbox list (data already loaded by coordinator into $script:permAllAccounts)
    $lbWizMailboxes.ItemsSource = $script:permAllAccounts

    # ── Wizard state ──────────────────────────────────────────────────────────
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

    Reset-Wizard
}
