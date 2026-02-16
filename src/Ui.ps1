function Initialize-Ui {
    param(
        [Parameter(Mandatory=$true)] $Window,
        [Parameter(Mandatory=$true)] [string] $ScriptRoot
    )

    # -- Control references --
    # Signatures tab
    $lbSignatures        = $Window.FindName('LbSignatures')
    $btnNew              = $Window.FindName('BtnNewSignature')
    $btnDuplicate        = $Window.FindName('BtnDuplicateSignature')
    $btnRename           = $Window.FindName('BtnRenameSignature')
    $btnDelete           = $Window.FindName('BtnDeleteSignature')
    $btnRefreshSigs      = $Window.FindName('BtnRefreshSignatures')
    $txtEditorTitle      = $Window.FindName('TxtEditorTitle')
    $btnToggleEditor     = $Window.FindName('BtnToggleEditor')
    $btnSave             = $Window.FindName('BtnSaveSignature')
    $previewBrowser      = $Window.FindName('PreviewBrowser')
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
    $lbAccounts          = $Window.FindName('LbAccounts')
    $btnRefreshAccounts  = $Window.FindName('BtnRefreshAccounts')
    $btnSetPermission    = $Window.FindName('BtnSetPermission')
    $txtPermissionUser   = $Window.FindName('TxtPermissionUser')
    $cbPermissionLevel   = $Window.FindName('CbPermissionLevel')
    $lbPermissions       = $Window.FindName('LbPermissions')

    # Status bar
    $txtStatus           = $Window.FindName('TxtStatus')

    # -- State --
    $script:editorMode   = 'preview'   # 'preview' | 'html'
    $script:currentSig   = $null
    $script:assignRows   = @{}         # RegistryPath -> @{ NewCb; ReplyCb }

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

        if (Test-Path $htmlPath) {
            try {
                $previewBrowser.Navigate((New-Object System.Uri($htmlPath)))
                $warn = if ($status.Warning) { " [!] $($status.Warning)" } else { '' }
                $txtSignatureInfo.Text = "$name$warn"
            } catch {
                $txtSignatureInfo.Text = "Preview failed: $_"
            }
        } else {
            $previewBrowser.NavigateToString('<html><body style="background:#1E1E2E;color:#9898A8;font-family:Segoe UI;padding:16px">No HTML file found.</body></html>')
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
        $lblAccount.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#EAEAEF')
        $lblAccount.FontSize = 11
        $lblAccount.VerticalAlignment = 'Center'
        $lblAccount.TextTrimming = 'CharacterEllipsis'
        $lblAccount.ToolTip = "$($assignment.AccountName)`n$($assignment.RegistryPath)"
        [System.Windows.Controls.Grid]::SetColumn($lblAccount, 0)

        # Helper: build a small sig picker ListBox
        function New-SigPicker([string]$currentValue) {
            $lb = New-Object System.Windows.Controls.ListBox
            $lb.Background = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#1E1E2E')
            $lb.BorderThickness = [System.Windows.Thickness]::new(1)
            $lb.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#3E3E55')
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
            $lbl.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#9898A8')
            $lbl.FontSize = 12
            $lbl.Margin = [System.Windows.Thickness]::new(4, 8, 0, 0)
            $panelRows.Children.Add($lbl) | Out-Null
        } else {
            foreach ($a in $assignments) { Add-AssignmentRow $a $sigNames }
        }
        $txtOutlookWarning.Visibility = if (Test-OutlookRunning) { 'Visible' } else { 'Collapsed' }
    }

    # Reusable: refresh accounts list (Permissions tab)
    $refreshAccounts = {
        $accounts = Get-SignedInAccounts
        $lbAccounts.Items.Clear()
        foreach ($a in $accounts) { $lbAccounts.Items.Add("$($a.Name) <$($a.SmtpAddress)>") }
        return $accounts
    }

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

    # -- Initial load --

    $previewBrowser.NavigateToString('<html><body style="background:#1E1E2E;margin:0;padding:0;overflow:hidden"></body></html>')

    try { & $refreshSignatures } catch { Set-Status "Error loading signatures: $_" }
    try { & $refreshAssignments } catch { Set-Status "Error loading assignments: $_" }
    try { & $refreshAccounts | Out-Null } catch { }  # silent - Outlook COM may not be available

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
            $previewBrowser.NavigateToString('<html><body style="background:#1E1E2E;margin:0;padding:0;"></body></html>')
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

    # -- Permissions tab --

    $btnRefreshAccounts.Add_Click({
        try { & $refreshAccounts | Out-Null; Set-Status "Refreshed accounts" }
        catch { Set-Status "Error refreshing accounts: $_" }
    })

    $btnSetPermission.Add_Click({
        $targetUser = $txtPermissionUser.Text.Trim()
        $level      = $cbPermissionLevel.SelectedItem
        if ([string]::IsNullOrWhiteSpace($targetUser) -or $null -eq $level) {
            Show-Error 'Enter a user and select a permission level.'
            return
        }
        $levelText = $level.Content
        if (-not (Confirm-Action "Set '$levelText' for $targetUser?")) { return }
        try {
            Set-CalendarPermission -User $targetUser -Level $levelText -Confirm:$false
            Set-Status "Permission set for ${targetUser}: $levelText"
        } catch { Show-Error "$_"; Set-Status "Permission set failed" }
    })
}
