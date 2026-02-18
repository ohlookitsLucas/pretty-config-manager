function Initialize-PermissionsTab {
    param(
        [Parameter(Mandatory=$true)] $Window,
        [Parameter(Mandatory=$true)] [string] $ScriptRoot
    )

    # -- Control references --
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

    # Store controls needed by AD search functions at script scope
    $script:permLbAdResults = $lbAdResults
    $script:permPopAddUser  = $popAddUser
    $script:permTxtAddUser  = $txtAddUser

    # Reusable: refresh mailbox list — updates $script:permAllAccounts (set by coordinator)
    # and re-binds both mailbox list controls
    $script:refreshPermMailboxes = {
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

    # Bind mailbox list (data already loaded by coordinator into $script:permAllAccounts)
    $lbMailboxes.ItemsSource = $script:permAllAccounts
    $isEmpty = $script:permAllAccounts.Count -eq 0
    if ($null -ne $script:permTxtMailboxEmptyHint) {
        $script:permTxtMailboxEmptyHint.Visibility = if ($isEmpty) { 'Visible' } else { 'Collapsed' }
    }

    # ── Mailbox selection → populate folder list ─────────────────────────────
    $lbMailboxes.Add_SelectionChanged({
        $sel = $script:permLbMailboxes.SelectedItem
        if ($null -eq $sel) { return }
        $script:permLastSelectedSmtp = $sel.SmtpAddress   # persist for search-filter reselect
        $script:permLbFolders.ItemsSource   = $null
        $script:permDgPerms.ItemsSource     = $null
        $script:permSelectedFolder          = $null
        $script:permTxtFolderHint.Visibility    = 'Visible'
        $script:permTxtFoldersHint.Text         = 'Loading folders...'
        $script:permTxtFoldersHint.Visibility   = 'Visible'
        Set-PermRightEnabled $false
        Set-PermStatus ''
        try {
            $raw = @(Get-MailboxFolders -SmtpAddress $sel.SmtpAddress)
            if ($raw.Count -eq 0) {
                $script:permTxtFoldersHint.Text = 'No folders found (is Outlook running?)'
            } else {
                $script:permTxtFoldersHint.Visibility = 'Collapsed'
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
            $script:permTxtFoldersHint.Text       = Get-Str 'PermErrorFolders'
            $script:permTxtFoldersHint.Visibility = 'Visible'
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
            & $script:refreshPermMailboxes
            $script:permLbFolders.ItemsSource      = $null
            $script:permDgPerms.ItemsSource        = $null
            $script:permSelectedFolder             = $null
            $script:permTxtFolderHint.Visibility     = 'Visible'
            $script:permTxtFoldersHint.Text          = Get-Str 'TxtFoldersHint'
            $script:permTxtFoldersHint.Visibility    = 'Visible'
            Set-PermRightEnabled $false
            Set-PermStatus ''
            Set-Status (Get-Str 'StatusMailboxesRefreshed')
        } catch {
            Set-Status (Get-Str 'StatusErrorRefresh' $_)
        }
    })
}
