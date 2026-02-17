# ═══ Translations ════════════════════════════════════════════════════════════
$script:Strings = @{
    'en' = @{
        # Tab headers
        TabSignatures  = 'Signatures'
        TabPermissions = 'Permissions'
        TabSettings    = 'Settings'
        TabExtras      = 'Extras'
        # Settings section labels
        LblAppearance      = 'APPEARANCE'
        LblColourTheme     = 'Colour theme'
        LblLanguageSection = 'LANGUAGE'
        LblLanguage        = 'Language'
        # Signatures tab
        LblSignatures  = 'SIGNATURES'
        LblPreview     = 'PREVIEW'
        LblAssign      = 'Assign'
        BtnNew         = '+ New'
        BtnRename      = 'Rename'
        BtnDelete      = 'Delete'
        BtnEdit        = 'Edit in Outlook'
        BtnReload      = 'Reload preview'
        TxtPreviewHint = 'Select a signature on the left to preview it here.'
        TxtSelectedSig = '(select a signature above)'
        # Permissions tab
        LblMailboxes     = 'MAILBOXES'
        LblFolders       = 'FOLDERS'
        LblWhoHasAccess  = 'WHO HAS ACCESS'
        LblAddOrChange   = 'ADD OR CHANGE ACCESS'
        BtnRefreshPerm   = 'Refresh mailboxes'
        BtnSavePerm      = 'Save'
        BtnRemovePerm    = 'Remove'
        TxtMailboxHint   = 'Open Outlook, then click Refresh mailboxes.'
        TxtFoldersHint   = 'Select a mailbox to see its folders'
        TxtFolderHint    = 'Select a folder to see who has access'
        TbMailboxPlaceholder = 'Filter mailboxes...'
        TxtAddUserPlaceholder = 'Search by name or email...'
        ColPersonGroup   = 'Person or group'
        ColAccessLevel   = 'Access level'
        TipMailboxSearch = 'Type to filter the list below'
        TipAddUser       = 'Search Active Directory or type an email address directly'
        TipPermLevel     = 'Choose what this person is allowed to do'
        # Status messages
        StatusReady           = 'Ready'
        StatusNoAccounts      = 'No Outlook accounts found'
        StatusNoAccountsOl    = 'No Outlook accounts found - is Outlook running?'
        StatusLoadedMailboxes = 'Loaded {0} mailbox(es)'
        StatusLoadedFolders   = 'Loaded {0} folder(s) for {1}'
        StatusErrorMailboxes  = 'Error loading mailboxes: {0}'
        StatusCreated         = "Created '{0}' - click it, then use 'Edit in Outlook' to edit"
        StatusCreateFailed    = 'Create failed'
        StatusRenamed         = "Renamed '{0}' -> '{1}'"
        StatusRenameFailed    = 'Rename failed'
        StatusDeleted         = "Deleted '{0}'"
        StatusDeleteFailed    = 'Delete failed'
        StatusOpened          = "Opened '{0}' in Outlook signature editor"
        StatusPreviewRefreshed = 'Preview refreshed'
        StatusAssigned        = "Assigned '{0}' to {1}"
        StatusUnassigned      = "Unassigned '{0}' from {1}"
        StatusAssignFailed    = 'Assign failed'
        StatusMailboxesUnavailable = 'Mailboxes unavailable: {0}'
        StatusPermissions     = 'Permissions for: {0}'
        StatusMailboxesRefreshed = 'Mailboxes refreshed'
        StatusErrorRefresh    = 'Error refreshing mailboxes: {0}'
        StatusExtrasLocked    = 'Extras locked'
        StatusExtrasUnlocked  = 'Extras unlocked'
        StatusUnlockError     = 'Unlock error: {0}'
        # Permission status messages
        PermShowing           = 'Showing {0} permission(s)'
        PermCouldNotRead      = 'Could not read permissions: {0}'
        PermNoFolders         = 'No folders found (is Outlook running?)'
        PermCouldNotLoad      = 'Could not load folders: {0}'
        PermErrorReading      = 'Error reading permissions: {0}'
        PermNoAdMatches       = 'No AD matches - type an email address directly.'
        PermSelectFolder      = 'Please select a folder first.'
        PermEnterName         = 'Please enter a name or email address.'
        PermSelectLevel       = 'Please select an access level.'
        PermSaved             = 'Saved: {0} - {1}'
        PermAutoGranted       = "Also granted 'Can view' on {0} for {1}"
        PermCouldNotSave      = 'Could not save: {0}'
        PermSelectPerson      = 'Select a person in the list above to remove them.'
        PermCannotRemove      = "The '{0}' entry cannot be removed - change its level instead."
        PermConfirmRemove     = 'Remove access for {0}? Click Remove again to confirm.'
        PermRemoved           = 'Removed: {0}'
        PermCouldNotRemove    = 'Could not remove: {0}'
        PermErrorFolders      = 'Error loading folders'
        # Dialogs
        DlgNewSigPrompt = 'Enter a name for the new signature:'
        DlgNewSigTitle  = 'New Signature'
    }
    'de' = @{
        # Tab headers
        TabSignatures  = 'Signaturen'
        TabPermissions = 'Berechtigungen'
        TabSettings    = 'Einstellungen'
        TabExtras      = 'Extras'
        # Settings section labels
        LblAppearance      = 'DARSTELLUNG'
        LblColourTheme     = 'Farbschema'
        LblLanguageSection = 'SPRACHE'
        LblLanguage        = 'Sprache'
        # Signatures tab
        LblSignatures  = 'SIGNATUREN'
        LblPreview     = 'VORSCHAU'
        LblAssign      = 'Zuweisen'
        BtnNew         = '+ Neu'
        BtnRename      = 'Umbenennen'
        BtnDelete      = 'Löschen'
        BtnEdit        = 'In Outlook bearbeiten'
        BtnReload      = 'Vorschau neu laden'
        TxtPreviewHint = 'Wählen Sie links eine Signatur aus, um sie hier anzuzeigen.'
        TxtSelectedSig = '(Signatur oben auswählen)'
        # Permissions tab
        LblMailboxes     = 'POSTFÄCHER'
        LblFolders       = 'ORDNER'
        LblWhoHasAccess  = 'WER HAT ZUGRIFF'
        LblAddOrChange   = 'ZUGRIFF HINZUFÜGEN ODER ÄNDERN'
        BtnRefreshPerm   = 'Postfächer aktualisieren'
        BtnSavePerm      = 'Speichern'
        BtnRemovePerm    = 'Entfernen'
        TxtMailboxHint   = 'Outlook öffnen und auf „Postfächer aktualisieren" klicken.'
        TxtFoldersHint   = 'Postfach auswählen, um Ordner anzuzeigen'
        TxtFolderHint    = 'Ordner auswählen, um Zugriffsberechtigte anzuzeigen'
        TbMailboxPlaceholder = 'Postfächer filtern...'
        TxtAddUserPlaceholder = 'Nach Name oder E-Mail suchen...'
        ColPersonGroup   = 'Person oder Gruppe'
        ColAccessLevel   = 'Zugriffsebene'
        TipMailboxSearch = 'Zum Filtern der Liste eingeben'
        TipAddUser       = 'Active Directory durchsuchen oder E-Mail-Adresse eingeben'
        TipPermLevel     = 'Zugriffsebene für diese Person auswählen'
        # Status messages
        StatusReady           = 'Bereit'
        StatusNoAccounts      = 'Keine Outlook-Konten gefunden'
        StatusNoAccountsOl    = 'Keine Outlook-Konten gefunden – läuft Outlook?'
        StatusLoadedMailboxes = '{0} Postfach/Postfächer geladen'
        StatusLoadedFolders   = '{0} Ordner für {1} geladen'
        StatusErrorMailboxes  = 'Fehler beim Laden der Postfächer: {0}'
        StatusCreated         = "'{0}' erstellt – anklicken und mit 'In Outlook bearbeiten' bearbeiten"
        StatusCreateFailed    = 'Erstellen fehlgeschlagen'
        StatusRenamed         = "'{0}' umbenannt in '{1}'"
        StatusRenameFailed    = 'Umbenennen fehlgeschlagen'
        StatusDeleted         = "'{0}' gelöscht"
        StatusDeleteFailed    = 'Löschen fehlgeschlagen'
        StatusOpened          = "'{0}' im Outlook-Signatureneditor geöffnet"
        StatusPreviewRefreshed = 'Vorschau aktualisiert'
        StatusAssigned        = "'{0}' {1} zugewiesen"
        StatusUnassigned      = "'{0}' von {1} entfernt"
        StatusAssignFailed    = 'Zuweisen fehlgeschlagen'
        StatusMailboxesUnavailable = 'Postfächer nicht verfügbar: {0}'
        StatusPermissions     = 'Berechtigungen für: {0}'
        StatusMailboxesRefreshed = 'Postfächer aktualisiert'
        StatusErrorRefresh    = 'Fehler beim Aktualisieren der Postfächer: {0}'
        StatusExtrasLocked    = 'Extras gesperrt'
        StatusExtrasUnlocked  = 'Extras entsperrt'
        StatusUnlockError     = 'Entsperrfehler: {0}'
        # Permission status messages
        PermShowing           = '{0} Berechtigung(en) angezeigt'
        PermCouldNotRead      = 'Berechtigungen konnten nicht gelesen werden: {0}'
        PermNoFolders         = 'Keine Ordner gefunden (läuft Outlook?)'
        PermCouldNotLoad      = 'Ordner konnten nicht geladen werden: {0}'
        PermErrorReading      = 'Fehler beim Lesen der Berechtigungen: {0}'
        PermNoAdMatches       = 'Keine AD-Treffer – E-Mail-Adresse direkt eingeben.'
        PermSelectFolder      = 'Bitte zuerst einen Ordner auswählen.'
        PermEnterName         = 'Bitte Name oder E-Mail-Adresse eingeben.'
        PermSelectLevel       = 'Bitte eine Zugriffsebene auswählen.'
        PermSaved             = 'Gespeichert: {0} - {1}'
        PermAutoGranted       = "'Anzeigen' auch auf {0} für {1} gewährt"
        PermCouldNotSave      = 'Speichern fehlgeschlagen: {0}'
        PermSelectPerson      = 'Person in der Liste oben auswählen, um sie zu entfernen.'
        PermCannotRemove      = "Der Eintrag '{0}' kann nicht entfernt werden – Zugriffsebene stattdessen ändern."
        PermConfirmRemove     = 'Zugriff für {0} entfernen? Erneut auf „Entfernen" klicken zur Bestätigung.'
        PermRemoved           = 'Entfernt: {0}'
        PermCouldNotRemove    = 'Entfernen fehlgeschlagen: {0}'
        PermErrorFolders      = 'Fehler beim Laden der Ordner'
        # Dialogs
        DlgNewSigPrompt = 'Name für die neue Signatur eingeben:'
        DlgNewSigTitle  = 'Neue Signatur'
    }
}

function Get-Str {
    param(
        [string]$Key,
        [Parameter(ValueFromRemainingArguments=$true)]
        [object[]]$FormatArgs
    )
    $lang = if ($script:AppSettings -and $script:Strings.ContainsKey($script:AppSettings.Language)) {
        $script:AppSettings.Language
    } else { 'en' }
    $tbl = $script:Strings[$lang]
    $s   = if ($tbl.ContainsKey($Key)) { $tbl[$Key] } else { $script:Strings['en'][$Key] }
    if ($null -ne $FormatArgs -and $FormatArgs.Count -gt 0) { return [string]::Format($s, $FormatArgs) } else { return $s }
}

function Apply-Lang {
    param($Window)
    $w = $Window
    # Helper to safely set Text on a named element
    function ST([string]$name, [string]$key) {
        $el = $w.FindName($name)
        if ($null -ne $el) { $el.Text = Get-Str $key }
    }
    function SC([string]$name, [string]$key) {
        $el = $w.FindName($name)
        if ($null -ne $el) { $el.Content = Get-Str $key }
    }
    function SH([string]$name, [string]$key) {
        $el = $w.FindName($name)
        if ($null -ne $el) { $el.Header = Get-Str $key }
    }
    function STip([string]$name, [string]$key) {
        $el = $w.FindName($name)
        if ($null -ne $el) { $el.ToolTip = Get-Str $key }
    }
    function STag([string]$name, [string]$key) {
        $el = $w.FindName($name)
        if ($null -ne $el) { $el.Tag = Get-Str $key }
    }
    # Tab headers
    SH 'TabSignatures'  'TabSignatures'
    SH 'TabPermissions' 'TabPermissions'
    SH 'TabSettings'    'TabSettings'
    SH 'TabExtras'      'TabExtras'
    # Settings
    ST 'TxtLblAppearance'      'LblAppearance'
    ST 'TxtLblColourTheme'     'LblColourTheme'
    ST 'TxtLblLanguageSection' 'LblLanguageSection'
    ST 'TxtLblLanguage'        'LblLanguage'
    # Signatures tab
    ST  'TxtLblSignatures'    'LblSignatures'
    ST  'TxtLblPreview'       'LblPreview'
    ST  'TxtLblAssign'        'LblAssign'
    SC  'BtnNewSignature'     'BtnNew'
    SC  'BtnRenameSignature'  'BtnRename'
    SC  'BtnDeleteSignature'  'BtnDelete'
    SC  'BtnEditSignature'    'BtnEdit'
    SC  'BtnReloadPreview'    'BtnReload'
    ST  'TxtPreviewEmptyHint' 'TxtPreviewHint'
    # Permissions tab
    ST   'TxtLblMailboxes'           'LblMailboxes'
    ST   'TxtLblFolders'             'LblFolders'
    ST   'TxtLblWhoHasAccess'        'LblWhoHasAccess'
    ST   'TxtLblAddOrChange'         'LblAddOrChange'
    SC   'BtnRefreshPerm'            'BtnRefreshPerm'
    SC   'BtnSavePerm'               'BtnSavePerm'
    SC   'BtnRemovePerm'             'BtnRemovePerm'
    ST   'TxtMailboxEmptyHint'       'TxtMailboxHint'
    ST   'TxtFoldersHint'            'TxtFoldersHint'
    ST   'TxtFolderHint'             'TxtFolderHint'
    STag 'TbMailboxSearch' 'TbMailboxPlaceholder'
    STag 'TxtAddUser'     'TxtAddUserPlaceholder'
    STip 'TbMailboxSearch'           'TipMailboxSearch'
    STip 'TxtAddUser'                'TipAddUser'
    STip 'CbPermLevel'               'TipPermLevel'
    # DataGrid column headers
    $dg = $w.FindName('DgCurrentPerms')
    if ($null -ne $dg -and $dg.Columns.Count -ge 2) {
        $dg.Columns[0].Header = Get-Str 'ColPersonGroup'
        $dg.Columns[1].Header = Get-Str 'ColAccessLevel'
    }
    # Status bar initial text (only if it still shows the default)
    $sb = $w.FindName('TxtStatus')
    if ($null -ne $sb -and ($sb.Text -eq 'Ready' -or $sb.Text -eq 'Bereit')) {
        $sb.Text = Get-Str 'StatusReady'
    }
}
