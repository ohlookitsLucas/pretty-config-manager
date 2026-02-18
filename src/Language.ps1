# ═══ Translations ════════════════════════════════════════════════════════════
$script:Strings = @{
    'en' = @{
        # Tab headers (sub-tabs)
        TabSignatures  = 'Signatures'
        TabPermissions = 'Permissions'
        TabExtras      = 'Extras'
        # Navigation buttons
        NavOutlook       = 'Outlook'
        NavSkype         = 'Skype'
        NavMUD           = 'MUD'
        NavSettings      = 'Settings'
        NavExtras        = 'Extras'
        # Navigation dropdown menu items
        NavMenuSignatures      = 'Signatures'
        NavMenuPermissions     = 'Permissions'
        NavMenuDev             = 'Dev'
        # Dev page box labels
        TxtDevProfileGen       = 'Profile generator'
        TxtDevAdressbook       = 'Refresh Outlook adressbook'
        TxtDevChangeFont       = 'Outlook change font'
        # Skype page box labels
        TxtSkypeCacheMsg       = 'Clear cache'
        TxtSkypeGeneratorMsg   = 'Profile generator'
        # MUD placeholder
        TxtMUDMsg              = 'Feature not implemented'
        # Settings section labels
        LblAppearance      = 'APPEARANCE'
        LblColourTheme     = 'Colour theme'
        LblLanguageSection = 'LANGUAGE'
        LblLanguage        = 'Language'
        # Signatures tab
        LblSignatures  = 'SIGNATURES'
        LblPreview     = 'PREVIEW'
        LblAssign      = 'Assign'
        BtnRefreshMailboxes = 'Refresh mailboxes'
        BtnNew         = 'New'
        BtnRename      = 'Rename'
        BtnDelete      = 'Delete'
        BtnEdit        = 'Edit in Outlook'
        BtnReload      = 'Reload preview'
        TxtPreviewHint = 'Select a signature on the left to preview it here.'
        TxtSelectedSig = '(select a signature above)'
        # Permissions tab
        LblMailboxes     = 'MAILBOXES'
        LblFolders       = 'FOLDERS'
        LblWhoHasAccess  = 'PERMISSIONS'
        BtnRefreshPerm   = 'Refresh mailboxes'
        BtnSavePerm      = 'Save'
        BtnRemovePerm    = 'Remove'
        TxtMailboxHint   = 'Open Outlook, then click Refresh mailboxes.'
        TxtFoldersHint   = 'Select a mailbox to see its folders'
        TxtFolderHint    = 'Select a folder to see who has access'
        TxtAddUserPlaceholder = 'Search by name or email...'
        ColPersonGroup   = 'Person or group'
        ColAccessLevel   = 'Access level'
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
        StatusDevUnlocked     = 'DEV features unlocked'
        StatusExtrasLocked    = 'Extras locked'
        StatusExtrasUnlocked  = 'Extras unlocked'
        StatusUnlockError     = 'Unlock error: {0}'
        # Permission status messages
        PermShowing           = 'Showing {0} permission(s)'
        PermCouldNotRead      = 'Could not read permissions: {0}'
        PermLoadingFolders    = 'Loading folders...'
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
        # Permissions sub-tabs (Easy / Advanced)
        TabPermEasy           = 'Easy'
        TabPermAdvanced       = 'Advanced'
        # Overview section
        LblPermOverview       = 'PERMISSIONS OVERVIEW'
        TxtPermOverviewDesc   = 'See who has access to your mailboxes and calendars at a glance.'
        BtnGenerateOverview   = 'Generate overview'
        OverviewEmpty         = 'No custom permissions found across your mailboxes.'
        OverviewLoading       = 'Generating overview...'
        OverviewGenerated     = 'Overview generated for {0} mailbox(es)'
        OverviewRemoveConfirm = 'Sure?'
        OverviewRemoved       = 'Removed {0} from {1}'
        OverviewRemoveError   = 'Could not remove: {0}'
        # Wizard section
        LblPermWizard         = 'QUICK SETUP'
        TxtWizardDesc         = 'Set up folder permissions step by step.'
        WizStep1Title         = 'Step 1: Choose a mailbox you want to edit'
        WizStep2Title         = 'Step 2: Choose users you want to give permissions to'
        WizStep2Placeholder   = 'Search by name or email...'
        WizStep2Tip           = 'Search Active Directory or type an email address directly'
        WizStep3Title         = 'Step 3: Select folders you want to share'
        WizStep4Title         = 'Step 4: What should they be able to do?'
        WizStep5Title         = 'Step 5: Review and apply'
        BtnWizNext            = 'Next'
        BtnWizBack            = 'Back'
        BtnWizApply           = 'Apply permissions'
        BtnWizReset           = 'Start over'
        WizSelectMailbox      = 'Please select a mailbox first.'
        WizAddUser            = 'Please add at least one user.'
        WizSelectFolder       = 'Please select at least one folder.'
        WizSelectLevel        = 'Please select a permission level.'
        WizApplySuccess       = 'Permissions applied successfully for {0} user(s) on {1} folder(s).'
        WizApplyPartial       = 'Completed with {0} error(s). See details below.'
        WizRemoveUser         = 'Remove'
        WizSummaryMailbox     = 'Mailbox'
        WizSummaryUsers       = 'Users'
        WizSummaryFolders     = 'Folders'
        WizSummaryLevel       = 'Permission'
        # Simplified permission labels
        WizPermViewOnly       = 'Just look (read only)'
        WizPermViewOnlyDesc   = 'Can see items but cannot change anything'
        WizPermCreate         = 'Add new items'
        WizPermCreateDesc     = 'Can create new items in the folder'
        WizPermEditAll        = 'Add and edit everything'
        WizPermEditAllDesc    = 'Can create, edit, and delete any items'
        WizPermFull           = 'Full control'
        WizPermFullDesc       = 'Can do everything including manage settings'
        WizPermNone           = 'No access (block)'
        WizPermNoneDesc       = 'Cannot see or access this folder at all'
        # Dialogs
        DlgNewSigPrompt = 'Enter a name for the new signature:'
        DlgNewSigTitle  = 'New Signature'
    }
    'de' = @{
        # Tab headers (sub-tabs)
        TabSignatures  = 'Signaturen'
        TabPermissions = 'Berechtigungen'
        TabExtras      = 'Extras'
        # Navigation buttons
        NavOutlook       = 'Outlook'
        NavSkype         = 'Skype'
        NavMUD           = 'MUD'
        NavSettings      = 'Einstellungen'
        NavExtras        = 'Extras'
        # Navigation dropdown menu items
        NavMenuSignatures      = 'Signaturen'
        NavMenuPermissions     = 'Berechtigungen'
        NavMenuDev             = 'Dev'
        # Dev page box labels
        TxtDevProfileGen       = 'Profilgenerator'
        TxtDevAdressbook       = 'Outlook-Adressbuch aktualisieren'
        TxtDevChangeFont       = 'Outlook-Schriftart ändern'
        # Skype page box labels
        TxtSkypeCacheMsg       = 'Cache leeren'
        TxtSkypeGeneratorMsg   = 'Profilgenerator'
        # MUD placeholder
        TxtMUDMsg              = 'Funktion nicht implementiert'
        # Settings section labels
        LblAppearance      = 'DARSTELLUNG'
        LblColourTheme     = 'Farbschema'
        LblLanguageSection = 'SPRACHE'
        LblLanguage        = 'Sprache'
        # Signatures tab
        LblSignatures  = 'SIGNATUREN'
        LblPreview     = 'VORSCHAU'
        LblAssign      = 'Zuweisen'
        BtnRefreshMailboxes = 'Postfächer aktualisieren'
        BtnNew         = 'Neu'
        BtnRename      = 'Umbenennen'
        BtnDelete      = 'Löschen'
        BtnEdit        = 'In Outlook bearbeiten'
        BtnReload      = 'Vorschau neu laden'
        TxtPreviewHint = 'Wählen Sie links eine Signatur aus, um sie hier anzuzeigen.'
        TxtSelectedSig = '(Signatur oben auswählen)'
        # Permissions tab
        LblMailboxes     = 'POSTFÄCHER'
        LblFolders       = 'ORDNER'
        LblWhoHasAccess  = 'BERECHTIGUNGEN'
        BtnRefreshPerm   = 'Postfächer aktualisieren'
        BtnSavePerm      = 'Speichern'
        BtnRemovePerm    = 'Entfernen'
        TxtMailboxHint   = 'Outlook öffnen und auf „Postfächer aktualisieren" klicken.'
        TxtFoldersHint   = 'Postfach auswählen, um Ordner anzuzeigen'
        TxtFolderHint    = 'Ordner auswählen, um Zugriffsberechtigte anzuzeigen'
        TxtAddUserPlaceholder = 'Nach Name oder E-Mail suchen...'
        ColPersonGroup   = 'Person oder Gruppe'
        ColAccessLevel   = 'Zugriffsebene'
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
        StatusDevUnlocked     = 'DEV-Funktionen freigeschaltet'
        StatusExtrasLocked    = 'Extras gesperrt'
        StatusExtrasUnlocked  = 'Extras entsperrt'
        StatusUnlockError     = 'Entsperrfehler: {0}'
        # Permission status messages
        PermShowing           = '{0} Berechtigung(en) angezeigt'
        PermCouldNotRead      = 'Berechtigungen konnten nicht gelesen werden: {0}'
        PermLoadingFolders    = 'Ordner werden geladen...'
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
        # Permissions sub-tabs (Easy / Advanced)
        TabPermEasy           = 'Einfach'
        TabPermAdvanced       = 'Erweitert'
        # Overview section
        LblPermOverview       = 'BERECHTIGUNGSÜBERSICHT'
        TxtPermOverviewDesc   = 'Sehen Sie auf einen Blick, wer Zugriff auf Ihre Postfächer und Kalender hat.'
        BtnGenerateOverview   = 'Übersicht erstellen'
        OverviewEmpty         = 'Keine benutzerdefinierten Berechtigungen gefunden.'
        OverviewLoading       = 'Übersicht wird erstellt...'
        OverviewGenerated     = 'Übersicht für {0} Postfach/Postfächer erstellt'
        OverviewRemoveConfirm = 'Sicher?'
        OverviewRemoved       = '{0} von {1} entfernt'
        OverviewRemoveError   = 'Entfernen fehlgeschlagen: {0}'
        # Wizard section
        LblPermWizard         = 'SCHNELLEINRICHTUNG'
        TxtWizardDesc         = 'Ordnerberechtigungen Schritt für Schritt einrichten.'
        WizStep1Title         = 'Schritt 1: Postfach auswählen'
        WizStep2Title         = 'Schritt 2: Benutzer auswählen'
        WizStep2Placeholder   = 'Nach Name oder E-Mail suchen...'
        WizStep2Tip           = 'Active Directory durchsuchen oder E-Mail-Adresse eingeben'
        WizStep3Title         = 'Schritt 3: Ordner auswählen'
        WizStep4Title         = 'Schritt 4: Was dürfen sie tun?'
        WizStep5Title         = 'Schritt 5: Überprüfen und anwenden'
        BtnWizNext            = 'Weiter'
        BtnWizBack            = 'Zurück'
        BtnWizApply           = 'Berechtigungen anwenden'
        BtnWizReset           = 'Neu starten'
        WizSelectMailbox      = 'Bitte zuerst ein Postfach auswählen.'
        WizAddUser            = 'Bitte mindestens einen Benutzer hinzufügen.'
        WizSelectFolder       = 'Bitte mindestens einen Ordner auswählen.'
        WizSelectLevel        = 'Bitte eine Berechtigungsstufe auswählen.'
        WizApplySuccess       = 'Berechtigungen erfolgreich für {0} Benutzer auf {1} Ordner(n) angewendet.'
        WizApplyPartial       = 'Mit {0} Fehler(n) abgeschlossen. Details unten.'
        WizRemoveUser         = 'Entfernen'
        WizSummaryMailbox     = 'Postfach'
        WizSummaryUsers       = 'Benutzer'
        WizSummaryFolders     = 'Ordner'
        WizSummaryLevel       = 'Berechtigung'
        # Simplified permission labels
        WizPermViewOnly       = 'Nur ansehen (schreibgeschützt)'
        WizPermViewOnlyDesc   = 'Kann Elemente sehen, aber nichts ändern'
        WizPermCreate         = 'Neue Elemente hinzufügen'
        WizPermCreateDesc     = 'Kann neue Elemente im Ordner erstellen'
        WizPermEditAll        = 'Alles hinzufügen und bearbeiten'
        WizPermEditAllDesc    = 'Kann alle Elemente erstellen, bearbeiten und löschen'
        WizPermFull           = 'Volle Kontrolle'
        WizPermFullDesc       = 'Kann alles tun, einschließlich Einstellungen verwalten'
        WizPermNone           = 'Kein Zugriff (sperren)'
        WizPermNoneDesc       = 'Kann diesen Ordner nicht sehen oder darauf zugreifen'
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
    function Set-ElContent([string]$name, [string]$key) {
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
    # Navigation buttons
    Set-ElContent 'NavBtnOutlook'          'NavOutlook'
    Set-ElContent 'NavBtnSkype'            'NavSkype'
    Set-ElContent 'NavBtnMUD'              'NavMUD'
    Set-ElContent 'NavBtnExtras'           'NavExtras'
    # Note: NavBtnSettings uses a gear icon (&#x2699;), not translated text
    # Outlook dropdown items
    Set-ElContent 'NavMenuSignatures'      'NavMenuSignatures'
    Set-ElContent 'NavMenuPermissions'     'NavMenuPermissions'
    Set-ElContent 'NavMenuDev'             'NavMenuDev'
    # Dev page box labels
    ST 'TxtDevProfileGen'    'TxtDevProfileGen'
    ST 'TxtDevAdressbook'    'TxtDevAdressbook'
    ST 'TxtDevChangeFont'    'TxtDevChangeFont'
    # Skype page box labels
    ST 'TxtSkypeCacheMsg'     'TxtSkypeCacheMsg'
    ST 'TxtSkypeGeneratorMsg' 'TxtSkypeGeneratorMsg'
    ST 'TxtMUDMsg'            'TxtMUDMsg'
    # Settings
    ST 'TxtLblColourTheme'     'LblColourTheme'
    ST 'TxtLblLanguage'        'LblLanguage'
    # Signatures tab
    ST  'TxtLblAssign'        'LblAssign'
    STip 'BtnRefreshMailboxes' 'BtnRefreshMailboxes'
    Set-ElContent  'BtnNewSignature'     'BtnNew'
    Set-ElContent  'BtnRenameSignature'  'BtnRename'
    Set-ElContent  'BtnDeleteSignature'  'BtnDelete'
    Set-ElContent  'BtnEditSignature'    'BtnEdit'
    STip 'BtnReloadPreview'    'BtnReload'
    ST  'TxtPreviewEmptyHint' 'TxtPreviewHint'
    # Permissions tab
    ST   'TxtLblMailboxes'           'LblMailboxes'
    ST   'TxtLblFolders'             'LblFolders'
    ST   'TxtLblWhoHasAccess'        'LblWhoHasAccess'
    Set-ElContent   'BtnRefreshPerm'            'BtnRefreshPerm'
    Set-ElContent   'BtnSavePerm'               'BtnSavePerm'
    Set-ElContent   'BtnRemovePerm'             'BtnRemovePerm'
    ST   'TxtMailboxEmptyHint'       'TxtMailboxHint'
    ST   'TxtFoldersHint'            'TxtFoldersHint'
    ST   'TxtFolderHint'             'TxtFolderHint'
    STag 'TxtAddUser'     'TxtAddUserPlaceholder'
    STip 'TxtAddUser'                'TipAddUser'
    # Permissions sub-tabs
    SH   'TabPermEasy'              'TabPermEasy'
    SH   'TabPermAdvanced'          'TabPermAdvanced'
    # Overview section
    ST   'TxtLblPermOverview'       'LblPermOverview'
    ST   'TxtPermOverviewDesc'      'TxtPermOverviewDesc'
    Set-ElContent   'BtnGenerateOverview'      'BtnGenerateOverview'
    # Wizard section
    ST   'TxtLblPermWizard'         'LblPermWizard'
    ST   'TxtWizardDesc'            'TxtWizardDesc'
    ST   'TxtWizStep1'              'WizStep1Title'
    ST   'TxtWizStep2'              'WizStep2Title'
    STag 'TxtWizUserSearch'         'WizStep2Placeholder'
    STip 'TxtWizUserSearch'         'WizStep2Tip'
    ST   'TxtWizStep3'              'WizStep3Title'
    ST   'TxtWizStep4'              'WizStep4Title'
    ST   'TxtWizStep5'              'WizStep5Title'
    Set-ElContent   'BtnWizNext1'              'BtnWizNext'
    Set-ElContent   'BtnWizNext2'              'BtnWizNext'
    Set-ElContent   'BtnWizNext3'              'BtnWizNext'
    Set-ElContent   'BtnWizNext4'              'BtnWizNext'
    Set-ElContent   'BtnWizBack2'              'BtnWizBack'
    Set-ElContent   'BtnWizBack3'              'BtnWizBack'
    Set-ElContent   'BtnWizBack4'              'BtnWizBack'
    Set-ElContent   'BtnWizBack5'              'BtnWizBack'
    Set-ElContent   'BtnWizApply'              'BtnWizApply'
    Set-ElContent   'BtnWizReset'              'BtnWizReset'
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
