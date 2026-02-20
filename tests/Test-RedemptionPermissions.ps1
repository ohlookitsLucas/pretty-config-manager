<#
  Test-RedemptionPermissions.ps1
  Standalone tool to test Exchange folder permission management via Redemption.dll.
  Replicates the PCM Advanced Permissions tab layout using RDOSession + RDOACL
  instead of Outlook COM FolderPermissions.

  Usage:  powershell -File tests\Test-RedemptionPermissions.ps1
  Requires: Outlook running (or default MAPI profile), Redemption COM registered.
#>

# ═══ WPF assemblies ═══════════════════════════════════════════════════════════
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# ═══ Redemption session ═══════════════════════════════════════════════════════
try {
    $rdoSession = New-Object -ComObject Redemption.RDOSession
} catch {
    [System.Windows.MessageBox]::Show(
        "Could not create Redemption.RDOSession COM object.`nMake sure Redemption is registered (regsvr32 Redemption64.dll).`n`n$_",
        'Redemption Permissions Test', 'OK', 'Error') | Out-Null
    return
}
try {
    $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
    $rdoSession.MAPIOBJECT = $outlook.Session.MAPIOBJECT
} catch {
    try {
        $rdoSession.Logon()
    } catch {
        [System.Windows.MessageBox]::Show(
            "Could not connect to MAPI.`nMake sure Outlook is running or a default profile exists.`n`n$_",
            'Redemption Permissions Test', 'OK', 'Error') | Out-Null
        return
    }
}

# ═══ MAPI rights constants ════════════════════════════════════════════════════
$RIGHTS_READ_ITEMS        = 0x0001
$RIGHTS_CREATE_ITEMS      = 0x0002
$RIGHTS_EDIT_OWN          = 0x0008
$RIGHTS_DELETE_OWN        = 0x0010
$RIGHTS_EDIT_ALL          = 0x0020
$RIGHTS_DELETE_ALL        = 0x0040
$RIGHTS_CREATE_SUBFOLDERS = 0x0080
$RIGHTS_FOLDER_OWNER      = 0x0100
$RIGHTS_FOLDER_CONTACT    = 0x0200
$RIGHTS_FOLDER_VISIBLE    = 0x0400

# Friendly name -> MAPI rights bitmask (ordered for ComboBox display)
$FriendlyToRights = [ordered]@{
    'No access'                   = 0x0000
    'Can view'                    = 0x0401   # ROLE_REVIEWER
    'Can create items'            = 0x0402   # ROLE_CONTRIBUTOR
    'Can view & create'           = 0x0403   # READ + CREATE + VISIBLE
    'Can create, edit & delete'   = 0x0413   # ROLE_NONEDITING_AUTHOR
    'Can create & edit own items' = 0x041B   # ROLE_AUTHOR
    'Can create & edit all items' = 0x047B   # ROLE_EDITOR
    'Full access'                 = 0x04FB   # ROLE_PUBLISH_EDITOR
    'Owner'                       = 0x07FB   # ROLE_OWNER
}

# Reverse: MAPI rights bitmask -> friendly name
# Uses string keys to avoid PowerShell's OrderedDictionary int-key indexing trap
# (int keys are treated as positional indices, not key lookups)
$RightsToFriendly = @{
    '07FB' = 'Owner'
    '04FB' = 'Full access'
    '047B' = 'Can create & edit all items'
    '041B' = 'Can create & edit own items'
    '0413' = 'Can create, edit & delete'
    '0403' = 'Can view & create'
    '0402' = 'Can create items'
    '0401' = 'Can view'
    '0400' = 'Folder visible only'
    '0000' = 'No access'
}

# Bit flag names for diagnostic decomposition
# Uses arrays of [flag, name] to avoid OrderedDictionary int-key indexing trap
$RightsBitFlags = @(
    @(0x0001, 'READ'),
    @(0x0002, 'CREATE'),
    @(0x0008, 'EDIT_OWN'),
    @(0x0010, 'DELETE_OWN'),
    @(0x0020, 'EDIT_ALL'),
    @(0x0040, 'DELETE_ALL'),
    @(0x0080, 'CREATE_SUBFOLDERS'),
    @(0x0100, 'FOLDER_OWNER'),
    @(0x0200, 'FOLDER_CONTACT'),
    @(0x0400, 'FOLDER_VISIBLE')
)

function Get-FriendlyRightsName {
    param([int]$Rights)
    $hex = $Rights.ToString('X4')
    if ($RightsToFriendly.ContainsKey($hex)) { return $RightsToFriendly[$hex] }
    return "Custom (0x$hex)"
}

function Get-RightsDecomposition {
    param([int]$Rights)
    $parts = @()
    foreach ($entry in $RightsBitFlags) {
        if ($Rights -band $entry[0]) { $parts += $entry[1] }
    }
    if ($parts.Count -eq 0) { return 'NONE' }
    return $parts -join ' | '
}

# ═══ Hidden folder filter ═════════════════════════════════════════════════════
$HiddenFolderNames = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase)
@(
    'Yammer Root', 'Conversation Action Settings', 'Conversation History',
    'Synchronisierungsprobleme', 'Sync Issues', 'Quick Step Settings',
    'ExternalContacts', 'GAL Contacts', 'Recipient Cache',
    'Social Activity Notifications', 'Files', 'OneNote',
    'Suggested Contacts', 'PersonMetadata', 'Recoverable Items',
    'Purges', 'Versions', 'DiscoveryHolds', 'Audits',
    'Calendar Logging', 'Conflicts', 'Local Failures', 'Server Failures',
    'Common Views', 'Finder', 'Reminders', 'Schedule',
    'To-Do', 'Tracked Mail Processing', 'PeopleConnect',
    'Orion Notes', 'Spooler Queue', 'Wichtige E-Mails'
) | ForEach-Object { $HiddenFolderNames.Add($_) | Out-Null }

# ═══ Folder icon function ═════════════════════════════════════════════════════
function Get-FolderIcon {
    param([string]$Name, [int]$Depth)
    if ($Depth -eq 0) { return [char]0x2709 }
    switch -Wildcard ($Name.ToLower()) {
        'inbox'     { return [char]0x25BC }
        'sent*'     { return [char]0x25B2 }
        'deleted*'  { return [char]0x2715 }
        'trash'     { return [char]0x2715 }
        'drafts'    { return [char]0x270F }
        'calendar*' { return [char]0x25A6 }
        'contacts*' { return [char]0x263A }
        'tasks'     { return [char]0x2611 }
        'junk*'     { return [char]0x26A0 }
        'spam'      { return [char]0x26A0 }
        'archive*'  { return [char]0x25A4 }
        'outbox'    { return [char]0x2191 }
        'notes'     { return [char]0x270E }
        'journal'   { return [char]0x25D4 }
        default     { return [char]0x25B7 }
    }
}

# ═══ Redemption account/folder discovery ══════════════════════════════════════

function Get-RdoAccounts {
    $accounts = @()
    $stores = $rdoSession.Stores
    for ($i = 1; $i -le $stores.Count; $i++) {
        $store = $stores.Item($i)
        $name = ''
        try { $name = $store.Name } catch { continue }
        $smtp = ''
        try { $smtp = $store.Owner.SmtpAddress } catch {
            try { $smtp = $store.Owner.Address } catch {}
        }
        if (-not [string]::IsNullOrEmpty($name)) {
            $accounts += [PSCustomObject]@{
                Name       = $name
                SmtpAddress = $smtp
                StoreIndex = $i
            }
        }
    }
    return $accounts
}

function Get-RdoFoldersRecursive {
    param($Folder, [int]$Depth = 0, [string]$StoreID)
    $results = @()
    $folderName = ''
    try { $folderName = $Folder.Name } catch { return $results }

    if ($Depth -gt 0 -and $HiddenFolderNames.Contains($folderName)) {
        return $results
    }

    $entryID = ''
    try { $entryID = $Folder.EntryID } catch {}
    $folderPath = ''
    try { $folderPath = $Folder.FolderPath } catch { $folderPath = $folderName }

    $results += [PSCustomObject]@{
        Name       = $folderName
        FolderPath = $folderPath
        EntryID    = $entryID
        StoreID    = $StoreID
        Depth      = $Depth
        RdoFolder  = $Folder
    }

    try {
        $subs = $Folder.Folders
        if ($null -ne $subs -and $subs.Count -gt 0) {
            for ($i = 1; $i -le $subs.Count; $i++) {
                $sub = $subs.Item($i)
                $results += Get-RdoFoldersRecursive -Folder $sub -Depth ($Depth + 1) -StoreID $StoreID
            }
        }
    } catch {}
    return $results
}

function Get-RdoMailboxFolders {
    param([int]$StoreIndex)
    try {
        $store = $rdoSession.Stores.Item($StoreIndex)
        $storeID = ''
        try { $storeID = $store.StoreID } catch {}
        $root = $store.IPMRootFolder
        return @(Get-RdoFoldersRecursive -Folder $root -Depth 0 -StoreID $storeID)
    } catch {
        Write-Host "Get-RdoMailboxFolders failed: $_"
        return @()
    }
}

# ═══ Permission CRUD via Redemption ACL ═══════════════════════════════════════

function Get-RdoFolderPermissions {
    param($RdoFolder)
    $results = @()
    try {
        $acl = $RdoFolder.ACL
        for ($i = 1; $i -le $acl.Count; $i++) {
            $ace = $acl.Item($i)
            $name = $ace.Name
            if ($ace.IsDefault)   { $name = 'Default' }
            if ($ace.IsAnonymous) { $name = 'Anonymous' }
            $rights = [int]$ace.Rights
            $results += [PSCustomObject]@{
                User                = $name
                Rights              = $rights
                RightsHex           = '0x' + $rights.ToString('X4')
                PermissionLevelName = Get-FriendlyRightsName -Rights $rights
                Index               = $i
                IsDefault           = [bool]$ace.IsDefault
                IsAnonymous         = [bool]$ace.IsAnonymous
            }
        }
    } catch {
        Write-Host "Get-RdoFolderPermissions failed: $_"
    }
    return $results
}

function Set-RdoFolderPermission {
    param($RdoFolder, [string]$User, [int]$RightsValue)
    $acl = $RdoFolder.ACL

    # Check if user already has an ACE
    $existingAce = $null
    for ($i = 1; $i -le $acl.Count; $i++) {
        $ace = $acl.Item($i)
        if ($ace.Name -eq $User) { $existingAce = $ace; break }
    }

    if ($null -ne $existingAce) {
        $existingAce.Rights = $RightsValue
    } else {
        # Resolve user via GAL
        $addressEntry = $null
        try {
            $addressEntry = $rdoSession.AddressBook.GAL.ResolveName($User)
        } catch {}

        if ($null -ne $addressEntry) {
            $newAce = $acl.Add($addressEntry)
            $newAce.Rights = $RightsValue
        } else {
            throw "Could not resolve user '$User' in the Global Address List."
        }
    }
    $acl.Save()
}

function Set-RdoPermissionWithAncestors {
    param($RdoFolder, [string]$User, [int]$RightsValue)

    # 1. Set on target folder
    Set-RdoFolderPermission -RdoFolder $RdoFolder -User $User -RightsValue $RightsValue

    # 2. Walk parents, auto-grant ROLE_REVIEWER (0x0401) where user has no ACE
    $autoGranted = @()
    $current = $RdoFolder
    try { $current = $RdoFolder.Parent } catch { $current = $null }

    while ($null -ne $current) {
        # Stop at store root (parent is null)
        $parent = $null
        try { $parent = $current.Parent } catch {}
        if ($null -eq $parent) { break }

        $acl = $current.ACL
        $hasEntry = $false
        for ($i = 1; $i -le $acl.Count; $i++) {
            if ($acl.Item($i).Name -eq $User) { $hasEntry = $true; break }
        }

        if (-not $hasEntry) {
            $addressEntry = $null
            try { $addressEntry = $rdoSession.AddressBook.GAL.ResolveName($User) } catch {}
            if ($null -ne $addressEntry) {
                $newAce = $acl.Add($addressEntry)
                $newAce.Rights = 0x0401   # ROLE_REVIEWER = Can view
                $acl.Save()
                $folderName = ''
                try { $folderName = $current.Name } catch {}
                $autoGranted += @{ FolderName = $folderName; Level = 'Can view' }
            } else {
                $folderName = ''
                try { $folderName = $current.Name } catch {}
                $autoGranted += @{ FolderName = $folderName; Level = 'SKIPPED (resolve failed)' }
            }
        }

        $current = $parent
    }

    return [PSCustomObject]@{ Success = $true; AutoGranted = $autoGranted }
}

function Remove-RdoFolderPermission {
    param($RdoFolder, [string]$User)
    $acl = $RdoFolder.ACL
    for ($i = $acl.Count; $i -ge 1; $i--) {
        $ace = $acl.Item($i)
        if ($ace.Name -eq $User) {
            $acl.Remove($i)
            break
        }
    }
    $acl.Save()
}

# ═══ Active Directory search ══════════════════════════════════════════════════
$script:ADCache = @{}

function Search-ADUsers {
    param([string]$Query)
    if ([string]::IsNullOrWhiteSpace($Query) -or $Query.Length -lt 5) { return @() }

    $cacheKey = $Query.ToLower()
    if ($script:ADCache.ContainsKey($cacheKey)) {
        $entry = $script:ADCache[$cacheKey]
        if ([datetime]::Now -lt $entry.Expires) { return $entry.Results }
        $script:ADCache.Remove($cacheKey)
    }

    try {
        Add-Type -AssemblyName System.DirectoryServices -ErrorAction Stop
        $searcher = New-Object System.DirectoryServices.DirectorySearcher
        $safe = $Query -replace '[\\*()\x00]',''

        if ($Query -match '@') {
            $parts     = $Query -split '@', 2
            $safePart0 = $parts[0] -replace '[\\*()\x00]',''
            $safePart1 = if ($parts.Count -gt 1) { $parts[1] -replace '[\\*()\x00]','' } else { '' }
            if ([string]::IsNullOrEmpty($safePart0) -or [string]::IsNullOrEmpty($safePart1)) { return @() }
            $sam   = "OG-$safePart1-$safePart0"
            $ldap  = "(&(objectCategory=group)(sAMAccountName=$sam))"
        } elseif ($Query -like 'OG-*') {
            $ldap  = "(&(objectCategory=group)(sAMAccountName=$safe*))"
        } else {
            $ldap  = "(&(objectCategory=group)(displayName=OG-*$safe*))"
        }

        $searcher.Filter = $ldap
        $searcher.PropertiesToLoad.AddRange(@('displayName','mail','sAMAccountName')) | Out-Null
        $searcher.SizeLimit       = 20
        $searcher.ServerTimeLimit = [System.TimeSpan]::FromSeconds(4)
        $searcher.CacheResults    = $false
        $results = $searcher.FindAll()
        $out = @()
        foreach ($r in $results) {
            $dn   = if ($r.Properties['displayName'].Count -gt 0)    { $r.Properties['displayName'][0] }    else { '' }
            $mail = if ($r.Properties['mail'].Count -gt 0)           { $r.Properties['mail'][0] }           else { '' }
            $sam  = if ($r.Properties['sAMAccountName'].Count -gt 0) { $r.Properties['sAMAccountName'][0] } else { '' }
            if ($dn -or $mail -or $sam) {
                $out += [PSCustomObject]@{ DisplayName = $dn; Mail = $mail; SamAccountName = $sam }
            }
        }
        $script:ADCache[$cacheKey] = @{ Results = $out; Expires = [datetime]::Now.AddSeconds(60) }
        return $out
    } catch {
        return @()
    }
}

# ═══ XAML layout ══════════════════════════════════════════════════════════════
$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Redemption Permissions Test Tool" Width="980" Height="680"
        WindowStartupLocation="CenterScreen"
        Background="#1E1E2E">
  <Window.Resources>
    <SolidColorBrush x:Key="BgBrush"            Color="#1E1E2E"/>
    <SolidColorBrush x:Key="SurfaceBrush"        Color="#2A2A3C"/>
    <SolidColorBrush x:Key="SurfaceHoverBrush"   Color="#33334A"/>
    <SolidColorBrush x:Key="BorderBrush"         Color="#3E3E55"/>
    <SolidColorBrush x:Key="AccentBrush"         Color="#7C6FE0"/>
    <SolidColorBrush x:Key="AccentSubtleBrush"   Color="#2E2B4A"/>
    <SolidColorBrush x:Key="TextPrimaryBrush"    Color="#EAEAEF"/>
    <SolidColorBrush x:Key="TextSecondaryBrush"  Color="#9898A8"/>
  </Window.Resources>

  <Grid Margin="16">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="4"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="10"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="6"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- Header -->
    <TextBlock Grid.Row="0" Text="Redemption Permissions Test Tool" FontSize="18" FontWeight="Bold"
               Foreground="{StaticResource TextPrimaryBrush}"/>
    <TextBlock Grid.Row="2" Text="Compare values against Outlook > Folder Properties > Permissions to validate mapping"
               FontSize="11" FontStyle="Italic" Foreground="{StaticResource TextSecondaryBrush}"/>

    <!-- Main 3-column layout -->
    <Grid Grid.Row="4">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="220"/>
        <ColumnDefinition Width="12"/>
        <ColumnDefinition Width="240"/>
        <ColumnDefinition Width="12"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>

      <!-- Col 0: Mailboxes -->
      <Grid Grid.Column="0">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="6"/>
          <RowDefinition Height="*"/>
          <RowDefinition Height="10"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" Text="MAILBOXES" FontWeight="SemiBold" FontSize="11"
                   Foreground="{StaticResource TextSecondaryBrush}" Margin="2,0,0,0"/>

        <Grid Grid.Row="2">
          <ListBox Name="LbMailboxes" Background="{StaticResource SurfaceBrush}"
                   BorderBrush="{StaticResource BorderBrush}" BorderThickness="1"
                   Foreground="{StaticResource TextPrimaryBrush}">
            <ListBox.ItemContainerStyle>
              <Style TargetType="ListBoxItem">
                <Setter Property="Padding" Value="10,6"/>
                <Setter Property="Margin"  Value="0"/>
                <Setter Property="Background" Value="Transparent"/>
                <Setter Property="Foreground" Value="{StaticResource TextPrimaryBrush}"/>
                <Style.Triggers>
                  <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="{StaticResource AccentSubtleBrush}"/>
                  </Trigger>
                  <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{StaticResource SurfaceHoverBrush}"/>
                  </Trigger>
                </Style.Triggers>
              </Style>
            </ListBox.ItemContainerStyle>
            <ListBox.ItemTemplate>
              <DataTemplate>
                <StackPanel>
                  <TextBlock Text="{Binding Name}" FontWeight="SemiBold" FontSize="13"
                             Foreground="{StaticResource TextPrimaryBrush}" TextTrimming="CharacterEllipsis"/>
                  <TextBlock Text="{Binding SmtpAddress}" FontSize="11"
                             Foreground="{StaticResource TextSecondaryBrush}" TextTrimming="CharacterEllipsis"/>
                </StackPanel>
              </DataTemplate>
            </ListBox.ItemTemplate>
          </ListBox>
          <TextBlock Name="TxtMailboxEmptyHint"
                     Text="Open Outlook, then click Refresh mailboxes."
                     TextWrapping="Wrap" Foreground="{StaticResource TextSecondaryBrush}"
                     FontSize="12" FontStyle="Italic"
                     HorizontalAlignment="Center" VerticalAlignment="Center"
                     Margin="12" Opacity="0.7" IsHitTestVisible="False"/>
        </Grid>

        <Button Grid.Row="4" Name="BtnRefreshMailboxes" Content="Refresh mailboxes"
                HorizontalAlignment="Stretch" Padding="8,6"
                Background="{StaticResource SurfaceBrush}" Foreground="{StaticResource TextPrimaryBrush}"
                BorderBrush="{StaticResource BorderBrush}" BorderThickness="1"/>
      </Grid>

      <!-- Col 2: Folders -->
      <Grid Grid.Column="2">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="6"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" Text="FOLDERS" FontWeight="SemiBold" FontSize="11"
                   Foreground="{StaticResource TextSecondaryBrush}" Margin="2,0,0,0"/>

        <Grid Grid.Row="2">
          <ListBox Name="LbFolders" Background="{StaticResource SurfaceBrush}"
                   BorderBrush="{StaticResource BorderBrush}" BorderThickness="1"
                   Foreground="{StaticResource TextPrimaryBrush}">
            <ListBox.ItemContainerStyle>
              <Style TargetType="ListBoxItem">
                <Setter Property="Padding" Value="0"/>
                <Setter Property="Margin"  Value="0"/>
                <Setter Property="MinHeight" Value="20"/>
                <Setter Property="Background" Value="Transparent"/>
                <Setter Property="Foreground" Value="{StaticResource TextPrimaryBrush}"/>
                <Style.Triggers>
                  <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="{StaticResource AccentSubtleBrush}"/>
                  </Trigger>
                  <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{StaticResource SurfaceHoverBrush}"/>
                  </Trigger>
                </Style.Triggers>
              </Style>
            </ListBox.ItemContainerStyle>
            <ListBox.ItemTemplate>
              <DataTemplate>
                <StackPanel Orientation="Horizontal" Margin="{Binding IndentPad}">
                  <TextBlock Text="{Binding Icon}" FontSize="12" FontFamily="Segoe UI Symbol"
                             Margin="0,0,5,0" Width="16"
                             Foreground="{StaticResource TextSecondaryBrush}"
                             VerticalAlignment="Center" TextAlignment="Center"/>
                  <TextBlock Text="{Binding Name}" TextTrimming="CharacterEllipsis" FontSize="11"
                             Foreground="{StaticResource TextPrimaryBrush}" VerticalAlignment="Center"/>
                </StackPanel>
              </DataTemplate>
            </ListBox.ItemTemplate>
          </ListBox>
          <TextBlock Name="TxtFoldersHint"
                     Text="Select a mailbox to see its folders"
                     TextWrapping="Wrap" Foreground="{StaticResource TextSecondaryBrush}"
                     FontSize="12" FontStyle="Italic"
                     HorizontalAlignment="Center" VerticalAlignment="Center"
                     Margin="12" Opacity="0.7" IsHitTestVisible="False"/>
        </Grid>
      </Grid>

      <!-- Col 4: Permissions -->
      <Grid Grid.Column="4" Name="PanelPermRight">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="6"/>
          <RowDefinition Height="*" MinHeight="80"/>
          <RowDefinition Height="10"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="6"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="6"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="6"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" Text="PERMISSIONS" FontWeight="SemiBold" FontSize="11"
                   Foreground="{StaticResource TextSecondaryBrush}"/>

        <Grid Grid.Row="2">
          <DataGrid Name="DgCurrentPerms"
                    AutoGenerateColumns="False" IsReadOnly="True"
                    HeadersVisibility="Column"
                    CanUserReorderColumns="False" CanUserSortColumns="False" CanUserResizeRows="False"
                    SelectionMode="Single" GridLinesVisibility="None"
                    Background="{StaticResource SurfaceBrush}"
                    Foreground="{StaticResource TextPrimaryBrush}"
                    BorderBrush="{StaticResource BorderBrush}" BorderThickness="1"
                    RowBackground="{StaticResource SurfaceBrush}"
                    AlternatingRowBackground="{StaticResource BgBrush}"
                    HorizontalScrollBarVisibility="Disabled"
                    ColumnHeaderHeight="30">
            <DataGrid.ColumnHeaderStyle>
              <Style TargetType="DataGridColumnHeader">
                <Setter Property="Background" Value="{StaticResource BgBrush}"/>
                <Setter Property="Foreground" Value="{StaticResource TextSecondaryBrush}"/>
                <Setter Property="FontSize" Value="11"/>
                <Setter Property="FontWeight" Value="SemiBold"/>
                <Setter Property="Padding" Value="10,6"/>
                <Setter Property="BorderThickness" Value="0"/>
              </Style>
            </DataGrid.ColumnHeaderStyle>
            <DataGrid.CellStyle>
              <Style TargetType="DataGridCell">
                <Setter Property="BorderThickness" Value="0"/>
                <Setter Property="Padding" Value="10,6"/>
                <Setter Property="Foreground" Value="{StaticResource TextPrimaryBrush}"/>
                <Style.Triggers>
                  <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="{StaticResource AccentSubtleBrush}"/>
                    <Setter Property="Foreground" Value="{StaticResource TextPrimaryBrush}"/>
                  </Trigger>
                </Style.Triggers>
              </Style>
            </DataGrid.CellStyle>
            <DataGrid.RowStyle>
              <Style TargetType="DataGridRow">
                <Style.Triggers>
                  <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="{StaticResource AccentSubtleBrush}"/>
                  </Trigger>
                </Style.Triggers>
              </Style>
            </DataGrid.RowStyle>
            <DataGrid.Columns>
              <DataGridTextColumn Header="Person or group" Binding="{Binding User}" Width="*"/>
              <DataGridTextColumn Header="Access level"    Binding="{Binding PermissionLevelName}" Width="150"/>
              <DataGridTextColumn Header="Rights (hex)"    Binding="{Binding RightsHex}" Width="90"/>
            </DataGrid.Columns>
          </DataGrid>
          <TextBlock Name="TxtFolderHint"
                     Text="Select a folder to see who has access"
                     TextWrapping="Wrap" Foreground="{StaticResource TextSecondaryBrush}"
                     FontSize="12" FontStyle="Italic"
                     HorizontalAlignment="Center" VerticalAlignment="Center"
                     Margin="12" Opacity="0.7" IsHitTestVisible="False"/>
        </Grid>

        <Grid Grid.Row="4" Height="180">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>
          <TextBox Grid.Row="0" Name="TxtAddUser" Padding="8,6"
                   Background="{StaticResource SurfaceBrush}"
                   Foreground="{StaticResource TextPrimaryBrush}"
                   BorderBrush="{StaticResource BorderBrush}" BorderThickness="1"
                   ToolTip="Search Active Directory or type an email address directly"/>
          <ListBox Grid.Row="1" Name="LbAdResults" Margin="0,2,0,0" BorderThickness="1"
                   Background="{StaticResource SurfaceBrush}"
                   BorderBrush="{StaticResource BorderBrush}"
                   ScrollViewer.VerticalScrollBarVisibility="Auto">
            <ListBox.ItemContainerStyle>
              <Style TargetType="ListBoxItem">
                <Setter Property="Padding" Value="8,2"/>
                <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
                <Setter Property="Background" Value="Transparent"/>
                <Setter Property="Foreground" Value="{StaticResource TextPrimaryBrush}"/>
                <Style.Triggers>
                  <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="{StaticResource AccentSubtleBrush}"/>
                  </Trigger>
                </Style.Triggers>
              </Style>
            </ListBox.ItemContainerStyle>
            <ListBox.ItemTemplate>
              <DataTemplate>
                <StackPanel Margin="0,1">
                  <TextBlock Text="{Binding DisplayName}" FontWeight="SemiBold" FontSize="12"
                             Foreground="{StaticResource TextPrimaryBrush}"/>
                  <TextBlock Text="{Binding Mail}" FontSize="11"
                             Foreground="{StaticResource TextSecondaryBrush}"/>
                </StackPanel>
              </DataTemplate>
            </ListBox.ItemTemplate>
          </ListBox>
        </Grid>

        <ComboBox Grid.Row="6" Name="CbPermLevel" Padding="8,6"
                  Background="{StaticResource SurfaceBrush}"
                  Foreground="{StaticResource TextPrimaryBrush}"
                  BorderBrush="{StaticResource BorderBrush}" BorderThickness="1"
                  ToolTip="Choose what this person is allowed to do"/>

        <StackPanel Grid.Row="8" Orientation="Horizontal">
          <Button Name="BtnSavePerm" Content="Save" Margin="0,0,8,0" Width="90" Padding="8,6"
                  Background="{StaticResource AccentBrush}" Foreground="White"
                  BorderThickness="0"/>
          <Button Name="BtnRemovePerm" Content="Remove" Width="90" Padding="8,6"
                  Background="{StaticResource SurfaceBrush}" Foreground="{StaticResource TextPrimaryBrush}"
                  BorderBrush="{StaticResource BorderBrush}" BorderThickness="1"/>
        </StackPanel>

        <TextBlock Grid.Row="10" Name="TxtPermStatus" FontSize="12"
                   Foreground="{StaticResource TextSecondaryBrush}" TextWrapping="Wrap"/>
      </Grid>
    </Grid>

    <!-- Status bar -->
    <Border Grid.Row="6" Name="StatusBar" Background="{StaticResource SurfaceBrush}"
            CornerRadius="4" Padding="10,6" Visibility="Collapsed">
      <TextBlock Name="TxtStatus" FontSize="11" Foreground="{StaticResource TextPrimaryBrush}"/>
    </Border>
  </Grid>
</Window>
'@

# ═══ Parse XAML and get control references ════════════════════════════════════
$reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)

$lbMailboxes    = $window.FindName('LbMailboxes')
$lbFolders      = $window.FindName('LbFolders')
$dgPerms        = $window.FindName('DgCurrentPerms')
$txtAddUser     = $window.FindName('TxtAddUser')
$lbAdResults    = $window.FindName('LbAdResults')
$cbPermLevel    = $window.FindName('CbPermLevel')
$btnSave        = $window.FindName('BtnSavePerm')
$btnRemove      = $window.FindName('BtnRemovePerm')
$txtPermStatus  = $window.FindName('TxtPermStatus')
$txtFolderHint  = $window.FindName('TxtFolderHint')
$txtFoldersHint = $window.FindName('TxtFoldersHint')
$panelPermRight = $window.FindName('PanelPermRight')
$btnRefresh     = $window.FindName('BtnRefreshMailboxes')
$txtMbxHint     = $window.FindName('TxtMailboxEmptyHint')
$statusBar      = $window.FindName('StatusBar')
$txtStatus      = $window.FindName('TxtStatus')

# Populate permission level ComboBox
foreach ($key in $FriendlyToRights.Keys) { $cbPermLevel.Items.Add($key) | Out-Null }
$cbPermLevel.SelectedIndex = 1   # default: "Can view"

# ═══ Shared state ═════════════════════════════════════════════════════════════
$script:selectedFolder  = $null
$script:removePending   = $null
$script:adLastQuery     = ''
$script:adTimer         = $null

# ═══ Helper functions ═════════════════════════════════════════════════════════

function Set-Status {
    param([string]$Msg)
    $txtStatus.Text = $Msg
    $statusBar.Visibility = 'Visible'
    if ($null -ne $script:statusHideTimer) { $script:statusHideTimer.Stop() }
    $script:statusHideTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:statusHideTimer.Interval = [System.TimeSpan]::FromSeconds(4)
    $script:statusHideTimer.Add_Tick({
        $script:statusHideTimer.Stop()
        $statusBar.Visibility = 'Collapsed'
    })
    $script:statusHideTimer.Start()
}

function Set-PermStatus {
    param([string]$Msg, [string]$Colour = '')
    $txtPermStatus.Text = $Msg
    if ($Colour) {
        $txtPermStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($Colour)
    } else {
        $txtPermStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#9898A8')
    }
}

function Set-PermRightEnabled {
    param([bool]$Enabled)
    $panelPermRight.IsEnabled = $Enabled
    $panelPermRight.Opacity   = if ($Enabled) { 1.0 } else { 0.45 }
}

function Refresh-PermGrid {
    if ($null -eq $script:selectedFolder) { return }
    try {
        $rows = @(Get-RdoFolderPermissions -RdoFolder $script:selectedFolder.RdoFolder)
        $dgPerms.ItemsSource = $null
        if ($rows.Count -gt 0) { $dgPerms.ItemsSource = $rows }
        Set-PermStatus "Showing $($rows.Count) permission entries" ''
    } catch {
        Set-PermStatus "Could not read permissions: $_" '#E74C3C'
    }
}

# Initial state
Set-PermRightEnabled $false

# ═══ Load mailboxes ═══════════════════════════════════════════════════════════

function Load-Mailboxes {
    $accounts = @(Get-RdoAccounts)
    $lbMailboxes.SelectedIndex = -1
    $lbMailboxes.ItemsSource = $null
    $lbMailboxes.ItemsSource = $accounts
    $isEmpty = $accounts.Count -eq 0
    $txtMbxHint.Visibility = if ($isEmpty) { 'Visible' } else { 'Collapsed' }
    if ($isEmpty) {
        Set-Status 'No mailboxes found. Open Outlook and click Refresh.'
    } else {
        Set-Status "Loaded $($accounts.Count) mailbox store(s)"
    }
}

Load-Mailboxes

# ═══ Event wiring ═════════════════════════════════════════════════════════════

# Mailbox selection -> populate folder list
$lbMailboxes.Add_SelectionChanged({
    $sel = $lbMailboxes.SelectedItem
    if ($null -eq $sel) { return }
    $lbFolders.ItemsSource   = $null
    $dgPerms.ItemsSource     = $null
    $script:selectedFolder   = $null
    $txtFolderHint.Visibility  = 'Visible'
    $txtFoldersHint.Text       = 'Loading folders...'
    $txtFoldersHint.Visibility = 'Visible'
    Set-PermRightEnabled $false
    Set-PermStatus ''
    try {
        $raw = @(Get-RdoMailboxFolders -StoreIndex $sel.StoreIndex)
        if ($raw.Count -eq 0) {
            $txtFoldersHint.Text = 'No folders found'
        } else {
            $txtFoldersHint.Visibility = 'Collapsed'
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
                    RdoFolder  = $f.RdoFolder
                })
            }
            $lbFolders.ItemsSource = $items
        }
        Set-Status "Loaded $($raw.Count) folders for $($sel.Name)"
    } catch {
        $txtFoldersHint.Text       = 'Error loading folders'
        $txtFoldersHint.Visibility = 'Visible'
        Set-PermStatus "Could not load folders: $_" '#E74C3C'
    }
})

# Folder selection -> show permissions
$lbFolders.Add_SelectionChanged({
    $sel = $lbFolders.SelectedItem
    if ($null -eq $sel) { return }
    $script:selectedFolder       = $sel
    $txtFolderHint.Visibility    = 'Collapsed'
    $dgPerms.ItemsSource         = $null
    $script:removePending        = $null
    Set-PermStatus ''
    Set-PermRightEnabled $true
    try {
        Refresh-PermGrid
        Set-Status "Permissions for $($sel.Name)"
    } catch {
        Set-PermStatus "Error reading permissions: $_" '#E74C3C'
    }
})

# DataGrid selection -> bit inspector in status + clear remove pending
$dgPerms.Add_SelectionChanged({
    $script:removePending = $null
    $row = $dgPerms.SelectedItem
    if ($null -ne $row) {
        $decomp = Get-RightsDecomposition -Rights $row.Rights
        Set-PermStatus "$($row.RightsHex): $decomp" ''
    } else {
        Set-PermStatus '' ''
    }
})

# AD search: debounced via DispatcherTimer
function Invoke-AdSearchTick {
    if ($null -ne $script:adTimer) { $script:adTimer.Stop() }
    $script:adTimer = $null
    try {
        $hits = Search-ADUsers -Query $script:adLastQuery
        $lbAdResults.ItemsSource = $null
        if ($hits.Count -gt 0) {
            $lbAdResults.ItemsSource = $hits
        } else {
            Set-PermStatus 'No AD matches found' ''
        }
    } catch {}
}

function Start-AdSearchDebounce {
    param([string]$Query)
    if ($null -ne $script:adTimer) { $script:adTimer.Stop(); $script:adTimer = $null }
    if ($Query.Length -lt 5) { $lbAdResults.ItemsSource = $null; return }
    $script:adLastQuery = $Query
    $t = New-Object System.Windows.Threading.DispatcherTimer
    $t.Interval = [System.TimeSpan]::FromMilliseconds(350)
    $t.Add_Tick({ Invoke-AdSearchTick })
    $script:adTimer = $t
    $t.Start()
}

$txtAddUser.Add_TextChanged({
    Start-AdSearchDebounce -Query $txtAddUser.Text.Trim()
})

# AD result selected -> fill search box
$lbAdResults.Add_SelectionChanged({
    $hit = $lbAdResults.SelectedItem
    if ($null -eq $hit) { return }
    $value = if ($hit.Mail) { $hit.Mail } else { $hit.DisplayName }
    $txtAddUser.Text       = $value
    $txtAddUser.CaretIndex = $value.Length
    $lbAdResults.ItemsSource   = $null
    $lbAdResults.SelectedIndex = -1
    Set-PermStatus ''
})

# Save permission
$btnSave.Add_Click({
    if ($null -eq $script:selectedFolder) {
        Set-PermStatus 'Select a folder first' '#E74C3C'; return
    }
    $user  = $txtAddUser.Text.Trim()
    $level = $cbPermLevel.SelectedItem
    if ([string]::IsNullOrWhiteSpace($user)) {
        Set-PermStatus 'Enter a user or group name' '#E74C3C'; return
    }
    if ($null -eq $level) {
        Set-PermStatus 'Select a permission level' '#E74C3C'; return
    }
    $rightsValue = $FriendlyToRights[$level]
    try {
        $result = Set-RdoPermissionWithAncestors -RdoFolder $script:selectedFolder.RdoFolder `
                                                  -User $user -RightsValue $rightsValue
        Refresh-PermGrid
        $msg = "Saved: $user = $level"
        if ($result.AutoGranted.Count -gt 0) {
            $folderNames = ($result.AutoGranted | ForEach-Object { $_.FolderName }) -join ', '
            $msg += " | Auto-granted Can view on: $folderNames"
        }
        Set-PermStatus $msg '#27AE60'
        Set-Status $msg
    } catch {
        Set-PermStatus "Could not save: $_" '#E74C3C'
        Set-Status "Could not save: $_"
    }
})

# Remove permission (two-click confirm)
$btnRemove.Add_Click({
    if ($null -eq $script:selectedFolder) {
        $script:removePending = $null
        Set-PermStatus 'Select a folder first' '#E74C3C'; return
    }
    $row = $dgPerms.SelectedItem
    if ($null -eq $row) {
        $script:removePending = $null
        Set-PermStatus 'Select a person to remove' '#E74C3C'; return
    }
    if ($row.User -eq 'Default' -or $row.User -eq 'Anonymous') {
        $script:removePending = $null
        Set-PermStatus "Cannot remove $($row.User)" '#E74C3C'; return
    }
    if ($script:removePending -ne $row.User) {
        $script:removePending = $row.User
        Set-PermStatus "Click Remove again to confirm removing $($row.User)" ''; return
    }
    # Second click confirmed
    $script:removePending = $null
    try {
        Remove-RdoFolderPermission -RdoFolder $script:selectedFolder.RdoFolder -User $row.User
        Refresh-PermGrid
        Set-PermStatus "Removed $($row.User)" '#27AE60'
        Set-Status "Removed $($row.User)"
    } catch {
        Set-PermStatus "Could not remove: $_" '#E74C3C'
        Set-Status "Could not remove: $_"
    }
})

# Refresh mailboxes
$btnRefresh.Add_Click({
    $lbFolders.ItemsSource   = $null
    $dgPerms.ItemsSource     = $null
    $script:selectedFolder   = $null
    $txtFolderHint.Visibility  = 'Visible'
    $txtFoldersHint.Text       = 'Select a mailbox to see its folders'
    $txtFoldersHint.Visibility = 'Visible'
    Set-PermRightEnabled $false
    Set-PermStatus ''
    Load-Mailboxes
})

# ═══ Show window ══════════════════════════════════════════════════════════════
$window.ShowDialog() | Out-Null

# Cleanup
try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($rdoSession) | Out-Null } catch {}
