<#
  PermissionsManager.psm1
  Helpers to enumerate Outlook accounts, mailbox folders, and folder permissions
  using the Outlook COM object model (MAPI). Includes Active Directory user search
  via System.DirectoryServices (built-in .NET — no external modules required).
#>

# ─── OlPermission enum values used by Outlook COM ────────────────────────────
# These correspond to the OlPermissions enumeration in the Outlook type library.
$script:OlPermissionMap = [ordered]@{
    0  = 'No access'
    1  = 'Can view'
    2  = 'Can create items'
    4  = 'Can create, edit & delete'
    7  = 'Full access'
    8  = 'Owner'
    9  = 'Can create & edit own items'
    5  = 'Can create & edit all items'
    3  = 'Can view & create'
}

# Friendly → OlPermissions numeric (for setting)
$script:FriendlyToOlPerm = [ordered]@{
    'No access'                    = 0
    'Can view'                     = 1
    'Can create items'             = 2
    'Can create, edit & delete'    = 4
    'Full access'                  = 7
    'Owner'                        = 8
    'Can create & edit own items'  = 9
    'Can create & edit all items'  = 5
    'Can view & create'            = 3
}

$script:OutlookNSFactory = $null
$script:ADSearchFactory  = $null

# ─── AD search cache ──────────────────────────────────────────────────────────
# Keyed by query string (lowercase). Each entry: @{ Results=@(); Expires=[datetime] }
$script:ADCache = @{}

function script:Get-OutlookNS {
    if ($script:OutlookNSFactory) { return & $script:OutlookNSFactory }
    $ol = New-Object -ComObject Outlook.Application
    return $ol.GetNameSpace('MAPI')
}

function Set-OutlookNSFactory {
    <# Inject a scriptblock that returns a mock MAPI namespace (for offline testing). #>
    param([scriptblock]$Factory)
    $script:OutlookNSFactory = $Factory
}

function Set-ADSearchFactory {
    <# Inject a scriptblock that returns mock AD search results (for offline testing). #>
    param([scriptblock]$Factory)
    $script:ADSearchFactory = $Factory
}

# ─── Account enumeration (reused by both Signatures and Permissions tabs) ─────

function Get-SignedInAccounts {
    <#
      Returns a list of accounts from the current Outlook profile.
      Each item: [PSCustomObject]@{ Name; SmtpAddress }
    #>
    try {
        $ns = Get-OutlookNS
        $accounts = @()
        if ($ns -and $ns.Accounts) {
            for ($i = 1; $i -le $ns.Accounts.Count; $i++) {
                $acc  = $ns.Accounts.Item($i)
                $smtp = ''
                try { $smtp = $acc.SmtpAddress } catch {}
                $accounts += [PSCustomObject]@{ Name = $acc.DisplayName; SmtpAddress = $smtp }
            }
        }
        return $accounts
    } catch {
        Write-Verbose "Outlook accounts could not be enumerated: $_"
        return @()
    }
}

# ─── Folder enumeration ───────────────────────────────────────────────────────

# System/hidden folder names to suppress — covers Exchange internals, add-in folders,
# localized German names, and other folders users should not set permissions on.
$script:HiddenFolderNames = [System.Collections.Generic.HashSet[string]]([System.StringComparer]::OrdinalIgnoreCase)
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
) | ForEach-Object { $script:HiddenFolderNames.Add($_) | Out-Null }

function script:Get-FoldersRecursive {
    param($Folder, [int]$Depth = 0)
    $results = @()

    $folderName = $Folder.Name

    # Depth 0 is the store root — include it so the mailbox itself appears as a header
    # but skip any root whose name is in the hidden list
    if ($Depth -gt 0 -and $script:HiddenFolderNames.Contains($folderName)) {
        return $results   # skip this folder and all its children
    }

    $results += [PSCustomObject]@{
        Name        = $folderName
        FolderPath  = $Folder.FolderPath
        EntryID     = $Folder.EntryID
        StoreID     = $Folder.StoreID
        Depth       = $Depth
    }
    try {
        if ($Folder.Folders -and $Folder.Folders.Count -gt 0) {
            for ($i = 1; $i -le $Folder.Folders.Count; $i++) {
                $sub = $Folder.Folders.Item($i)
                $results += Get-FoldersRecursive -Folder $sub -Depth ($Depth + 1)
            }
        }
    } catch {}
    return $results
}

function Get-MailboxFolders {
    <#
      Returns a flat list of folders for the mailbox matching SmtpAddress.
      Each item: [PSCustomObject]@{ Name; FolderPath; EntryID; StoreID; Depth }
      Depth=0 is the root; subfolders have Depth=1, 2, etc.
    #>
    param([string]$SmtpAddress)
    try {
        $ns = Get-OutlookNS

        # Single pass: collect display names and try DeliveryStore (preferred — exact match)
        $accDisplayNames = @()
        for ($j = 1; $j -le $ns.Accounts.Count; $j++) {
            $a = $ns.Accounts.Item($j)
            $aSmtp = ''; try { $aSmtp = $a.SmtpAddress } catch {}
            if ($aSmtp -eq $SmtpAddress) {
                $accDisplayNames += $a.DisplayName
                try {
                    $delivStore = $a.DeliveryStore
                    if ($null -ne $delivStore) {
                        return Get-FoldersRecursive -Folder $delivStore.GetRootFolder() -Depth 0
                    }
                } catch {}
            }
        }

        # Fallback: walk every Store and match by display name
        for ($i = 1; $i -le $ns.Stores.Count; $i++) {
            $store = $ns.Stores.Item($i)
            $storeName = ''; try { $storeName = $store.DisplayName } catch {}
            $isMatch = $false
            foreach ($dn in $accDisplayNames) {
                if ($storeName -like "*$dn*" -or $dn -like "*$storeName*") {
                    $isMatch = $true; break
                }
            }
            # Also match if the store name directly contains the SMTP address
            if (-not $isMatch -and $storeName -like "*$SmtpAddress*") { $isMatch = $true }
            # If only one store exists, just use it
            if (-not $isMatch -and $ns.Stores.Count -eq 1) { $isMatch = $true }

            if ($isMatch) {
                try {
                    $root = $store.GetRootFolder()
                    return Get-FoldersRecursive -Folder $root -Depth 0
                } catch {}
            }
        }

        # Tier 2b: FilePath match for non-standard OST/PST directories
        # Handles custom store locations (e.g. AppData\Local\AA_OST, AppData\Local\AA_PST)
        $ostDir    = Join-Path $env:LOCALAPPDATA 'AA_OST'
        $pstDir    = Join-Path $env:LOCALAPPDATA 'AA_PST'
        $smtpLocal = ($SmtpAddress -split '@')[0]
        for ($i = 1; $i -le $ns.Stores.Count; $i++) {
            $store = $ns.Stores.Item($i)
            try {
                $fp = $store.FilePath
                if ([string]::IsNullOrEmpty($fp)) { continue }
                $inAaDir = ($fp.StartsWith($ostDir, [System.StringComparison]::OrdinalIgnoreCase) -or
                            $fp.StartsWith($pstDir, [System.StringComparison]::OrdinalIgnoreCase))
                if (-not $inAaDir) { continue }
                # Secondary guard: filename or store display name loosely matches SMTP local part
                $fn = [System.IO.Path]::GetFileNameWithoutExtension($fp)
                $sn = ''; try { $sn = $store.DisplayName } catch {}
                if ($fn -like "*$smtpLocal*" -or $sn -like "*$smtpLocal*" -or
                    $fn -like "*$SmtpAddress*" -or $sn -like "*$SmtpAddress*") {
                    try {
                        $root = $store.GetRootFolder()
                        return Get-FoldersRecursive -Folder $root -Depth 0
                    } catch {}
                }
            } catch {}
        }

        # Last resort: first store
        $firstRoot = $ns.Stores.Item(1).GetRootFolder()
        return Get-FoldersRecursive -Folder $firstRoot -Depth 0
    } catch {
        Write-Verbose "Get-MailboxFolders failed: $_"
        return @()
    }
}

# ─── Read permissions ─────────────────────────────────────────────────────────

function Get-FolderPermissions {
    <#
      Returns the permission entries for a folder identified by EntryID + StoreID.
      Each item: [PSCustomObject]@{ User; PermissionLevel; PermissionLevelName }
    #>
    param(
        [Parameter(Mandatory=$true)][string]$EntryID,
        [Parameter(Mandatory=$true)][string]$StoreID
    )
    try {
        $ns     = Get-OutlookNS
        $folder = $ns.GetFolderFromID($EntryID, $StoreID)
        $perms  = $folder.FolderPermissions
        $results = @()
        for ($i = 1; $i -le $perms.Count; $i++) {
            $p     = $perms.Item($i)
            $level = $p.OlFolderPermission
            $name  = if ($script:OlPermissionMap.Contains([int]$level)) {
                         $script:OlPermissionMap[[int]$level]
                     } else { "Level $level" }
            $results += [PSCustomObject]@{
                User               = $p.UserName
                PermissionLevel    = [int]$level
                PermissionLevelName = $name
            }
        }
        return $results
    } catch {
        Write-Verbose "Get-FolderPermissions failed: $_"
        return @()
    }
}

# ─── Write permissions ────────────────────────────────────────────────────────

function Set-FolderPermission {
    <#
      Adds or updates a user's permission on a folder.
      $Level should be one of the friendly strings from $FriendlyToOlPerm keys,
      or a numeric OlPermissions value (0-9).
    #>
    param(
        [Parameter(Mandatory=$true)][string]$EntryID,
        [Parameter(Mandatory=$true)][string]$StoreID,
        [Parameter(Mandatory=$true)][string]$User,
        [Parameter(Mandatory=$true)]$Level
    )
    $levelNum = $null
    if ($Level -is [int]) {
        $levelNum = $Level
    } elseif ($script:FriendlyToOlPerm.Contains($Level)) {
        $levelNum = $script:FriendlyToOlPerm[$Level]
    } else {
        throw "Unknown permission level: $Level"
    }

    $ns     = Get-OutlookNS
    $folder = $ns.GetFolderFromID($EntryID, $StoreID)
    $perms  = $folder.FolderPermissions

    # Check if user already exists
    $existing = $null
    for ($i = 1; $i -le $perms.Count; $i++) {
        if ($perms.Item($i).UserName -eq $User) {
            $existing = $perms.Item($i)
            break
        }
    }

    if ($null -ne $existing) {
        $existing.OlFolderPermission = $levelNum
    } else {
        $newPerm = $perms.Add($User)
        $newPerm.OlFolderPermission = $levelNum
    }
    $perms.Save()
}

function Set-FolderPermissionWithAncestors {
    <#
      Sets a user's permission on the target folder AND auto-grants "Can view"
      on every ancestor folder where the user has NO existing permission entry.
      This ensures the user can navigate to the subfolder in Outlook.

      Returns [PSCustomObject]@{ Success=$true; AutoGranted=@( @{FolderName; Level} ) }
      AutoGranted lists each parent folder that was automatically given "Can view".
      If user already has ANY permission on a parent (even "No access"), that parent
      is left alone — we don't override explicit admin decisions.
    #>
    param(
        [Parameter(Mandatory=$true)][string]$EntryID,
        [Parameter(Mandatory=$true)][string]$StoreID,
        [Parameter(Mandatory=$true)][string]$User,
        [Parameter(Mandatory=$true)]$Level
    )

    # 1. Set permission on the target folder
    Set-FolderPermission -EntryID $EntryID -StoreID $StoreID -User $User -Level $Level

    # 2. Walk parents upward, auto-granting "Can view" where needed
    $autoGranted = @()
    $ns     = Get-OutlookNS
    $folder = $ns.GetFolderFromID($EntryID, $StoreID)

    $current = $folder.Parent
    while ($null -ne $current) {
        # Stop at the store root (Depth 0): a folder whose own Parent is $null
        # is the mailbox root — we never auto-grant on that level.
        if ($null -eq $current.Parent) { break }

        $perms = $current.FolderPermissions
        $hasEntry = $false
        for ($i = 1; $i -le $perms.Count; $i++) {
            if ($perms.Item($i).UserName -eq $User) {
                $hasEntry = $true
                break
            }
        }

        if (-not $hasEntry) {
            # Auto-grant "Can view" (level 1)
            $newPerm = $perms.Add($User)
            $newPerm.OlFolderPermission = 1
            $perms.Save()
            $autoGranted += @{ FolderName = $current.Name; Level = 'Can view' }
        }

        $current = $current.Parent
    }

    return [PSCustomObject]@{
        Success     = $true
        AutoGranted = $autoGranted
    }
}

function Remove-FolderPermission {
    <#
      Removes a user's permission entry from a folder.
    #>
    param(
        [Parameter(Mandatory=$true)][string]$EntryID,
        [Parameter(Mandatory=$true)][string]$StoreID,
        [Parameter(Mandatory=$true)][string]$User
    )
    $ns     = Get-OutlookNS
    $folder = $ns.GetFolderFromID($EntryID, $StoreID)
    $perms  = $folder.FolderPermissions
    for ($i = $perms.Count; $i -ge 1; $i--) {
        if ($perms.Item($i).UserName -eq $User) {
            $perms.Remove($i)
            break
        }
    }
    $perms.Save()
}

# ─── Active Directory user search ────────────────────────────────────────────

function Search-ADUsers {
    <#
      Searches Active Directory for OG- mail-enabled groups matching the query.
      Returns up to 20 results.
      Each item: [PSCustomObject]@{ DisplayName; Mail; SamAccountName }
      Returns empty array gracefully if not domain-joined or AD unreachable.

      3-branch query logic (matches working reference implementation):
        @  in query  →  rewrite as sAMAccountName=OG-{domain}-{user}  (email input)
        OG-* prefix  →  sAMAccountName wildcard OG-input*
        free text    →  displayName wildcard OG-*input*

      Performance notes:
      - Minimum query length is 5 to avoid massive unindexed wildcard scans
      - Results are cached in-memory for 60 seconds per unique query
      - ServerTimeLimit caps the DC wait at 4 seconds to avoid UI hangs
      - objectCategory=group scopes to AD groups only (OG- are mail-enabled groups)
    #>
    param([string]$Query)
    if ([string]::IsNullOrWhiteSpace($Query) -or $Query.Length -lt 5) { return @() }
    if ($script:ADSearchFactory) { return & $script:ADSearchFactory $Query }

    # Cache lookup
    $cacheKey = $Query.ToLower()
    if ($script:ADCache.ContainsKey($cacheKey)) {
        $entry = $script:ADCache[$cacheKey]
        if ([datetime]::Now -lt $entry.Expires) {
            return $entry.Results
        }
        $script:ADCache.Remove($cacheKey)
    }

    try {
        Add-Type -AssemblyName System.DirectoryServices -ErrorAction Stop
        $searcher = New-Object System.DirectoryServices.DirectorySearcher
        # Escape special LDAP chars in query
        $safe = $Query -replace '[\\*()\x00]',''

        # 3-branch LDAP filter — always scoped to objectCategory=group
        if ($Query -match '@') {
            # Email format (user@domain.de) → rewrite to OG-domain.de-user sAMAccountName
            $parts     = $Query -split '@', 2
            $safePart0 = $parts[0] -replace '[\\*()\x00]',''
            $safePart1 = if ($parts.Count -gt 1) { $parts[1] -replace '[\\*()\x00]','' } else { '' }
            if ([string]::IsNullOrEmpty($safePart0) -or [string]::IsNullOrEmpty($safePart1)) {
                return @()   # incomplete email — skip query
            }
            $sam   = "OG-$safePart1-$safePart0"
            $ldap  = "(&(objectCategory=group)(sAMAccountName=$sam))"
        } elseif ($Query -like 'OG-*') {
            # Already OG- prefixed → direct sAMAccountName wildcard
            $ldap  = "(&(objectCategory=group)(sAMAccountName=$safe*))"
        } else {
            # Free text → displayName wildcard with OG- prefix anchored
            $ldap  = "(&(objectCategory=group)(displayName=OG-*$safe*))"
        }

        $searcher.Filter = $ldap
        $searcher.PropertiesToLoad.AddRange(@('displayName','mail','sAMAccountName')) | Out-Null
        $searcher.SizeLimit        = 20
        $searcher.ServerTimeLimit  = [System.TimeSpan]::FromSeconds(4)
        $searcher.CacheResults     = $false   # we do our own caching
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
        # Cache for 60 seconds
        $script:ADCache[$cacheKey] = @{ Results = $out; Expires = [datetime]::Now.AddSeconds(60) }
        return $out
    } catch {
        Write-Verbose "AD search failed (may not be domain-joined): $_"
        return @()
    }
}

# ─── Friendly permission level list (for UI dropdowns) ───────────────────────

function Get-PermissionLevels {
    <# Returns ordered list of friendly permission level strings for UI display. #>
    return $script:FriendlyToOlPerm.Keys | ForEach-Object { $_ }
}

# ─── Simplified permission levels (for Easy-mode wizard) ─────────────────────

function Get-SimplifiedPermissionLevels {
    <#
      Returns a concise list of permission choices for the guided wizard.
      Each item: [PSCustomObject]@{ Label; Description; OlLevel }
    #>
    return @(
        [PSCustomObject]@{ Label = 'Just look (read only)';       Description = 'Can see items but cannot change anything';     OlLevel = 1 }
        [PSCustomObject]@{ Label = 'Add new items';               Description = 'Can create new items in the folder';           OlLevel = 2 }
        [PSCustomObject]@{ Label = 'Add and edit everything';     Description = 'Can create, edit, and delete any items';       OlLevel = 5 }
        [PSCustomObject]@{ Label = 'Full control';                Description = 'Can do everything including manage settings';  OlLevel = 7 }
        [PSCustomObject]@{ Label = 'No access (block)';           Description = 'Cannot see or access this folder at all';      OlLevel = 0 }
    )
}

# ─── Permissions overview (for Easy-mode overview report) ────────────────────

function Get-PermissionsOverview {
    <#
      Scans all accounts, all folders, and all permissions to build an overview.
      Returns @( [PSCustomObject]@{
          Mailbox     = 'Alice Johnson'
          SmtpAddress = 'alice@contoso.com'
          Entries     = @( [PSCustomObject]@{
              User    = 'Bob Smith'
              Folders = @( [PSCustomObject]@{ FolderName; FolderPath; EntryID; StoreID; Level } )
          } )
      } )
      Filters out Default/Anonymous users and "No access" (level 0) entries.
    #>
    $accounts = Get-SignedInAccounts
    $results = @()
    foreach ($acct in $accounts) {
        $folders = Get-MailboxFolders -SmtpAddress $acct.SmtpAddress
        $userMap = [ordered]@{}
        foreach ($f in $folders) {
            $perms = Get-FolderPermissions -EntryID $f.EntryID -StoreID $f.StoreID
            foreach ($p in $perms) {
                if ($p.User -eq 'Default' -or $p.User -eq 'Anonymous') { continue }
                if ($p.PermissionLevel -eq 0) { continue }
                if (-not $userMap.Contains($p.User)) { $userMap[$p.User] = @() }
                $userMap[$p.User] += [PSCustomObject]@{
                    FolderName = $f.Name
                    FolderPath = $f.FolderPath
                    EntryID    = $f.EntryID
                    StoreID    = $f.StoreID
                    Level      = $p.PermissionLevelName
                }
            }
        }
        $entries = @()
        foreach ($user in $userMap.Keys) {
            $entries += [PSCustomObject]@{ User = $user; Folders = $userMap[$user] }
        }
        $results += [PSCustomObject]@{
            Mailbox     = $acct.Name
            SmtpAddress = $acct.SmtpAddress
            Entries     = $entries
        }
    }
    return $results
}

Export-ModuleMember -Function `
    Get-SignedInAccounts, `
    Get-MailboxFolders, `
    Get-FolderPermissions, `
    Set-FolderPermission, `
    Set-FolderPermissionWithAncestors, `
    Remove-FolderPermission, `
    Search-ADUsers, `
    Get-PermissionLevels, `
    Get-SimplifiedPermissionLevels, `
    Get-PermissionsOverview, `
    Set-OutlookNSFactory, `
    Set-ADSearchFactory
