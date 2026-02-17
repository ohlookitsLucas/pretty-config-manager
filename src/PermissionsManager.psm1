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

function script:Get-OutlookNS {
    $ol = New-Object -ComObject Outlook.Application
    return $ol.GetNameSpace('MAPI')
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

function script:Get-FoldersRecursive {
    param($Folder, [int]$Depth = 0)
    $results = @()
    $results += [PSCustomObject]@{
        Name        = $Folder.Name
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
        # Find the Store whose root display matches the account
        for ($i = 1; $i -le $ns.Stores.Count; $i++) {
            $store = $ns.Stores.Item($i)
            try {
                $root = $store.GetRootFolder()
                # Match by SMTP embedded in the folder path or display name
                $storeSmtp = ''
                try {
                    $acc = $ns.Accounts | Where-Object { $_.SmtpAddress -eq $SmtpAddress } | Select-Object -First 1
                    if ($null -eq $acc) { continue }
                    # The store for this account typically has the display name matching the account
                    if ($store.DisplayName -ne $acc.DisplayName -and
                        $store.DisplayName -ne $SmtpAddress -and
                        $root.Name -ne $acc.DisplayName) {
                        # Try matching via ExchangeStoreType or fallback: use first store for primary account
                    }
                } catch {}
                # Just enumerate folders for every store that matches any account with this SMTP
                # A simple heuristic: return folders for the first store whose name contains the SMTP or display name
                $accNames = @()
                for ($j = 1; $j -le $ns.Accounts.Count; $j++) {
                    $a = $ns.Accounts.Item($j)
                    if ($a.SmtpAddress -eq $SmtpAddress) {
                        $accNames += $a.DisplayName
                        $accNames += $a.SmtpAddress
                    }
                }
                $storeName = $store.DisplayName
                $match = $accNames | Where-Object { $storeName -like "*$_*" -or $_ -like "*$storeName*" }
                if ($match -or ($ns.Stores.Count -eq 1)) {
                    return Get-FoldersRecursive -Folder $root -Depth 0
                }
            } catch {}
        }
        # Fallback: return root folders of first store
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
        $perms  = $folder.Permission
        $results = @()
        for ($i = 1; $i -le $perms.Count; $i++) {
            $p     = $perms.Item($i)
            $level = $p.PermissionSet
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
    $perms  = $folder.Permission

    # Check if user already exists
    $existing = $null
    for ($i = 1; $i -le $perms.Count; $i++) {
        if ($perms.Item($i).UserName -eq $User) {
            $existing = $perms.Item($i)
            break
        }
    }

    if ($null -ne $existing) {
        $existing.PermissionSet = $levelNum
    } else {
        $newPerm = $perms.Add($User)
        $newPerm.PermissionSet = $levelNum
    }
    $perms.Save()
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
    $perms  = $folder.Permission
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
      Searches Active Directory for users matching a display name, email, or
      sAMAccountName. Returns up to 20 results.
      Each item: [PSCustomObject]@{ DisplayName; Mail; SamAccountName }
      Returns empty array gracefully if not domain-joined or AD unreachable.
    #>
    param([string]$Query)
    if ([string]::IsNullOrWhiteSpace($Query) -or $Query.Length -lt 2) { return @() }
    try {
        Add-Type -AssemblyName System.DirectoryServices -ErrorAction Stop
        $searcher = New-Object System.DirectoryServices.DirectorySearcher
        # Escape special LDAP chars in query
        $safe = $Query -replace '[\\*()\x00]',''
        $searcher.Filter = "(|" +
            "(displayName=*$safe*)" +
            "(mail=*$safe*)" +
            "(sAMAccountName=*$safe*)" +
            "(cn=*$safe*)" +
            ")"
        $searcher.PropertiesToLoad.AddRange(@('displayName','mail','sAMAccountName')) | Out-Null
        $searcher.SizeLimit = 20
        $results = $searcher.FindAll()
        $out = @()
        foreach ($r in $results) {
            $dn   = if ($r.Properties['displayName'].Count -gt 0) { $r.Properties['displayName'][0] } else { '' }
            $mail = if ($r.Properties['mail'].Count -gt 0)        { $r.Properties['mail'][0] }        else { '' }
            $sam  = if ($r.Properties['sAMAccountName'].Count -gt 0) { $r.Properties['sAMAccountName'][0] } else { '' }
            if ($dn -or $mail -or $sam) {
                $out += [PSCustomObject]@{ DisplayName = $dn; Mail = $mail; SamAccountName = $sam }
            }
        }
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

Export-ModuleMember -Function `
    Get-SignedInAccounts, `
    Get-MailboxFolders, `
    Get-FolderPermissions, `
    Set-FolderPermission, `
    Remove-FolderPermission, `
    Search-ADUsers, `
    Get-PermissionLevels
