<#
  New-MockOutlookNS.ps1
  Returns a PSObject tree that mimics the Outlook COM MAPI namespace.
  Used by both Pester tests and the sandbox UI launcher.

  Structure mirrors:
    $ns.Accounts.Count / .Item(i) -> { DisplayName, SmtpAddress }
    $ns.Stores.Count   / .Item(i) -> { DisplayName, GetRootFolder() }
    folder -> { Name, FolderPath, EntryID, StoreID, Folders, FolderPermissions }
    permissions -> { Count, Item(i), Add(user), Remove(i), Save() }
    permission entry -> { UserName, OlFolderPermission }

  Permissions are fully mutable — Add/Remove/Save work in-memory.
  Each folder gets its own independent permission list (no shared references).

  Every folder has a .Parent property pointing to its parent folder (or $null for root).
  This mirrors the Outlook COM object model and is used by Set-FolderPermissionWithAncestors.

  Pre-loaded permissions by account:
    Alice Johnson   – Inbox: Bob=1, Charlie=2; Calendar: Bob=7, Diana=3; Sent/Contacts/Drafts: Default+Anon
                      Inbox > Projects (Depth 2): Default+Anon only
                      Inbox > Projects > Client A (Depth 3): Default+Anon only
                      Inbox > Archive (Depth 2): Default+Anon only
    Shared Services – Inbox: Alice=7, Bob=7, Eva=4; Calendar: Alice=7, Frank=3; Sent/Contacts/Drafts: Default+Anon
    Bob (Delegate)  – Inbox/Calendar/etc.: Default+Anon only (clean slate for testing Add)
#>

function New-MockOutlookNS {
    param(
        [hashtable[]]$Accounts = @(
            @{ DisplayName = 'Alice Johnson';   SmtpAddress = 'alice@contoso.com' }
            @{ DisplayName = 'Shared Services'; SmtpAddress = 'shared@contoso.com' }
            @{ DisplayName = 'Bob (Delegate)';  SmtpAddress = 'bob.delegate@contoso.com' }
        )
    )

    # ── Permission entry builder ──────────────────────────────────────────────
    function New-MockPermEntry {
        param([string]$UserName, [int]$Level)
        return [PSCustomObject]@{ UserName = $UserName; OlFolderPermission = $Level }
    }

    # ── Permission collection builder ─────────────────────────────────────────
    # Each call returns a *new* independent ArrayList — no shared references.
    function New-MockPermissions {
        param([array]$Entries = @())
        $list = [System.Collections.ArrayList]::new()
        foreach ($e in $Entries) {
            # Clone each entry so callers can't alias into a shared object
            $list.Add([PSCustomObject]@{
                UserName           = $e.UserName
                OlFolderPermission = $e.OlFolderPermission
            }) | Out-Null
        }

        $perms = [PSCustomObject]@{ _list = $list }

        $perms | Add-Member -MemberType ScriptProperty -Name Count -Value {
            $this._list.Count
        }

        $perms | Add-Member -MemberType ScriptMethod -Name Item -Value {
            param([int]$i)
            return $this._list[$i - 1]   # COM is 1-based
        }

        $perms | Add-Member -MemberType ScriptMethod -Name Add -Value {
            param([string]$User)
            $entry = [PSCustomObject]@{ UserName = $User; OlFolderPermission = 0 }
            $this._list.Add($entry) | Out-Null
            return $entry
        }

        $perms | Add-Member -MemberType ScriptMethod -Name Remove -Value {
            param([int]$i)
            $this._list.RemoveAt($i - 1)   # COM is 1-based
        }

        $perms | Add-Member -MemberType ScriptMethod -Name Save -Value {
            # No-op in mock — changes are already live in _list
        }

        return $perms
    }

    # ── Folder builder ────────────────────────────────────────────────────────
    function New-MockFolder {
        param(
            [string]$Name,
            [string]$FolderPath,
            [string]$EntryID,
            [string]$StoreID,
            [array]$SubFolders  = @(),
            [array]$PermEntries = @()
        )
        $childList = [System.Collections.ArrayList]::new()
        foreach ($sf in $SubFolders) { $childList.Add($sf) | Out-Null }

        $folders = [PSCustomObject]@{ _list = $childList }
        $folders | Add-Member -MemberType ScriptProperty -Name Count -Value { $this._list.Count }
        $folders | Add-Member -MemberType ScriptMethod   -Name Item  -Value {
            param([int]$i)
            return $this._list[$i - 1]
        }

        $f = [PSCustomObject]@{
            Name              = $Name
            FolderPath        = $FolderPath
            EntryID           = $EntryID
            StoreID           = $StoreID
            Folders           = $folders
            FolderPermissions = (New-MockPermissions -Entries $PermEntries)
            Parent            = $null   # wired up later by Wire-ParentRefs
        }
        return $f
    }

    # ── Standard entries reused across folders ────────────────────────────────
    # Helper so every call produces fresh objects (no reference aliasing)
    function _DefaultAnon { @(
        (New-MockPermEntry 'Default'   0)
        (New-MockPermEntry 'Anonymous' 0)
    )}

    # ── Wire .Parent references recursively ─────────────────────────────────
    # Mirrors the Outlook COM model where every folder has a .Parent pointing
    # to its containing folder. Root folders get Parent = $null.
    function Wire-ParentRefs {
        param($Folder)
        if ($Folder.Folders -and $Folder.Folders.Count -gt 0) {
            for ($k = 1; $k -le $Folder.Folders.Count; $k++) {
                $child = $Folder.Folders.Item($k)
                $child.Parent = $Folder
                Wire-ParentRefs -Folder $child
            }
        }
    }

    # ── Per-account folder trees ──────────────────────────────────────────────

    function New-AliceTree {
        param([string]$StoreID)
        $root = '\\Alice Johnson'

        # Subfolders of Inbox — for testing ancestor permission auto-grant
        $clientA = New-MockFolder 'Client A' "$root\Inbox\Projects\Client A" "entry-$StoreID-clienta" $StoreID `
            -PermEntries (_DefaultAnon)
        $projects = New-MockFolder 'Projects' "$root\Inbox\Projects" "entry-$StoreID-projects" $StoreID `
            -SubFolders @($clientA) -PermEntries (_DefaultAnon)
        $archive = New-MockFolder 'Archive' "$root\Inbox\Archive" "entry-$StoreID-archive" $StoreID `
            -PermEntries (_DefaultAnon)

        $inbox = New-MockFolder 'Inbox' "$root\Inbox" "entry-$StoreID-inbox" $StoreID `
            -SubFolders @($projects, $archive) -PermEntries @(
            (New-MockPermEntry 'Default'         0)
            (New-MockPermEntry 'Anonymous'       0)
            (New-MockPermEntry 'Bob Smith'       1)   # Can view
            (New-MockPermEntry 'Charlie Brown'   2)   # Can create items
        )
        $sent = New-MockFolder 'Sent Items' "$root\Sent Items" "entry-$StoreID-sent" $StoreID `
            -PermEntries (_DefaultAnon)
        $calendar = New-MockFolder 'Calendar' "$root\Calendar" "entry-$StoreID-cal" $StoreID -PermEntries @(
            (New-MockPermEntry 'Default'         1)   # Can view (typical calendar default)
            (New-MockPermEntry 'Anonymous'       0)
            (New-MockPermEntry 'Bob Smith'       7)   # Full access
            (New-MockPermEntry 'Diana Prince'    3)   # Can view & create
        )
        $contacts = New-MockFolder 'Contacts' "$root\Contacts" "entry-$StoreID-contacts" $StoreID `
            -PermEntries (_DefaultAnon)
        $drafts = New-MockFolder 'Drafts' "$root\Drafts" "entry-$StoreID-drafts" $StoreID `
            -PermEntries (_DefaultAnon)

        return New-MockFolder 'Alice Johnson' $root "entry-$StoreID-root" $StoreID `
            -SubFolders @($inbox, $sent, $calendar, $contacts, $drafts) `
            -PermEntries (_DefaultAnon)
    }

    function New-SharedTree {
        param([string]$StoreID)
        $root = '\\Shared Services'

        $inbox = New-MockFolder 'Inbox' "$root\Inbox" "entry-$StoreID-inbox" $StoreID -PermEntries @(
            (New-MockPermEntry 'Default'         0)
            (New-MockPermEntry 'Anonymous'       0)
            (New-MockPermEntry 'Alice Johnson'   7)   # Full access
            (New-MockPermEntry 'Bob Smith'       7)   # Full access
            (New-MockPermEntry 'Eva Mueller'     4)   # Can create, edit & delete
        )
        $sent = New-MockFolder 'Sent Items' "$root\Sent Items" "entry-$StoreID-sent" $StoreID -PermEntries @(
            (New-MockPermEntry 'Default'         0)
            (New-MockPermEntry 'Anonymous'       0)
            (New-MockPermEntry 'Alice Johnson'   7)
            (New-MockPermEntry 'Bob Smith'       7)
        )
        $calendar = New-MockFolder 'Calendar' "$root\Calendar" "entry-$StoreID-cal" $StoreID -PermEntries @(
            (New-MockPermEntry 'Default'         1)
            (New-MockPermEntry 'Anonymous'       0)
            (New-MockPermEntry 'Alice Johnson'   7)
            (New-MockPermEntry 'Frank Weber'     3)   # Can view & create
        )
        $contacts = New-MockFolder 'Contacts' "$root\Contacts" "entry-$StoreID-contacts" $StoreID `
            -PermEntries (_DefaultAnon)
        $drafts = New-MockFolder 'Drafts' "$root\Drafts" "entry-$StoreID-drafts" $StoreID `
            -PermEntries (_DefaultAnon)

        return New-MockFolder 'Shared Services' $root "entry-$StoreID-root" $StoreID `
            -SubFolders @($inbox, $sent, $calendar, $contacts, $drafts) `
            -PermEntries (_DefaultAnon)
    }

    function New-BobTree {
        param([string]$StoreID)
        $root = '\\Bob (Delegate)'

        # Clean slate — only Default + Anonymous, so you can test adding from scratch
        $inbox    = New-MockFolder 'Inbox'      "$root\Inbox"      "entry-$StoreID-inbox"    $StoreID -PermEntries (_DefaultAnon)
        $sent     = New-MockFolder 'Sent Items' "$root\Sent Items" "entry-$StoreID-sent"     $StoreID -PermEntries (_DefaultAnon)
        $calendar = New-MockFolder 'Calendar'   "$root\Calendar"   "entry-$StoreID-cal"      $StoreID -PermEntries @(
            (New-MockPermEntry 'Default'   1)
            (New-MockPermEntry 'Anonymous' 0)
        )
        $contacts = New-MockFolder 'Contacts'   "$root\Contacts"   "entry-$StoreID-contacts" $StoreID -PermEntries (_DefaultAnon)
        $drafts   = New-MockFolder 'Drafts'     "$root\Drafts"     "entry-$StoreID-drafts"   $StoreID -PermEntries (_DefaultAnon)

        return New-MockFolder 'Bob (Delegate)' $root "entry-$StoreID-root" $StoreID `
            -SubFolders @($inbox, $sent, $calendar, $contacts, $drafts) `
            -PermEntries (_DefaultAnon)
    }

    # ── Router: pick the right tree builder per account ───────────────────────
    $treeBuilders = @{
        'Alice Johnson'   = { param($sid) New-AliceTree  -StoreID $sid }
        'Shared Services' = { param($sid) New-SharedTree -StoreID $sid }
        'Bob (Delegate)'  = { param($sid) New-BobTree    -StoreID $sid }
    }

    # ── Build accounts collection ─────────────────────────────────────────────
    $acctList = [System.Collections.ArrayList]::new()
    foreach ($a in $Accounts) {
        $acctList.Add([PSCustomObject]@{
            DisplayName = $a.DisplayName
            SmtpAddress = $a.SmtpAddress
        }) | Out-Null
    }
    $acctsColl = [PSCustomObject]@{ _list = $acctList }
    $acctsColl | Add-Member -MemberType ScriptProperty -Name Count -Value { $this._list.Count }
    $acctsColl | Add-Member -MemberType ScriptMethod   -Name Item  -Value {
        param([int]$i)
        return $this._list[$i - 1]
    }

    # ── Build stores collection ───────────────────────────────────────────────
    $storeList    = [System.Collections.ArrayList]::new()
    $folderLookup = @{}   # "$EntryID|$StoreID" -> folder object

    function Index-Folders {
        param($Folder)
        $folderLookup["$($Folder.EntryID)|$($Folder.StoreID)"] = $Folder
        if ($Folder.Folders -and $Folder.Folders.Count -gt 0) {
            for ($j = 1; $j -le $Folder.Folders.Count; $j++) {
                Index-Folders -Folder $Folder.Folders.Item($j)
            }
        }
    }

    for ($idx = 0; $idx -lt $Accounts.Count; $idx++) {
        $a       = $Accounts[$idx]
        $storeID = "store-$idx"

        # Use named builder if available; fall back to a generic empty tree
        $rootFolder = $null
        if ($treeBuilders.ContainsKey($a.DisplayName)) {
            $rootFolder = & $treeBuilders[$a.DisplayName] $storeID
        } else {
            # Generic fallback: Inbox + Calendar with Default/Anon only
            $rp = "\\$($a.DisplayName)"
            $inbox    = New-MockFolder 'Inbox'    "$rp\Inbox"    "entry-$storeID-inbox" $storeID -PermEntries (_DefaultAnon)
            $calendar = New-MockFolder 'Calendar' "$rp\Calendar" "entry-$storeID-cal"   $storeID -PermEntries @(
                (New-MockPermEntry 'Default' 1); (New-MockPermEntry 'Anonymous' 0)
            )
            $rootFolder = New-MockFolder $a.DisplayName $rp "entry-$storeID-root" $storeID `
                -SubFolders @($inbox, $calendar)
        }

        Wire-ParentRefs -Folder $rootFolder
        Index-Folders -Folder $rootFolder

        $store = [PSCustomObject]@{ DisplayName = $a.DisplayName; _rootFolder = $rootFolder }
        $store | Add-Member -MemberType ScriptMethod -Name GetRootFolder -Value { return $this._rootFolder }
        $storeList.Add($store) | Out-Null
    }

    $storesColl = [PSCustomObject]@{ _list = $storeList }
    $storesColl | Add-Member -MemberType ScriptProperty -Name Count -Value { $this._list.Count }
    $storesColl | Add-Member -MemberType ScriptMethod   -Name Item  -Value {
        param([int]$i)
        return $this._list[$i - 1]
    }

    # ── MAPI namespace ────────────────────────────────────────────────────────
    $ns = [PSCustomObject]@{
        Accounts      = $acctsColl
        Stores        = $storesColl
        _folderLookup = $folderLookup
    }

    $ns | Add-Member -MemberType ScriptMethod -Name GetFolderFromID -Value {
        param([string]$EntryID, [string]$StoreID)
        $key = "$EntryID|$StoreID"
        if ($this._folderLookup.ContainsKey($key)) { return $this._folderLookup[$key] }
        throw "Folder not found: EntryID=$EntryID, StoreID=$StoreID"
    }

    return $ns
}

# ── Mock AD search factory ────────────────────────────────────────────────────

function New-MockADSearchFactory {
    <#
      Returns a scriptblock suitable for Set-ADSearchFactory.
      8 fake AD users — all names also appear in the mock permissions above,
      so AD search results can be directly used to add/remove permissions.
    #>
    $mockUsers = @(
        [PSCustomObject]@{ DisplayName = 'Alice Johnson';   Mail = 'alice@contoso.com';        SamAccountName = 'ajohnson' }
        [PSCustomObject]@{ DisplayName = 'Bob Smith';       Mail = 'bob.delegate@contoso.com';  SamAccountName = 'bsmith'   }
        [PSCustomObject]@{ DisplayName = 'Charlie Brown';   Mail = 'charlie@contoso.com';       SamAccountName = 'cbrown'   }
        [PSCustomObject]@{ DisplayName = 'Diana Prince';    Mail = 'diana@contoso.com';         SamAccountName = 'dprince'  }
        [PSCustomObject]@{ DisplayName = 'Shared Services'; Mail = 'shared@contoso.com';        SamAccountName = 'shared'   }
        [PSCustomObject]@{ DisplayName = 'Eva Mueller';     Mail = 'eva.mueller@contoso.com';   SamAccountName = 'emueller' }
        [PSCustomObject]@{ DisplayName = 'Frank Weber';     Mail = 'frank.weber@contoso.com';   SamAccountName = 'fweber'   }
        [PSCustomObject]@{ DisplayName = 'Grace Lee';       Mail = 'grace.lee@contoso.com';     SamAccountName = 'glee'     }
    )

    return {
        param([string]$Query)
        $q = $Query.ToLower()
        return @($mockUsers | Where-Object {
            $_.DisplayName.ToLower().Contains($q) -or
            $_.Mail.ToLower().Contains($q)        -or
            $_.SamAccountName.ToLower().Contains($q)
        } | Select-Object -First 20)
    }.GetNewClosure()
}
