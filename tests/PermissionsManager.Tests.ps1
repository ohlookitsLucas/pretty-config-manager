<#
  PermissionsManager.Tests.ps1
  Pester 5 unit tests for PermissionsManager.psm1
  Uses mock COM/AD factories — no Outlook or Active Directory required.
#>

BeforeAll {
    # Load mock helpers
    . (Join-Path $PSScriptRoot 'Mocks\New-MockOutlookNS.ps1')

    # Import the module
    $modulePath = Join-Path $PSScriptRoot '..\src\PermissionsManager.psm1'
    Import-Module $modulePath -Force

    # Wire up mock factories
    $mockNS = New-MockOutlookNS
    Set-OutlookNSFactory { $mockNS }.GetNewClosure()
    Set-ADSearchFactory  (New-MockADSearchFactory)
}

AfterAll {
    # Reset factories to default (null)
    Set-OutlookNSFactory { $null }
    Set-ADSearchFactory  { $null }
}

Describe 'Get-SignedInAccounts' {
    It 'Returns 3 mock accounts' {
        $accounts = Get-SignedInAccounts
        $accounts.Count | Should -Be 3
    }

    It 'Contains alice@contoso.com' {
        $accounts = Get-SignedInAccounts
        $alice = $accounts | Where-Object { $_.SmtpAddress -eq 'alice@contoso.com' }
        $alice | Should -Not -BeNullOrEmpty
        $alice.Name | Should -Be 'Alice Johnson'
    }

    It 'Contains shared@contoso.com' {
        $accounts = Get-SignedInAccounts
        $shared = $accounts | Where-Object { $_.SmtpAddress -eq 'shared@contoso.com' }
        $shared | Should -Not -BeNullOrEmpty
    }

    It 'Contains bob.delegate@contoso.com' {
        $accounts = Get-SignedInAccounts
        $bob = $accounts | Where-Object { $_.SmtpAddress -eq 'bob.delegate@contoso.com' }
        $bob | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-MailboxFolders' {
    It 'Returns folders for alice@contoso.com' {
        $folders = Get-MailboxFolders -SmtpAddress 'alice@contoso.com'
        $folders.Count | Should -BeGreaterThan 0
    }

    It 'Includes Inbox, Sent Items, Calendar, Contacts, Drafts' {
        $folders = Get-MailboxFolders -SmtpAddress 'alice@contoso.com'
        $names = $folders | ForEach-Object { $_.Name }
        $names | Should -Contain 'Inbox'
        $names | Should -Contain 'Sent Items'
        $names | Should -Contain 'Calendar'
        $names | Should -Contain 'Contacts'
        $names | Should -Contain 'Drafts'
    }

    It 'Root folder has Depth 0' {
        $folders = Get-MailboxFolders -SmtpAddress 'alice@contoso.com'
        $folders[0].Depth | Should -Be 0
    }

    It 'Child folders have Depth 1' {
        $folders = Get-MailboxFolders -SmtpAddress 'alice@contoso.com'
        $inbox = $folders | Where-Object { $_.Name -eq 'Inbox' }
        $inbox.Depth | Should -Be 1
    }
}

Describe 'Get-FolderPermissions' {
    It 'Returns Default and Anonymous permissions for Inbox' {
        $folders = Get-MailboxFolders -SmtpAddress 'alice@contoso.com'
        $inbox = $folders | Where-Object { $_.Name -eq 'Inbox' }
        $perms = Get-FolderPermissions -EntryID $inbox.EntryID -StoreID $inbox.StoreID
        # Alice's Inbox now has Default, Anonymous, Bob Smith, Charlie Brown pre-loaded
        $perms.Count | Should -BeGreaterOrEqual 2
        ($perms | Where-Object { $_.User -eq 'Default' })   | Should -Not -BeNullOrEmpty
        ($perms | Where-Object { $_.User -eq 'Anonymous' }) | Should -Not -BeNullOrEmpty
    }

    It 'Calendar has Default with Can view permission' {
        $folders = Get-MailboxFolders -SmtpAddress 'alice@contoso.com'
        $cal = $folders | Where-Object { $_.Name -eq 'Calendar' }
        $perms = Get-FolderPermissions -EntryID $cal.EntryID -StoreID $cal.StoreID
        $defaultPerm = $perms | Where-Object { $_.User -eq 'Default' }
        $defaultPerm.PermissionLevelName | Should -Be 'Can view'
    }
}

Describe 'Set-FolderPermission' {
    It 'Updates an existing permission' {
        $folders = Get-MailboxFolders -SmtpAddress 'alice@contoso.com'
        $inbox = $folders | Where-Object { $_.Name -eq 'Inbox' }

        Set-FolderPermission -EntryID $inbox.EntryID -StoreID $inbox.StoreID -User 'Default' -Level 'Can view'

        $perms = Get-FolderPermissions -EntryID $inbox.EntryID -StoreID $inbox.StoreID
        $defaultPerm = $perms | Where-Object { $_.User -eq 'Default' }
        $defaultPerm.PermissionLevel | Should -Be 1
    }

    It 'Adds a new permission for a new user' {
        $folders = Get-MailboxFolders -SmtpAddress 'alice@contoso.com'
        $inbox = $folders | Where-Object { $_.Name -eq 'Inbox' }

        $before = (Get-FolderPermissions -EntryID $inbox.EntryID -StoreID $inbox.StoreID).Count
        Set-FolderPermission -EntryID $inbox.EntryID -StoreID $inbox.StoreID -User 'bsmith' -Level 'Can view'
        $after = (Get-FolderPermissions -EntryID $inbox.EntryID -StoreID $inbox.StoreID).Count
        $after | Should -Be ($before + 1)
    }
}

Describe 'Remove-FolderPermission' {
    It 'Removes a permission entry' {
        $folders = Get-MailboxFolders -SmtpAddress 'alice@contoso.com'
        $inbox = $folders | Where-Object { $_.Name -eq 'Inbox' }

        # Add then remove
        Set-FolderPermission -EntryID $inbox.EntryID -StoreID $inbox.StoreID -User 'testuser' -Level 1
        $before = (Get-FolderPermissions -EntryID $inbox.EntryID -StoreID $inbox.StoreID).Count
        Remove-FolderPermission -EntryID $inbox.EntryID -StoreID $inbox.StoreID -User 'testuser'
        $after = (Get-FolderPermissions -EntryID $inbox.EntryID -StoreID $inbox.StoreID).Count
        $after | Should -Be ($before - 1)
    }
}

Describe 'Search-ADUsers' {
    It 'Returns results for "alice"' {
        $results = @(Search-ADUsers -Query 'alice')
        $results.Count | Should -BeGreaterThan 0
        $results[0].DisplayName | Should -Be 'Alice Johnson'
    }

    It 'Searches by email' {
        $results = @(Search-ADUsers -Query 'shared@contoso')
        $results.Count | Should -BeGreaterThan 0
    }

    It 'Searches by SAM account name' {
        $results = @(Search-ADUsers -Query 'cbrown')
        $results.Count | Should -Be 1
        $results[0].DisplayName | Should -Be 'Charlie Brown'
    }

    It 'Returns empty for query shorter than 2 chars' {
        $results = Search-ADUsers -Query 'a'
        $results | Should -HaveCount 0
    }

    It 'Returns empty for blank query' {
        $results = Search-ADUsers -Query ''
        $results | Should -HaveCount 0
    }

    It 'Returns multiple matches for partial name' {
        $results = Search-ADUsers -Query 'contoso'
        $results.Count | Should -BeGreaterThan 1
    }
}

Describe 'Get-PermissionLevels' {
    It 'Returns the known friendly names' {
        $levels = Get-PermissionLevels
        $levels | Should -Contain 'No access'
        $levels | Should -Contain 'Owner'
        $levels | Should -Contain 'Can view'
        $levels | Should -Contain 'Full access'
    }
}

Describe 'Get-SimplifiedPermissionLevels' {
    It 'Returns 5 simplified levels' {
        $levels = Get-SimplifiedPermissionLevels
        $levels.Count | Should -Be 5
    }

    It 'Each level has Label, Description, and OlLevel' {
        $levels = Get-SimplifiedPermissionLevels
        foreach ($l in $levels) {
            $l.Label       | Should -Not -BeNullOrEmpty
            $l.Description | Should -Not -BeNullOrEmpty
            $l.OlLevel     | Should -BeOfType [int]
        }
    }

    It 'OlLevel values are valid permission numbers' {
        $levels = Get-SimplifiedPermissionLevels
        $valid  = @(0, 1, 2, 5, 7)
        foreach ($l in $levels) {
            $l.OlLevel | Should -BeIn $valid
        }
    }
}

Describe 'Get-PermissionsOverview' {
    It 'Returns overview for all 3 mock accounts' {
        $overview = Get-PermissionsOverview
        $overview.Count | Should -Be 3
    }

    It 'Alice account shows Bob Smith with access' {
        $overview = Get-PermissionsOverview
        $alice = $overview | Where-Object { $_.SmtpAddress -eq 'alice@contoso.com' }
        $bob   = $alice.Entries | Where-Object { $_.User -eq 'Bob Smith' }
        $bob | Should -Not -BeNullOrEmpty
        $bob.Folders.Count | Should -BeGreaterThan 0
    }

    It 'Alice account Bob has Inbox and Calendar entries' {
        $overview = Get-PermissionsOverview
        $alice = $overview | Where-Object { $_.SmtpAddress -eq 'alice@contoso.com' }
        $bob   = $alice.Entries | Where-Object { $_.User -eq 'Bob Smith' }
        $folderNames = $bob.Folders | ForEach-Object { $_.FolderName }
        $folderNames | Should -Contain 'Inbox'
        $folderNames | Should -Contain 'Calendar'
    }

    It 'Excludes Default and Anonymous users' {
        $overview = Get-PermissionsOverview
        foreach ($acct in $overview) {
            foreach ($entry in $acct.Entries) {
                $entry.User | Should -Not -Be 'Default'
                $entry.User | Should -Not -Be 'Anonymous'
            }
        }
    }

    It 'Excludes entries with No access (level 0)' {
        $overview = Get-PermissionsOverview
        foreach ($acct in $overview) {
            foreach ($entry in $acct.Entries) {
                foreach ($f in $entry.Folders) {
                    $f.Level | Should -Not -Be 'No access'
                }
            }
        }
    }

    It 'Shared mailbox shows Frank Weber with Calendar access' {
        $overview = Get-PermissionsOverview
        $shared = $overview | Where-Object { $_.SmtpAddress -eq 'shared@contoso.com' }
        $frank = $shared.Entries | Where-Object { $_.User -eq 'Frank Weber' }
        $frank | Should -Not -BeNullOrEmpty
        $frankCal = $frank.Folders | Where-Object { $_.FolderName -eq 'Calendar' }
        $frankCal | Should -Not -BeNullOrEmpty
    }

    It 'Each folder entry has EntryID and StoreID for removal' {
        $overview = Get-PermissionsOverview
        $alice = $overview | Where-Object { $_.SmtpAddress -eq 'alice@contoso.com' }
        $bob   = $alice.Entries | Where-Object { $_.User -eq 'Bob Smith' }
        foreach ($f in $bob.Folders) {
            $f.EntryID | Should -Not -BeNullOrEmpty
            $f.StoreID | Should -Not -BeNullOrEmpty
        }
    }

    It 'Bob Delegate account has no custom entries (clean slate)' {
        $overview = Get-PermissionsOverview
        $bob = $overview | Where-Object { $_.SmtpAddress -eq 'bob.delegate@contoso.com' }
        $bob.Entries.Count | Should -Be 0
    }
}

Describe 'Subfolder hierarchy in mock' {
    It 'Alice Inbox has Projects subfolder at Depth 2' {
        $folders = Get-MailboxFolders -SmtpAddress 'alice@contoso.com'
        $projects = $folders | Where-Object { $_.Name -eq 'Projects' }
        $projects | Should -Not -BeNullOrEmpty
        $projects.Depth | Should -Be 2
    }

    It 'Projects has Client A subfolder at Depth 3' {
        $folders = Get-MailboxFolders -SmtpAddress 'alice@contoso.com'
        $clientA = $folders | Where-Object { $_.Name -eq 'Client A' }
        $clientA | Should -Not -BeNullOrEmpty
        $clientA.Depth | Should -Be 3
    }

    It 'Alice Inbox has Archive subfolder at Depth 2' {
        $folders = Get-MailboxFolders -SmtpAddress 'alice@contoso.com'
        $archive = $folders | Where-Object { $_.Name -eq 'Archive' }
        $archive | Should -Not -BeNullOrEmpty
        $archive.Depth | Should -Be 2
    }
}

Describe 'Set-FolderPermissionWithAncestors' {
    It 'Depth-1 folder (Sent Items): no auto-grant needed, AutoGranted is empty' {
        $folders = Get-MailboxFolders -SmtpAddress 'alice@contoso.com'
        $sent = $folders | Where-Object { $_.Name -eq 'Sent Items' }

        $result = Set-FolderPermissionWithAncestors -EntryID $sent.EntryID -StoreID $sent.StoreID `
            -User 'Grace Lee' -Level 'Can view'

        $result.Success | Should -BeTrue
        $result.AutoGranted.Count | Should -Be 0
    }

    It 'Depth-2 folder (Projects): auto-grants Can view on Inbox' {
        $folders = Get-MailboxFolders -SmtpAddress 'alice@contoso.com'
        $projects = $folders | Where-Object { $_.Name -eq 'Projects' }

        $result = Set-FolderPermissionWithAncestors -EntryID $projects.EntryID -StoreID $projects.StoreID `
            -User 'Grace Lee' -Level 'Full access'

        $result.Success | Should -BeTrue
        $result.AutoGranted.Count | Should -Be 1
        $result.AutoGranted[0].FolderName | Should -Be 'Inbox'
        $result.AutoGranted[0].Level | Should -Be 'Can view'
    }

    It 'Depth-3 folder (Client A): auto-grants on both Projects and Inbox' {
        $folders = Get-MailboxFolders -SmtpAddress 'alice@contoso.com'
        $clientA = $folders | Where-Object { $_.Name -eq 'Client A' }

        # Remove Grace Lee from Projects and Inbox first to test clean
        Remove-FolderPermission -EntryID ($folders | Where-Object { $_.Name -eq 'Projects' }).EntryID `
            -StoreID $clientA.StoreID -User 'Grace Lee'
        Remove-FolderPermission -EntryID ($folders | Where-Object { $_.Name -eq 'Inbox' }).EntryID `
            -StoreID $clientA.StoreID -User 'Grace Lee'

        $result = Set-FolderPermissionWithAncestors -EntryID $clientA.EntryID -StoreID $clientA.StoreID `
            -User 'Grace Lee' -Level 'Owner'

        $result.Success | Should -BeTrue
        $result.AutoGranted.Count | Should -Be 2
        $grantedNames = $result.AutoGranted | ForEach-Object { $_.FolderName }
        $grantedNames | Should -Contain 'Projects'
        $grantedNames | Should -Contain 'Inbox'
    }

    It 'Does not touch parent if user already has a permission entry (Bob on Inbox)' {
        $folders = Get-MailboxFolders -SmtpAddress 'alice@contoso.com'
        $projects = $folders | Where-Object { $_.Name -eq 'Projects' }
        $inbox    = $folders | Where-Object { $_.Name -eq 'Inbox' }

        # Bob Smith already has "Can view" (level 1) on Inbox
        $result = Set-FolderPermissionWithAncestors -EntryID $projects.EntryID -StoreID $projects.StoreID `
            -User 'Bob Smith' -Level 'Full access'

        $result.Success | Should -BeTrue
        $result.AutoGranted.Count | Should -Be 0

        # Verify Bob's Inbox permission was NOT changed (still level 1)
        $inboxPerms = Get-FolderPermissions -EntryID $inbox.EntryID -StoreID $inbox.StoreID
        $bobPerm = $inboxPerms | Where-Object { $_.User -eq 'Bob Smith' }
        $bobPerm.PermissionLevel | Should -Be 1
    }

    It 'Never auto-grants on the mailbox root (Depth 0)' {
        $folders = Get-MailboxFolders -SmtpAddress 'alice@contoso.com'
        $inbox = $folders | Where-Object { $_.Name -eq 'Inbox' }

        # Eva Mueller has no entry on Inbox — this should NOT create one on the root
        $result = Set-FolderPermissionWithAncestors -EntryID $inbox.EntryID -StoreID $inbox.StoreID `
            -User 'Eva Mueller' -Level 'Can view'

        $result.Success | Should -BeTrue
        # Root (Alice Johnson) should not be in AutoGranted
        $rootGrant = $result.AutoGranted | Where-Object { $_.FolderName -eq 'Alice Johnson' }
        $rootGrant | Should -BeNullOrEmpty
    }
}
