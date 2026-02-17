<#
  Ui.Tests.ps1
  Automated WPF UI tests for the Signatures tab.

  Launches the sandbox app in a background STA runspace, interacts with
  controls via DispatcherTimer commands, and verifies state changes.

  Tests are ordered: create -> rename -> delete to avoid interference.
  A single app instance is shared across all tests for speed.
#>

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot 'UiTestHelper.psm1') -Force
    $script:ctx = Start-SandboxApp
}

AfterAll {
    if ($null -ne $script:ctx) {
        Stop-SandboxApp $script:ctx
    }
}

# ─── Initial Load ──────────────────────────────────────────────────────────────

Describe 'Initial Load' {
    It 'Window is visible and ready' {
        $vis = Invoke-UiCommand $script:ctx '$window.IsVisible'
        $vis | Should -BeTrue
    }

    It 'PanelInboxList contains mailbox cards and signature rows' {
        $childCount = Invoke-UiCommand $script:ctx '$window.FindName("PanelInboxList").Children.Count'
        $childCount | Should -BeGreaterThan 3
    }

    It 'Bottom pill bar has pills for each mailbox' {
        $pillCount = Invoke-UiCommand $script:ctx '$window.FindName("PanelCopyTargets").Children.Count'
        $pillCount | Should -BeGreaterOrEqual 3
    }

    It 'Preview empty hint is visible initially' {
        Get-ControlVisibility $script:ctx 'TxtPreviewEmptyHint' | Should -Be 'Visible'
    }

    It 'Preview browser is collapsed initially' {
        Get-ControlVisibility $script:ctx 'PreviewBrowser' | Should -Be 'Collapsed'
    }

    It 'Edit button is collapsed initially' {
        Get-ControlVisibility $script:ctx 'BtnEditSignature' | Should -Be 'Collapsed'
    }

    It 'Reload preview button is collapsed initially' {
        Get-ControlVisibility $script:ctx 'BtnReloadPreview' | Should -Be 'Collapsed'
    }

    It 'Status bar shows mailbox count' {
        $status = Get-ControlText $script:ctx 'TxtStatus'
        $status | Should -BeLike '*mailbox*'
    }

    It 'All 3 fixture signatures appear in the panel' {
        $names = Get-InboxListSigNames $script:ctx
        $names | Should -Contain 'Corporate Standard'
        $names | Should -Contain 'Personal'
        $names | Should -Contain 'Shared Mailbox Sig'
    }
}

# ─── Preview Loading ───────────────────────────────────────────────────────────

Describe 'Preview Loading' {
    It 'Clicking a signature row shows the preview' {
        Click-SignatureRow $script:ctx 'Corporate Standard'

        $vis = Wait-ForCondition $script:ctx `
            '$window.FindName("PreviewBrowser").Visibility.ToString() -eq "Visible"' `
            -TimeoutMs 3000
        $vis | Should -BeTrue
    }

    It 'Empty hint is hidden after selecting a signature' {
        Get-ControlVisibility $script:ctx 'TxtPreviewEmptyHint' | Should -Be 'Collapsed'
    }

    It 'Signature info shows the selected name' {
        $info = Get-ControlText $script:ctx 'TxtSignatureInfo'
        $info | Should -BeLike '*Corporate Standard*'
    }

    It 'TxtSelectedSig shows the selected name' {
        $label = Get-ControlText $script:ctx 'TxtSelectedSig'
        $label | Should -Be 'Corporate Standard'
    }

    It 'Edit button becomes visible' {
        Get-ControlVisibility $script:ctx 'BtnEditSignature' | Should -Be 'Visible'
    }

    It 'Reload preview button becomes visible' {
        Get-ControlVisibility $script:ctx 'BtnReloadPreview' | Should -Be 'Visible'
    }

    It 'Reload preview updates status bar' {
        Click-Button $script:ctx 'BtnReloadPreview'
        $status = Get-ControlText $script:ctx 'TxtStatus'
        $status | Should -BeLike '*Preview*refresh*'
    }

    It 'Clicking a different signature switches preview' {
        Click-SignatureRow $script:ctx 'Personal'
        $info = Get-ControlText $script:ctx 'TxtSignatureInfo'
        $info | Should -BeLike '*Personal*'
    }
}

# ─── New Signature ─────────────────────────────────────────────────────────────

Describe 'New Signature' {
    It 'Creates a new signature when name is provided' {
        $script:ctx.SyncHash._inputBoxResult = 'Test UI Sig'
        Click-Button $script:ctx 'BtnNewSignature'

        $names = Get-InboxListSigNames $script:ctx
        $names | Should -Contain 'Test UI Sig'
    }

    It 'Shows success status after creation' {
        $status = Get-ControlText $script:ctx 'TxtStatus'
        $status | Should -BeLike "*Created*Test UI Sig*"
    }

    It 'Does nothing when user cancels (empty name)' {
        $beforeNames = Get-InboxListSigNames $script:ctx
        $script:ctx.SyncHash._inputBoxResult = ''
        Click-Button $script:ctx 'BtnNewSignature'
        $afterNames = Get-InboxListSigNames $script:ctx
        $afterNames.Count | Should -Be $beforeNames.Count
    }
}

# ─── Rename Signature ──────────────────────────────────────────────────────────

Describe 'Rename Signature' {
    It 'Requires a signature to be selected' {
        Invoke-UiCommand $script:ctx '$script:selectedSigName = $null'
        $script:ctx.SyncHash._lastError = ''
        Click-Button $script:ctx 'BtnRenameSignature'
        $script:ctx.SyncHash._lastError | Should -BeLike '*Click a signature*'
    }

    It 'Renames the selected signature' {
        Click-SignatureRow $script:ctx 'Test UI Sig'
        $script:ctx.SyncHash._inputBoxResult = 'Renamed Sig'
        Click-Button $script:ctx 'BtnRenameSignature'

        $names = Get-InboxListSigNames $script:ctx
        $names | Should -Contain 'Renamed Sig'
        $names | Should -Not -Contain 'Test UI Sig'
    }

    It 'Updates status bar after rename' {
        $status = Get-ControlText $script:ctx 'TxtStatus'
        $status | Should -BeLike "*Renamed*"
    }

    It 'Updates the selected signature label' {
        $label = Get-ControlText $script:ctx 'TxtSelectedSig'
        $label | Should -Be 'Renamed Sig'
    }
}

# ─── Delete Signature ──────────────────────────────────────────────────────────

Describe 'Delete Signature' {
    It 'Does not delete when user cancels confirmation' {
        Click-SignatureRow $script:ctx 'Renamed Sig'
        $script:ctx.SyncHash._confirmResult = $false
        $beforeNames = Get-InboxListSigNames $script:ctx
        Click-Button $script:ctx 'BtnDeleteSignature'
        $afterNames = Get-InboxListSigNames $script:ctx
        $afterNames.Count | Should -Be $beforeNames.Count
    }

    It 'Deletes the selected signature when confirmed' {
        Click-SignatureRow $script:ctx 'Renamed Sig'
        $script:ctx.SyncHash._confirmResult = $true
        Click-Button $script:ctx 'BtnDeleteSignature'

        $names = Get-InboxListSigNames $script:ctx
        $names | Should -Not -Contain 'Renamed Sig'
    }

    It 'Hides preview after delete' {
        Get-ControlVisibility $script:ctx 'PreviewBrowser' | Should -Be 'Collapsed'
    }

    It 'Shows empty hint after delete' {
        Get-ControlVisibility $script:ctx 'TxtPreviewEmptyHint' | Should -Be 'Visible'
    }

    It 'Resets selected sig label after delete' {
        $label = Get-ControlText $script:ctx 'TxtSelectedSig'
        $label | Should -Be '(none selected)'
    }

    It 'Shows deleted status' {
        $status = Get-ControlText $script:ctx 'TxtStatus'
        $status | Should -BeLike "*Deleted*Renamed Sig*"
    }

    It 'Edit and reload buttons are hidden after delete' {
        Get-ControlVisibility $script:ctx 'BtnEditSignature' | Should -Be 'Collapsed'
        Get-ControlVisibility $script:ctx 'BtnReloadPreview' | Should -Be 'Collapsed'
    }
}

# ─── Assign Signature ─────────────────────────────────────────────────────────

Describe 'Assign Signature' {
    It 'Assigns signature to mailbox via pill click' {
        Click-SignatureRow $script:ctx 'Personal'
        Click-Pill $script:ctx 'Bob (Delegate)'

        # refreshInboxList overwrites status, so verify via badge
        Click-SignatureRow $script:ctx 'Personal'
        $badge = Get-PillBadgeText $script:ctx 'Bob (Delegate)'
        $badge | Should -Be 'New + Reply'
    }

    It 'Pill shows New + Reply badge after assignment' {
        Click-SignatureRow $script:ctx 'Personal'
        $badge = Get-PillBadgeText $script:ctx 'Bob (Delegate)'
        $badge | Should -Be 'New + Reply'
    }
}

# ─── Unassign Signature ───────────────────────────────────────────────────────

Describe 'Unassign Signature' {
    It 'Unassigns signature from mailbox via pill click (toggle off)' {
        Click-SignatureRow $script:ctx 'Personal'
        Click-Pill $script:ctx 'Bob (Delegate)'

        Click-SignatureRow $script:ctx 'Personal'
        $badge = Get-PillBadgeText $script:ctx 'Bob (Delegate)'
        $badge | Should -Be ''
    }

    It 'Pill badge is gone after unassignment' {
        Click-SignatureRow $script:ctx 'Personal'
        $badge = Get-PillBadgeText $script:ctx 'Bob (Delegate)'
        $badge | Should -Be ''
    }
}
