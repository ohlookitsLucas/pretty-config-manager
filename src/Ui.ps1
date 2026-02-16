function Initialize-Ui {
    param(
        [Parameter(Mandatory=$true)] $Window,
        [Parameter(Mandatory=$true)] [string] $ScriptRoot
    )

    # Find controls
    $lbSignatures = $Window.FindName('LbSignatures')
    $btnCopy = $Window.FindName('BtnCopySignature')
    $btnExport = $Window.FindName('BtnExportSignature')
    $btnImport = $Window.FindName('BtnImportSignature')
    $preview = $Window.FindName('PreviewBrowser')
    $txtInfo = $Window.FindName('TxtSignatureInfo')
    $txtStatus = $Window.FindName('TxtStatus')

    $lbAccounts = $Window.FindName('LbAccounts')
    $btnRefreshAccounts = $Window.FindName('BtnRefreshAccounts')
    $btnSetPermission = $Window.FindName('BtnSetPermission')
    $txtPermissionUser = $Window.FindName('TxtPermissionUser')
    $cbPermissionLevel = $Window.FindName('CbPermissionLevel')
    $lbPermissions = $Window.FindName('LbPermissions')

    # Set WebBrowser to dark background on load so it's not a white flash
    try {
        $preview.NavigateToString('<html><body style="background:#1E1E2E;margin:0;padding:0;"></body></html>')
    } catch {}

    # Reusable: refresh signature list
    $refreshSignatures = {
        $lbSignatures.Items.Clear()
        foreach ($s in (Get-Signatures)) { $lbSignatures.Items.Add($s) }
    }

    # Reusable: refresh accounts list
    $refreshAccounts = {
        $accounts = Get-SignedInAccounts
        $lbAccounts.Items.Clear()
        foreach ($a in $accounts) { $lbAccounts.Items.Add( "$($a.Name) <$($a.SmtpAddress)>" ) }
        return $accounts
    }

    # Load initial data
    try {
        & $refreshSignatures
        $txtStatus.Text = "Loaded $($lbSignatures.Items.Count) signatures"
    } catch {
        $txtStatus.Text = "Error loading signatures: $_"
    }

    try {
        $accounts = & $refreshAccounts
        $txtStatus.Text = "Loaded $($accounts.Count) accounts"
    } catch {
        $txtStatus.Text = "Error loading accounts: $_"
    }

    # Signature selection -> preview
    $lbSignatures.add_SelectionChanged({
        $sel = $lbSignatures.SelectedItem
        if ($null -ne $sel) {
            $htmlPath = Get-SignatureHtmlPath -Name $sel
            if (Test-Path $htmlPath) {
                try {
                    $preview.Navigate((New-Object System.Uri($htmlPath)))
                    $txtInfo.Text = "Previewing: $sel"
                } catch {
                    $txtInfo.Text = "Unable to preview: $_"
                }
            } else {
                $txtInfo.Text = "HTML file not found for $sel"
            }
        }
    })

    # Copy signature
    $btnCopy.Add_Click({
        $sel = $lbSignatures.SelectedItem
        if ($null -eq $sel) { [System.Windows.MessageBox]::Show('Select a signature first','Info') | Out-Null; return }

        Add-Type -AssemblyName Microsoft.VisualBasic
        $target = [Microsoft.VisualBasic.Interaction]::InputBox("Enter target signature name (e.g. account alias):","Copy Signature","$sel")
        if ([string]::IsNullOrWhiteSpace($target)) { return }

        try {
            Copy-Signature -SourceName $sel -TargetName $target
            & $refreshSignatures
            $txtStatus.Text = "Copied '$sel' -> '$target'"
        } catch {
            [System.Windows.MessageBox]::Show("Failed to copy: $_","Error") | Out-Null
            $txtStatus.Text = "Copy failed"
        }
    })

    # Export signature
    $btnExport.Add_Click({
        $sel = $lbSignatures.SelectedItem
        if ($null -eq $sel) { [System.Windows.MessageBox]::Show('Select a signature first','Info') | Out-Null; return }

        $artifactsDir = Join-Path $ScriptRoot 'artifacts'
        if (-not (Test-Path $artifactsDir)) { New-Item -ItemType Directory -Path $artifactsDir | Out-Null }
        $zipPath = Join-Path $artifactsDir "$sel.zip"

        try {
            Export-Signature -Name $sel -Destination $zipPath
            [System.Windows.MessageBox]::Show("Exported to $zipPath","Export") | Out-Null
            $txtStatus.Text = "Exported $sel"
        } catch {
            [System.Windows.MessageBox]::Show("Failed to export: $_","Error") | Out-Null
            $txtStatus.Text = "Export failed"
        }
    })

    # Import signature
    $btnImport.Add_Click({
        Add-Type -AssemblyName System.Windows.Forms
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = 'Zip files (*.zip)|*.zip|All files (*.*)|*.*'
        if ($ofd.ShowDialog() -eq 'OK') {
            try {
                Import-Signature -ZipPath $ofd.FileName
                & $refreshSignatures
                $txtStatus.Text = "Imported signature"
            } catch {
                [System.Windows.MessageBox]::Show("Failed to import: $_","Error") | Out-Null
                $txtStatus.Text = "Import failed"
            }
        }
    })

    # Refresh accounts
    $btnRefreshAccounts.Add_Click({
        try {
            & $refreshAccounts
            $txtStatus.Text = "Refreshed accounts"
        } catch {
            $txtStatus.Text = "Error refreshing accounts: $_"
        }
    })

    # Set permission
    $btnSetPermission.Add_Click({
        $targetUser = $txtPermissionUser.Text.Trim()
        $level = $cbPermissionLevel.SelectedItem
        if ([string]::IsNullOrWhiteSpace($targetUser) -or $null -eq $level) {
            [System.Windows.MessageBox]::Show('Enter user and select level','Info') | Out-Null
            return
        }
        $levelText = $level.Content

        if ([System.Windows.MessageBox]::Show("Set permission '$levelText' for $targetUser?","Confirm",[System.Windows.MessageBoxButton]::YesNo) -ne [System.Windows.MessageBoxResult]::Yes) { return }

        try {
            Set-CalendarPermission -User $targetUser -Level $levelText -Confirm:$false
            [System.Windows.MessageBox]::Show("Permission set (or queued). Check Status for details.","Done") | Out-Null
            $txtStatus.Text = "Permission set for $($targetUser): $($levelText)"
        } catch {
            [System.Windows.MessageBox]::Show("Failed to set permission: $_","Error") | Out-Null
            $txtStatus.Text = "Permission set failed"
        }
    })
}
