Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ----------------------------
# Form
# ----------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Pigeon Transfer"
$form.Size = New-Object System.Drawing.Size(620, 620)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false

# ----------------------------
# Section: Files
# ----------------------------
$grpFiles = New-Object System.Windows.Forms.GroupBox
$grpFiles.Text = "Files / Folders to Pack  [0 items]"
$grpFiles.Size = New-Object System.Drawing.Size(590, 200)
$grpFiles.Location = New-Object System.Drawing.Point(10, 8)
$form.Controls.Add($grpFiles)

$lbFiles = New-Object System.Windows.Forms.ListBox
$lbFiles.Size = New-Object System.Drawing.Size(460, 155)
$lbFiles.Location = New-Object System.Drawing.Point(8, 20)
$lbFiles.AllowDrop = $true
$lbFiles.SelectionMode = "MultiExtended"
$grpFiles.Controls.Add($lbFiles)

# Drag & Drop into list box
$lbFiles.Add_DragEnter({
    param($s, $e)
    if ($e.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
    }
})
$lbFiles.Add_DragDrop({
    param($s, $e)
    foreach ($p in $e.Data.GetData([Windows.Forms.DataFormats]::FileDrop)) {
        if (-not $lbFiles.Items.Contains($p)) { $lbFiles.Items.Add($p) }
    }
    Update-FileCount
})

# Sidebar buttons for list management
$btnBrowseFolder = New-Object System.Windows.Forms.Button
$btnBrowseFolder.Text = "Add Folder"
$btnBrowseFolder.Size = New-Object System.Drawing.Size(105, 30)
$btnBrowseFolder.Location = New-Object System.Drawing.Point(476, 20)
$grpFiles.Controls.Add($btnBrowseFolder)

$btnBrowseFile = New-Object System.Windows.Forms.Button
$btnBrowseFile.Text = "Add File(s)"
$btnBrowseFile.Size = New-Object System.Drawing.Size(105, 30)
$btnBrowseFile.Location = New-Object System.Drawing.Point(476, 58)
$grpFiles.Controls.Add($btnBrowseFile)

$btnRemove = New-Object System.Windows.Forms.Button
$btnRemove.Text = "Remove"
$btnRemove.Size = New-Object System.Drawing.Size(105, 30)
$btnRemove.Location = New-Object System.Drawing.Point(476, 96)
$grpFiles.Controls.Add($btnRemove)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "Clear All"
$btnClear.Size = New-Object System.Drawing.Size(105, 30)
$btnClear.Location = New-Object System.Drawing.Point(476, 134)
$grpFiles.Controls.Add($btnClear)

# ----------------------------
# Section: Actions
# ----------------------------
$grpActions = New-Object System.Windows.Forms.GroupBox
$grpActions.Text = "Pack / Unpack"
$grpActions.Size = New-Object System.Drawing.Size(590, 60)
$grpActions.Location = New-Object System.Drawing.Point(10, 215)
$form.Controls.Add($grpActions)

$btnPack = New-Object System.Windows.Forms.Button
$btnPack.Text = "Pack -> TXT"
$btnPack.Size = New-Object System.Drawing.Size(120, 30)
$btnPack.Location = New-Object System.Drawing.Point(8, 20)
$grpActions.Controls.Add($btnPack)

$btnUnpack = New-Object System.Windows.Forms.Button
$btnUnpack.Text = "Unpack TXT -> Files"
$btnUnpack.Size = New-Object System.Drawing.Size(140, 30)
$btnUnpack.Location = New-Object System.Drawing.Point(136, 20)
$grpActions.Controls.Add($btnUnpack)

# ----------------------------
# Section: Email
# ----------------------------
$grpEmail = New-Object System.Windows.Forms.GroupBox
$grpEmail.Text = "Send via Email"
$grpEmail.Size = New-Object System.Drawing.Size(590, 220)
$grpEmail.Location = New-Object System.Drawing.Point(10, 283)
$form.Controls.Add($grpEmail)

# Method selector
$lblMethod = New-Object System.Windows.Forms.Label
$lblMethod.Text = "Method:"
$lblMethod.Size = New-Object System.Drawing.Size(60, 20)
$lblMethod.Location = New-Object System.Drawing.Point(8, 22)
$grpEmail.Controls.Add($lblMethod)

$cmbMethod = New-Object System.Windows.Forms.ComboBox
$cmbMethod.Size = New-Object System.Drawing.Size(160, 24)
$cmbMethod.Location = New-Object System.Drawing.Point(70, 20)
$cmbMethod.DropDownStyle = "DropDownList"
[void]$cmbMethod.Items.Add("Gmail (SMTP)")
[void]$cmbMethod.Items.Add("Outlook (COM)")
$cmbMethod.SelectedIndex = 0
$grpEmail.Controls.Add($cmbMethod)

# Recipient
$lblTo = New-Object System.Windows.Forms.Label
$lblTo.Text = "Recipient:"
$lblTo.Size = New-Object System.Drawing.Size(65, 20)
$lblTo.Location = New-Object System.Drawing.Point(8, 54)
$grpEmail.Controls.Add($lblTo)

$txtTo = New-Object System.Windows.Forms.TextBox
$txtTo.Size = New-Object System.Drawing.Size(300, 24)
$txtTo.Location = New-Object System.Drawing.Point(76, 52)
$grpEmail.Controls.Add($txtTo)

# Gmail sender (hidden for Outlook)
$lblFrom = New-Object System.Windows.Forms.Label
$lblFrom.Text = "Gmail address:"
$lblFrom.Size = New-Object System.Drawing.Size(90, 20)
$lblFrom.Location = New-Object System.Drawing.Point(8, 86)
$grpEmail.Controls.Add($lblFrom)

$txtFrom = New-Object System.Windows.Forms.TextBox
$txtFrom.Size = New-Object System.Drawing.Size(220, 24)
$txtFrom.Location = New-Object System.Drawing.Point(100, 84)
$grpEmail.Controls.Add($txtFrom)

# App password (hidden for Outlook)
$lblPass = New-Object System.Windows.Forms.Label
$lblPass.Text = "App Password:"
$lblPass.Size = New-Object System.Drawing.Size(90, 20)
$lblPass.Location = New-Object System.Drawing.Point(8, 118)
$grpEmail.Controls.Add($lblPass)

$txtPass = New-Object System.Windows.Forms.TextBox
$txtPass.Size = New-Object System.Drawing.Size(220, 24)
$txtPass.Location = New-Object System.Drawing.Point(100, 116)
$txtPass.UseSystemPasswordChar = $true
$grpEmail.Controls.Add($txtPass)

$lnkHelp = New-Object System.Windows.Forms.LinkLabel
$lnkHelp.Text = "How to get a Gmail App Password"
$lnkHelp.Size = New-Object System.Drawing.Size(200, 18)
$lnkHelp.Location = New-Object System.Drawing.Point(328, 120)
$lnkHelp.Add_LinkClicked({
    Start-Process "https://myaccount.google.com/apppasswords"
})
$grpEmail.Controls.Add($lnkHelp)

# Outlook mailbox selector (hidden for Gmail)
$lblMailbox = New-Object System.Windows.Forms.Label
$lblMailbox.Text = "Send from:"
$lblMailbox.Size = New-Object System.Drawing.Size(70, 20)
$lblMailbox.Location = New-Object System.Drawing.Point(8, 86)
$lblMailbox.Visible = $false
$grpEmail.Controls.Add($lblMailbox)

$cmbMailbox = New-Object System.Windows.Forms.ComboBox
$cmbMailbox.Size = New-Object System.Drawing.Size(300, 24)
$cmbMailbox.Location = New-Object System.Drawing.Point(80, 84)
$cmbMailbox.DropDownStyle = "DropDownList"
$cmbMailbox.Visible = $false
$grpEmail.Controls.Add($cmbMailbox)

$btnRefreshMailboxes = New-Object System.Windows.Forms.Button
$btnRefreshMailboxes.Text = "Refresh"
$btnRefreshMailboxes.Size = New-Object System.Drawing.Size(70, 24)
$btnRefreshMailboxes.Location = New-Object System.Drawing.Point(388, 84)
$btnRefreshMailboxes.Visible = $false
$grpEmail.Controls.Add($btnRefreshMailboxes)

# Pack + Send button
$btnPackSend = New-Object System.Windows.Forms.Button
$btnPackSend.Text = "Pack & Send"
$btnPackSend.Size = New-Object System.Drawing.Size(120, 30)
$btnPackSend.Location = New-Object System.Drawing.Point(460, 175)
$btnPackSend.BackColor = [System.Drawing.Color]::SteelBlue
$btnPackSend.ForeColor = [System.Drawing.Color]::White
$btnPackSend.FlatStyle = "Flat"
$grpEmail.Controls.Add($btnPackSend)

# Helper: populate Outlook mailbox list
function Refresh-OutlookMailboxes {
    $cmbMailbox.Items.Clear()
    try {
        $ol = New-Object -ComObject Outlook.Application
        foreach ($acct in $ol.Session.Accounts) {
            [void]$cmbMailbox.Items.Add($acct.SmtpAddress)
        }
        if ($cmbMailbox.Items.Count -gt 0) { $cmbMailbox.SelectedIndex = 0 }
        else { [void]$cmbMailbox.Items.Add("(no accounts found)") ; $cmbMailbox.SelectedIndex = 0 }
    } catch {
        [void]$cmbMailbox.Items.Add("(Outlook not available)")
        $cmbMailbox.SelectedIndex = 0
    }
}

$btnRefreshMailboxes.Add_Click({ Refresh-OutlookMailboxes })

# Toggle Gmail vs Outlook fields based on method
$cmbMethod.Add_SelectedIndexChanged({
    $isGmail = ($cmbMethod.SelectedIndex -eq 0)
    $lblFrom.Visible   = $isGmail
    $txtFrom.Visible   = $isGmail
    $lblPass.Visible   = $isGmail
    $txtPass.Visible   = $isGmail
    $lnkHelp.Visible   = $isGmail
    $lblMailbox.Visible        = -not $isGmail
    $cmbMailbox.Visible        = -not $isGmail
    $btnRefreshMailboxes.Visible = -not $isGmail
    if (-not $isGmail -and $cmbMailbox.Items.Count -eq 0) { Refresh-OutlookMailboxes }
})

# Watermark helper for plain TextBoxes (PlaceholderText not available in Windows PS)
# GetNewClosure() ensures $textbox and $hint are captured by value in PS5
function Add-Watermark($textbox, $hint) {
    $textbox.ForeColor = [System.Drawing.Color]::Gray
    $textbox.Text = $hint
    $textbox.Add_GotFocus(({
        if ($textbox.Text -eq $hint -and $textbox.ForeColor -eq [System.Drawing.Color]::Gray) {
            $textbox.Text = ""
            $textbox.ForeColor = [System.Drawing.Color]::Black
        }
    }).GetNewClosure())
    $textbox.Add_LostFocus(({
        if ($textbox.Text -eq "") {
            $textbox.ForeColor = [System.Drawing.Color]::Gray
            $textbox.Text = $hint
        }
    }).GetNewClosure())
}

# Returns the real text of a watermarked textbox, or "" if it still shows the hint
function Get-WatermarkText($textbox) {
    if ($textbox.ForeColor -eq [System.Drawing.Color]::Gray) { return "" }
    return $textbox.Text.Trim()
}

Add-Watermark $txtTo   "recipient@example.com"
Add-Watermark $txtFrom "you@gmail.com"

# For password field: show a grey hint label instead (can't show text while masked)
$lblPassHint = New-Object System.Windows.Forms.Label
$lblPassHint.Text = "App Password (16 chars, no spaces)"
$lblPassHint.ForeColor = [System.Drawing.Color]::Gray
$lblPassHint.Size = New-Object System.Drawing.Size(210, 20)
$lblPassHint.Location = New-Object System.Drawing.Point(100, 140)
$grpEmail.Controls.Add($lblPassHint)

$txtPass.Add_TextChanged({
    $lblPassHint.Visible = ($txtPass.Text.Length -eq 0)
})

# ----------------------------
# Status / Progress
# ----------------------------
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Size = New-Object System.Drawing.Size(590, 40)
$lblStatus.Location = New-Object System.Drawing.Point(10, 511)
$lblStatus.BorderStyle = "Fixed3D"
$lblStatus.TextAlign = "MiddleLeft"
$lblStatus.Text = "Ready."
$form.Controls.Add($lblStatus)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Size = New-Object System.Drawing.Size(590, 18)
$progress.Location = New-Object System.Drawing.Point(10, 558)
$progress.Minimum = 0
$progress.Value = 0
$form.Controls.Add($progress)

# ----------------------------
# Helper: Update file list count in group header
# ----------------------------
function Update-FileCount {
    $grpFiles.Text = "Files / Folders to Pack  [$($lbFiles.Items.Count) item$(if($lbFiles.Items.Count -ne 1){'s'})]"
}

# ----------------------------
# Helper: Set status
# ----------------------------
function Set-Status($msg, [System.Drawing.Color]$color = [System.Drawing.Color]::Black) {
    $lblStatus.ForeColor = $color
    $lblStatus.Text = $msg
    [System.Windows.Forms.Application]::DoEvents()
}

# ----------------------------
# Helper: Collect all files from list
# ----------------------------
function Get-AllFiles {
    $allFiles = @()
    foreach ($item in $lbFiles.Items) {
        if (Test-Path $item -PathType Leaf) {
            $allFiles += $item
        } elseif (Test-Path $item -PathType Container) {
            $allFiles += Get-ChildItem $item -Recurse -File | ForEach-Object { $_.FullName }
        }
    }
    return $allFiles
}

# ----------------------------
# Function: Pack to a given output path
# Returns $true on success
# ----------------------------
function Pack-ToFile($outfile) {
    $allFiles = Get-AllFiles
    if ($allFiles.Count -eq 0) { Set-Status "No files/folders to pack!" ([System.Drawing.Color]::OrangeRed); return $false }

    $progress.Maximum = $allFiles.Count
    $progress.Value = 0
    Set-Status "Packing $($allFiles.Count) file(s)..."

    $data = @()
    foreach ($file in $allFiles) {
        $relPath = $file
        foreach ($root in $lbFiles.Items) {
            if ((Test-Path $root -PathType Container) -and $file.StartsWith($root)) {
                $relPath = Join-Path (Split-Path $root -Leaf) $file.Substring($root.Length).TrimStart("\")
            }
        }
        $bytes = [IO.File]::ReadAllBytes($file)
        $data += [PSCustomObject]@{ Path = $relPath; Data = [Convert]::ToBase64String($bytes) }
        $progress.Value += 1
        [System.Windows.Forms.Application]::DoEvents()
    }

    $data | ConvertTo-Json -Depth 10 | Set-Content -Path $outfile -Encoding UTF8
    $progress.Value = 0
    return $true
}

# ----------------------------
# Function: Pack (save dialog)
# ----------------------------
function Pack-Files {
    if ($lbFiles.Items.Count -eq 0) { Set-Status "No files/folders to pack!" ([System.Drawing.Color]::OrangeRed); return }

    $save = New-Object System.Windows.Forms.SaveFileDialog
    $save.Filter = "Text File (*.txt)|*.txt"
    $save.FileName = "pigeon_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    if ($save.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    Set-Busy $true
    try {
        if (Pack-ToFile $save.FileName) {
            Set-Status "Packed to: $($save.FileName)" ([System.Drawing.Color]::DarkGreen)
        }
    } finally { Set-Busy $false }
}

# ----------------------------
# Function: Unpack
# ----------------------------
function Unpack-Files {
    $open = New-Object System.Windows.Forms.OpenFileDialog
    $open.Filter = "Text File (*.txt)|*.txt"
    if ($open.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $folder = New-Object System.Windows.Forms.FolderBrowserDialog
    $folder.Description = "Select destination folder for unpacked files"
    if ($folder.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    Set-Busy $true
    try {
        $data = Get-Content $open.FileName -Raw | ConvertFrom-Json
        $progress.Maximum = $data.Count
        $progress.Value = 0
        Set-Status "Unpacking $($data.Count) file(s)..."

        foreach ($f in $data) {
            $target = Join-Path $folder.SelectedPath $f.Path
            $dir = Split-Path $target
            if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            [IO.File]::WriteAllBytes($target, [Convert]::FromBase64String($f.Data))
            $progress.Value += 1
            [System.Windows.Forms.Application]::DoEvents()
        }

        Set-Status "Unpacked $($data.Count) file(s) to: $($folder.SelectedPath)" ([System.Drawing.Color]::DarkGreen)
        $progress.Value = 0
    } finally { Set-Busy $false }
}

# ----------------------------
# Function: Send via Gmail SMTP
# ----------------------------
function Send-ViaGmail($attachmentPath, $toAddress, $fromAddress, $appPassword) {
    Set-Status "Connecting to Gmail SMTP..."

    $subject = "Packed Files - $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
    $body    = "Please find the packed file bundle attached.`n`nSent by Pigeon Transfer."

    $smtp = New-Object System.Net.Mail.SmtpClient("smtp.gmail.com", 587)
    $smtp.EnableSsl   = $true
    $smtp.Credentials = New-Object System.Net.NetworkCredential($fromAddress, $appPassword)

    $msg = New-Object System.Net.Mail.MailMessage
    $msg.From       = $fromAddress
    $msg.To.Add($toAddress)
    $msg.Subject    = $subject
    $msg.Body       = $body
    $msg.Attachments.Add((New-Object System.Net.Mail.Attachment($attachmentPath)))

    try {
        $smtp.Send($msg)
        return $true
    } catch {
        throw $_
    } finally {
        $msg.Dispose()
        $smtp.Dispose()
    }
}

# ----------------------------
# Function: Send via Outlook COM
# ----------------------------
function Send-ViaOutlook($attachmentPath, $toAddress, $fromSmtp) {
    Set-Status "Creating Outlook mail item..."

    try {
        $outlook = New-Object -ComObject Outlook.Application
    } catch {
        throw "Outlook is not installed or could not be started. $_"
    }

    $mail         = $outlook.CreateItem(0)  # olMailItem
    $mail.Subject = "Packed Files - $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
    $mail.Body    = "Please find the packed file bundle attached.`n`nSent by Pigeon Transfer."
    $mail.To      = $toAddress

    # Set specific sending account if one was selected
    if ($fromSmtp) {
        foreach ($acct in $outlook.Session.Accounts) {
            if ($acct.SmtpAddress -eq $fromSmtp) {
                $mail.SendUsingAccount = $acct
                break
            }
        }
    }

    $mail.Attachments.Add($attachmentPath) | Out-Null

    try {
        $mail.Send()
        return $true
    } catch {
        throw $_
    }
}

# ----------------------------
# Helper: lock/unlock action buttons during long operations
# ----------------------------
function Set-Busy($busy) {
    $btnPack.Enabled     = -not $busy
    $btnUnpack.Enabled   = -not $busy
    $btnPackSend.Enabled = -not $busy
    $form.Cursor = if ($busy) { [System.Windows.Forms.Cursors]::WaitCursor } else { [System.Windows.Forms.Cursors]::Default }
    [System.Windows.Forms.Application]::DoEvents()
}

# ----------------------------
# Function: Pack & Send
# ----------------------------
function Pack-AndSend {
    # Validate list
    if ($lbFiles.Items.Count -eq 0) {
        Set-Status "No files/folders to pack!" ([System.Drawing.Color]::OrangeRed)
        return
    }

    # Validate recipient (use Get-WatermarkText so hint text isn't treated as input)
    $toAddress = Get-WatermarkText $txtTo
    if (-not $toAddress -or $toAddress -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
        Set-Status "Enter a valid recipient email address." ([System.Drawing.Color]::OrangeRed)
        $txtTo.Focus()
        return
    }

    # Validate Gmail credentials if needed
    $useGmail = ($cmbMethod.SelectedIndex -eq 0)
    if ($useGmail) {
        $fromAddress = Get-WatermarkText $txtFrom
        $appPassword = $txtPass.Text.Trim()
        if (-not $fromAddress -or $fromAddress -notmatch '@gmail\.com$') {
            Set-Status "Enter a valid Gmail address." ([System.Drawing.Color]::OrangeRed)
            $txtFrom.Focus()
            return
        }
        if (-not $appPassword) {
            Set-Status "Enter your Gmail App Password." ([System.Drawing.Color]::OrangeRed)
            $txtPass.Focus()
            return
        }
    }

    Set-Busy $true
    # Pack to temp file
    $tempFile = Join-Path $env:TEMP "pigeon_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    try {
        if (-not (Pack-ToFile $tempFile)) { return }

        # Send
        Set-Status "Sending email..."
        if ($useGmail) {
            Send-ViaGmail $tempFile $toAddress $fromAddress $appPassword | Out-Null
        } else {
            $selectedMailbox = if ($cmbMailbox.SelectedItem) { $cmbMailbox.SelectedItem.ToString() } else { $null }
            Send-ViaOutlook $tempFile $toAddress $selectedMailbox | Out-Null
        }
        Set-Status "Pigeon delivered! Email sent to $toAddress" ([System.Drawing.Color]::DarkGreen)
    } catch {
        Set-Status "Send failed: $_" ([System.Drawing.Color]::Red)
    } finally {
        Set-Busy $false
        if (Test-Path $tempFile) { Remove-Item $tempFile -Force }
    }
}

# ----------------------------
# Button events
# ----------------------------
$btnBrowseFolder.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Select a folder to add"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        if (-not $lbFiles.Items.Contains($dlg.SelectedPath)) {
            $lbFiles.Items.Add($dlg.SelectedPath)
            Update-FileCount
        }
    }
})

$btnBrowseFile.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Multiselect = $true
    $dlg.Title = "Select file(s) to add"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        foreach ($f in $dlg.FileNames) {
            if (-not $lbFiles.Items.Contains($f)) { $lbFiles.Items.Add($f) }
        }
        Update-FileCount
    }
})

$btnRemove.Add_Click({
    $selected = @($lbFiles.SelectedItems)
    foreach ($item in $selected) { $lbFiles.Items.Remove($item) }
    Update-FileCount
})

$btnClear.Add_Click({
    $lbFiles.Items.Clear()
    Update-FileCount
    Set-Status "List cleared."
    $progress.Value = 0
})

$btnPack.Add_Click({ Pack-Files })
$btnUnpack.Add_Click({ Unpack-Files })
$btnPackSend.Add_Click({ Pack-AndSend })

# ----------------------------
# Launch
# ----------------------------
[void]$form.ShowDialog()
