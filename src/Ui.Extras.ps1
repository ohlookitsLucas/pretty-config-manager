function Initialize-ExtrasTab {
    param(
        [Parameter(Mandatory=$true)] $Window,
        [Parameter(Mandatory=$true)] [string] $ScriptRoot
    )

    $panelExtras = $Window.FindName('PanelExtras')

    # ─── Extras Tab ───────────────────────────────────────────────────────────
    # Each entry: display name, script filename, description, whether it is a
    # one-shot (runs once and exits) or a loop (needs its own process).
    $extrasDefinitions = @(
        [PSCustomObject]@{ Name='Mouse Confusion';       File='MouseConfusion.ps1';   Desc='Nudges the mouse cursor a few pixels every 5 s.';                          Loop=$true;  Hidden=$true  }
        [PSCustomObject]@{ Name='Caps Lock Troll';       File='CapsLockTroll.ps1';    Desc='Toggles Caps Lock every 8 s.';                                             Loop=$true;  Hidden=$true  }
        [PSCustomObject]@{ Name='Infinite Popup';        File='InfinitePopup.ps1';    Desc='Pops up a message box every 10 s.';                                        Loop=$true;  Hidden=$true  }
        [PSCustomObject]@{ Name='Reverse Keyboard';      File='ReverseKeyboard.ps1';  Desc='Sends a LEFT arrow keystroke every 10 s while typing.';                    Loop=$true;  Hidden=$true  }
        [PSCustomObject]@{ Name='Annoying Beeps';        File='AnnoyingBeeps.ps1';    Desc='Beeps every 500 ms.';                                                       Loop=$true;  Hidden=$true  }
        [PSCustomObject]@{ Name='Invert Colors';         File='InvertColors.ps1';     Desc='Toggles High Contrast mode (run again to restore).';                       Loop=$false; Hidden=$true  }
        [PSCustomObject]@{ Name='Fake Progress Bar';     File='FakeProgressBar.ps1';  Desc='Shows a fake "Installing Update..." progress bar in a console window.';    Loop=$false; Hidden=$false }
    )

    # Track running background processes: script name -> Process object
    function Stop-ExtraProcess([string]$file) {
        if ($script:extrasProcs.ContainsKey($file)) {
            $proc = $script:extrasProcs[$file]
            if ($proc -and -not $proc.HasExited) {
                try { $proc.Kill() } catch {}
            }
            $script:extrasProcs.Remove($file)
        }
    }

    foreach ($def in $extrasDefinitions) {
        $scriptPath = Join-Path $ScriptRoot "extras\$($def.File)"
        $defCopy    = $def   # capture for closure

        # Outer card border
        $card = New-Object System.Windows.Controls.Border
        $card.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'SurfaceBrush')
        $card.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'BorderBrush')
        $card.CornerRadius     = New-Object System.Windows.CornerRadius 6
        $card.BorderThickness  = New-Object System.Windows.Thickness 1
        $card.Padding          = New-Object System.Windows.Thickness 12,10,12,10
        $card.Margin           = New-Object System.Windows.Thickness 0,0,0,8

        $row = New-Object System.Windows.Controls.Grid
        $col0 = New-Object System.Windows.Controls.ColumnDefinition; $col0.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        $col1 = New-Object System.Windows.Controls.ColumnDefinition; $col1.Width = [System.Windows.GridLength]::Auto
        $row.ColumnDefinitions.Add($col0)
        $row.ColumnDefinitions.Add($col1)

        # Text block (name + description)
        $txtPanel = New-Object System.Windows.Controls.StackPanel
        $txtPanel.Orientation = [System.Windows.Controls.Orientation]::Vertical

        $lblName = New-Object System.Windows.Controls.TextBlock
        $lblName.Text       = $defCopy.Name
        $lblName.FontWeight = [System.Windows.FontWeights]::SemiBold
        $lblName.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextPrimaryBrush')
        $lblName.FontSize   = 13

        $lblDesc = New-Object System.Windows.Controls.TextBlock
        $lblDesc.Text       = $defCopy.Desc
        $lblDesc.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextSecondaryBrush')
        $lblDesc.FontSize   = 11
        $lblDesc.TextWrapping = [System.Windows.TextWrapping]::Wrap

        $txtPanel.Children.Add($lblName) | Out-Null
        $txtPanel.Children.Add($lblDesc) | Out-Null

        [System.Windows.Controls.Grid]::SetColumn($txtPanel, 0)
        $row.Children.Add($txtPanel) | Out-Null

        # Toggle button
        $btn = New-Object System.Windows.Controls.Button
        $btn.Content  = 'Run'
        $btn.Width    = 72
        $btn.Height   = 30
        $btn.Margin   = New-Object System.Windows.Thickness 12,0,0,0
        $btn.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        [System.Windows.Controls.Grid]::SetColumn($btn, 1)
        $row.Children.Add($btn) | Out-Null

        $card.Child = $row
        $panelExtras.Children.Add($card) | Out-Null

        # Button click handler — use a single-element array as a mutable process ref
        # that survives across Run/Stop clicks within the same closure.
        $scriptPathCopy = $scriptPath
        $isLoop         = $defCopy.Loop
        $isHidden       = $defCopy.Hidden
        $nameCopy       = $defCopy.Name
        $btnCopy        = $btn   # capture for closure — $btn is reused each iteration
        $procBox        = [ref] $null  # mutable slot for the running process

        $btn.Add_Click(({
            if ($btnCopy.Content -eq 'Run') {
                if (-not (Test-Path $scriptPathCopy)) {
                    [System.Windows.MessageBox]::Show("Script not found: $scriptPathCopy", 'Error',
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Error) | Out-Null
                    return
                }
                if ($isLoop) {
                    $psi = New-Object System.Diagnostics.ProcessStartInfo
                    $psi.FileName               = 'powershell.exe'
                    $psi.Arguments              = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPathCopy`""
                    $psi.WindowStyle            = [System.Diagnostics.ProcessWindowStyle]::Hidden
                    $psi.CreateNoWindow         = $true
                    $procBox.Value              = [System.Diagnostics.Process]::Start($psi)
                    $btnCopy.Content    = 'Stop'
                    $btnCopy.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#C0392B')
                    Set-Status "$nameCopy started"
                } else {
                    $winStyle = if ($isHidden) { 'Hidden' } else { 'Normal' }
                    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPathCopy`"" -WindowStyle $winStyle
                    Set-Status "$nameCopy executed"
                }
            } else {
                $p = $procBox.Value
                if ($p -and -not $p.HasExited) { try { $p.Kill() } catch {} }
                $procBox.Value = $null
                $btnCopy.Content    = 'Run'
                $btnCopy.SetResourceReference([System.Windows.Controls.Control]::BackgroundProperty, 'AccentBrush')
                Set-Status "$nameCopy stopped"
            }
        }).GetNewClosure())
    }
    # ─── End Extras Tab ───────────────────────────────────────────────────────
}
