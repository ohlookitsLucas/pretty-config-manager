<#
  capture-window.ps1
  Takes screenshots of each tab in the outlookmAnAger window.
  Uses PrintWindow API - does NOT steal focus or bring window to front.
  Output: artifacts\screenshot-N-tabname.png (numbered sequentially)
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition @"
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text;

public class Win32Cap {
    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    // PrintWindow captures a window without requiring it to be in the foreground
    [DllImport("user32.dll")]
    public static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, uint nFlags);

    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    // PW_RENDERFULLCONTENT = 2 (captures DWM-composed content)
    public static Bitmap CaptureWindow(IntPtr hWnd) {
        RECT rect;
        GetWindowRect(hWnd, out rect);
        int width = rect.Right - rect.Left;
        int height = rect.Bottom - rect.Top;
        if (width <= 0 || height <= 0) return null;

        Bitmap bmp = new Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        using (Graphics g = Graphics.FromImage(bmp)) {
            IntPtr hdc = g.GetHdc();
            PrintWindow(hWnd, hdc, 2); // PW_RENDERFULLCONTENT
            g.ReleaseHdc(hdc);
        }
        return bmp;
    }
}
"@

# Find the outlookmAnAger window
$targetHwnd = [IntPtr]::Zero

$callback = [Win32Cap+EnumWindowsProc]{
    param([IntPtr]$hWnd, [IntPtr]$lParam)
    if ([Win32Cap]::IsWindowVisible($hWnd)) {
        $sb = New-Object System.Text.StringBuilder 256
        [Win32Cap]::GetWindowText($hWnd, $sb, 256) | Out-Null
        $title = $sb.ToString()
        if ($title -eq 'outlookmAnAger') {
            $script:targetHwnd = $hWnd
            return $false
        }
    }
    return $true
}

[Win32Cap]::EnumWindows($callback, [IntPtr]::Zero) | Out-Null

if ($targetHwnd -eq [IntPtr]::Zero) {
    Write-Error "Could not find the outlookmAnAger window. Is the app running?"
    exit 1
}

# NO SetForegroundWindow - we capture in the background!

# Setup
$projRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
$artifactsDir = Join-Path $projRoot 'artifacts'
if (-not (Test-Path $artifactsDir)) { New-Item -ItemType Directory -Path $artifactsDir | Out-Null }

# Find next batch number
$existing = Get-ChildItem -Path $artifactsDir -Filter 'screenshot-*.png' -ErrorAction SilentlyContinue |
    ForEach-Object { if ($_.BaseName -match 'screenshot-(\d+)') { [int]$Matches[1] } } |
    Sort-Object -Descending | Select-Object -First 1
$batchNum = if ($null -eq $existing) { 1 } else { $existing + 1 }

function Capture-Window {
    param([string]$Suffix)

    $bitmap = [Win32Cap]::CaptureWindow($targetHwnd)
    if ($null -eq $bitmap) {
        Write-Host "Failed to capture window for $Suffix"
        return
    }

    $outPath = Join-Path $artifactsDir "screenshot-$batchNum-$Suffix.png"
    $bitmap.Save($outPath, [System.Drawing.Imaging.ImageFormat]::Png)
    $bitmap.Dispose()
    Write-Host "Saved: $outPath"
    $script:batchNum++
}

# Use UI Automation to find tab items (does not require focus)
$rootElement = [System.Windows.Automation.AutomationElement]::FromHandle($targetHwnd)
$tabCondition = New-Object System.Windows.Automation.PropertyCondition(
    [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
    [System.Windows.Automation.ControlType]::TabItem
)
$tabItems = $rootElement.FindAll([System.Windows.Automation.TreeScope]::Descendants, $tabCondition)

$tabNames = @('signatures', 'permissions', 'extras')

for ($i = 0; $i -lt $tabItems.Count -and $i -lt $tabNames.Count; $i++) {
    $tab = $tabItems[$i]
    try {
        $selectPattern = $tab.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern)
        $selectPattern.Select()
    } catch {
        try {
            $invokePattern = $tab.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
            $invokePattern.Invoke()
        } catch {
            Write-Host "Could not switch to tab $i ($($tabNames[$i])): $_"
        }
    }
    Start-Sleep -Milliseconds 500
    Capture-Window -Suffix $tabNames[$i]
}

# Return to first tab
if ($tabItems.Count -gt 0) {
    try {
        $selectPattern = $tabItems[0].GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern)
        $selectPattern.Select()
    } catch {}
}

Write-Host "All tabs captured (no focus stolen)."
