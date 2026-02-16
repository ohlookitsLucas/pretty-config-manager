<#
  capture-window.ps1
  Takes a screenshot of one (or all) tabs in the outlookmAnAger window.
  Uses PrintWindow API - does NOT steal focus or bring window to front.
  Output: artifacts\screenshot-N-tabname.png (numbered sequentially)

  Usage:
    .\capture-window.ps1              # capture all three tabs
    .\capture-window.ps1 signatures   # capture only the Signatures tab
    .\capture-window.ps1 permissions  # capture only the Permissions tab
    .\capture-window.ps1 extras       # capture only the Extras tab
#>
param(
    [string]$Tab = ''   # optional: 'signatures', 'permissions', 'extras'
)

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

$allTabNames = @('signatures', 'permissions', 'extras')

# Decide which indices to capture
$normalised = $Tab.Trim().ToLower()
if ($normalised -eq '') {
    $indicesToCapture = 0..($allTabNames.Count - 1)
} else {
    $idx = $allTabNames.IndexOf($normalised)
    if ($idx -lt 0) {
        Write-Error "Unknown tab '$Tab'. Valid values: $($allTabNames -join ', ')"
        exit 1
    }
    $indicesToCapture = @($idx)
}

foreach ($i in $indicesToCapture) {
    $tabItem = $tabItems[$i]
    try {
        $selectPattern = $tabItem.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern)
        $selectPattern.Select()
    } catch {
        try {
            $invokePattern = $tabItem.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
            $invokePattern.Invoke()
        } catch {
            Write-Host "Could not switch to tab $i ($($allTabNames[$i])): $_"
        }
    }
    Start-Sleep -Milliseconds 500
    Capture-Window -Suffix $allTabNames[$i]
}

# Return to first tab only when capturing all tabs
if ($normalised -eq '' -and $tabItems.Count -gt 0) {
    try {
        $selectPattern = $tabItems[0].GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern)
        $selectPattern.Select()
    } catch {}
}

$captured = if ($normalised -eq '') { 'All tabs' } else { "Tab '$normalised'" }
Write-Host "$captured captured (no focus stolen)."
