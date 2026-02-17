# ═══ Theme Definitions ═══════════════════════════════════════════════════════
# To add a theme: add an entry here and a matching ListBoxItem in MainWindow.xaml.
$script:Themes = [ordered]@{
    'Dark' = @{
        BgBrush            = '#1E1E2E'
        SurfaceBrush       = '#2A2A3C'
        SurfaceHoverBrush  = '#33334A'
        BorderBrush        = '#3E3E55'
        AccentBrush        = '#7C6FE0'
        AccentHoverBrush   = '#9589E8'
        AccentSubtleBrush  = '#2E2B4A'
        TextPrimaryBrush   = '#EAEAEF'
        TextSecondaryBrush = '#9898A8'
        TextOnAccentBrush  = '#FFFFFF'
    }
    'Light' = @{
        BgBrush            = '#F0F6FF'
        SurfaceBrush       = '#FFFFFF'
        SurfaceHoverBrush  = '#E6F0FF'
        BorderBrush        = '#C8D8EE'
        AccentBrush        = '#6DA8E0'
        AccentHoverBrush   = '#4E90CC'
        AccentSubtleBrush  = '#E0EEFF'
        TextPrimaryBrush   = '#1A2333'
        TextSecondaryBrush = '#4A6080'
        TextOnAccentBrush  = '#FFFFFF'
    }
    'Green' = @{
        BgBrush            = '#1A2420'
        SurfaceBrush       = '#24332D'
        SurfaceHoverBrush  = '#2E3F37'
        BorderBrush        = '#3D5548'
        AccentBrush        = '#4E8A6E'
        AccentHoverBrush   = '#62A584'
        AccentSubtleBrush  = '#233D30'
        TextPrimaryBrush   = '#E2EAE6'
        TextSecondaryBrush = '#8FA99A'
        TextOnAccentBrush  = '#FFFFFF'
    }
    'Maroon' = @{
        BgBrush            = '#241A1E'
        SurfaceBrush       = '#332428'
        SurfaceHoverBrush  = '#3F2E33'
        BorderBrush        = '#553D44'
        AccentBrush        = '#8A4E5E'
        AccentHoverBrush   = '#A56275'
        AccentSubtleBrush  = '#3D2330'
        TextPrimaryBrush   = '#EAE2E5'
        TextSecondaryBrush = '#A98F96'
        TextOnAccentBrush  = '#FFFFFF'
    }
    'Retro' = @{
        BgBrush            = '#08060F'
        SurfaceBrush       = '#110D1F'
        SurfaceHoverBrush  = '#1C1630'
        BorderBrush        = '#00F0FF'
        AccentBrush        = '#FF2D78'
        AccentHoverBrush   = '#FF6BB0'
        AccentSubtleBrush  = '#1A0830'
        TextPrimaryBrush   = '#00F0FF'
        TextSecondaryBrush = '#B060FF'
        TextOnAccentBrush  = '#FFFFFF'
    }
}

function Apply-Theme {
    param($Window, [string]$ThemeName)
    if (-not $script:Themes.Contains($ThemeName)) { return }
    $colors = $script:Themes[$ThemeName]
    foreach ($key in $colors.Keys) {
        $color = [System.Windows.Media.ColorConverter]::ConvertFromString($colors[$key])
        $brush = $Window.Resources[$key]
        if ($null -ne $brush -and $brush -is [System.Windows.Media.SolidColorBrush] -and -not $brush.IsFrozen) {
            $brush.Color = [System.Windows.Media.Color] $color
        } else {
            $nb = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color] $color)
            $Window.Resources[$key] = [System.Windows.Media.SolidColorBrush] $nb
        }
    }
    # When a sig with a missing file is selected, re-render the error page with new colors.
    if ($null -ne $script:previewBrowserRef -and $null -ne $script:currentSig) {
        $htmlPath = Get-SignatureHtmlPath -Name $script:currentSig
        if (-not (Test-Path $htmlPath)) {
            $bgHex = $script:Themes[$ThemeName].BgBrush
            $fgHex = $script:Themes[$ThemeName].TextSecondaryBrush
            $script:previewBrowserRef.NavigateToString(
                "<!DOCTYPE html><html style=`"background:$bgHex`"><body style=`"background:$bgHex;color:$fgHex;font-family:Segoe UI;padding:16px`">No HTML file found.</body></html>")
        }
    }
}

# ═══ Retro Hue Animation (Retro/unlocked mode) ══════════════════════════════
$script:retroTimer = $null

# Neon palette: hue stops cycling pink→magenta→violet→blue→cyan and back
# Each entry: [AccentBrush, AccentHoverBrush, BorderBrush, TextPrimaryBrush, TextSecondaryBrush]
$script:retroStops = @(
    @('#FF2D78','#FF6BB0','#00F0FF','#00F0FF','#B060FF'),
    @('#FF60A0','#FFB0D0','#40DFFF','#40DFFF','#C080FF'),
    @('#CC00FF','#DD60FF','#00DFFF','#E0B0FF','#9040FF'),
    @('#8000FF','#AA40FF','#60A0FF','#C0A0FF','#4060FF'),
    @('#00BFFF','#60D8FF','#8000FF','#80E0FF','#6040FF'),
    @('#00F0FF','#40FFFF','#FF2D78','#00FFEE','#B060FF'),
    @('#40DFFF','#80EFFF','#FF60A0','#A0FFEE','#FF80C0'),
    @('#00F0FF','#40DFFF','#CC00FF','#80E0FF','#FF60A0')
)

function Invoke-RetroLerpHex {
    param([string]$hexA, [string]$hexB, [double]$t)
    $cA = [System.Windows.Media.ColorConverter]::ConvertFromString($hexA)
    $cB = [System.Windows.Media.ColorConverter]::ConvertFromString($hexB)
    $r  = [byte]([Math]::Round($cA.R * (1 - $t) + $cB.R * $t))
    $g  = [byte]([Math]::Round($cA.G * (1 - $t) + $cB.G * $t))
    $b  = [byte]([Math]::Round($cA.B * (1 - $t) + $cB.B * $t))
    return [System.Windows.Media.Color]::FromRgb($r, $g, $b)
}

function Invoke-RetroTick {
    try {
        $stops = $script:retroStops
        $n     = $stops.Count
        $phase = $script:retroPhase

        $script:retroPhase = ($phase + 0.004) % $n

        $iA  = [int][Math]::Floor($phase) % $n
        $iB  = ($iA + 1) % $n
        $t   = $phase - [Math]::Floor($phase)
        $rowA = $stops[$iA]
        $rowB = $stops[$iB]

        $keys = @('AccentBrush','AccentHoverBrush','BorderBrush','TextPrimaryBrush','TextSecondaryBrush')
        $win  = $script:retroWinRef
        for ($k = 0; $k -lt $keys.Count; $k++) {
            $col   = Invoke-RetroLerpHex -hexA $rowA[$k] -hexB $rowB[$k] -t $t
            $brush = $win.Resources[$keys[$k]]
            if ($null -ne $brush -and $brush -is [System.Windows.Media.SolidColorBrush] -and -not $brush.IsFrozen) {
                $brush.Color = $col
            } else {
                $win.Resources[$keys[$k]] = [System.Windows.Media.SolidColorBrush]::new($col)
            }
        }
    } catch {}
}

function Start-RetroHueTimer {
    param($Window)
    if ($null -ne $script:retroTimer -and $script:retroTimer.IsEnabled) { return }

    $script:retroPhase  = 0.0
    $script:retroWinRef = $Window

    $script:retroTimer          = New-Object System.Windows.Threading.DispatcherTimer
    $script:retroTimer.Interval = [TimeSpan]::FromMilliseconds(33)
    $script:retroTimer.Add_Tick({ Invoke-RetroTick })
    $script:retroTimer.Start()
}

function Stop-RetroHueTimer {
    if ($null -ne $script:retroTimer) {
        $script:retroTimer.Stop()
        $script:retroTimer = $null
    }
}

# ═══ Signature Colour Map ═══════════════════════════════════════════════════
function Build-SigColourMap {
    $palette = @('#D96C6C','#4FA89B','#C49A2A','#7B68C8','#4A8CC1','#6B9E3F','#B85C8A')
    $hashToColour = @{}
    $nameToColour = @{}
    $idx = 0
    foreach ($name in (Get-Signatures)) {
        $path = Get-SignatureHtmlPath -Name $name
        if (Test-Path $path) {
            $hash = (Get-FileHash $path -Algorithm MD5).Hash
            if (-not $hashToColour.ContainsKey($hash)) {
                $hashToColour[$hash] = $palette[$idx % $palette.Count]
                $idx++
            }
            $nameToColour[$name] = $hashToColour[$hash]
        }
    }
    return $nameToColour
}
