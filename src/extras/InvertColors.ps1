# Invert Screen Colors
# Toggles Windows High Contrast mode, making the screen look "wrong".
# Run again to toggle it back.
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

[StructLayout(LayoutKind.Sequential)]
public struct HIGHCONTRAST {
    public int cbSize;
    public int dwFlags;
    public string lpszDefaultScheme;
}

public class AccessibilityHelper {
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool SystemParametersInfo(int uAction, int uParam, ref HIGHCONTRAST lpvParam, int fuWinIni);
}
"@

$hc         = New-Object HIGHCONTRAST
$hc.cbSize  = [System.Runtime.InteropServices.Marshal]::SizeOf($hc)

# Read current state
[AccessibilityHelper]::SystemParametersInfo(66, 0, [ref]$hc, 0) | Out-Null   # SPI_GETHIGHCONTRAST

$HCF_HIGHCONTRASTON = 0x1
if ($hc.dwFlags -band $HCF_HIGHCONTRASTON) {
    $hc.dwFlags = $hc.dwFlags -band (-bnot $HCF_HIGHCONTRASTON)   # turn off
} else {
    $hc.dwFlags = $hc.dwFlags -bor $HCF_HIGHCONTRASTON            # turn on
}

[AccessibilityHelper]::SystemParametersInfo(67, 0, [ref]$hc, 3) | Out-Null   # SPI_SETHIGHCONTRAST
