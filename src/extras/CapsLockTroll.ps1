# Caps Lock Troll
# Toggles Caps Lock every few seconds — makes the victim wonder why their typing keeps shifting.
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class CapsLock {
    [DllImport("user32.dll")]
    public static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);
}
"@

while ($true) {
    Start-Sleep -Seconds 8
    [CapsLock]::keybd_event(0x14, 0, 1, 0)   # Toggle Caps Lock
}
