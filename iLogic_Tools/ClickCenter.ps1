# ClickCenter.ps1 - Click at center of Inventor window
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class MouseOps {
    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int X, int Y);
    
    [DllImport("user32.dll")]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, int dwExtraInfo);
    
    [DllImport("user32.dll")]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
    
    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left, Top, Right, Bottom;
    }
    
    public const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
    public const uint MOUSEEVENTF_LEFTUP = 0x0004;
    
    public static void Click(int x, int y) {
        SetCursorPos(x, y);
        System.Threading.Thread.Sleep(50);
        mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
        System.Threading.Thread.Sleep(50);
        mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
    }
}
"@

# Find Inventor window
$hwnd = [MouseOps]::FindWindow($null, "Autodesk Inventor Professional 2026")
if ($hwnd -eq [IntPtr]::Zero) {
    # Try partial match
    $proc = Get-Process | Where-Object { $_.MainWindowTitle -like "*Inventor*" } | Select-Object -First 1
    if ($proc) {
        $hwnd = $proc.MainWindowHandle
    }
}

if ($hwnd -ne [IntPtr]::Zero) {
    Write-Host "Found Inventor window: $hwnd"
    
    # Bring to foreground
    [MouseOps]::SetForegroundWindow($hwnd)
    Start-Sleep -Milliseconds 200
    
    # Get window rect
    $rect = New-Object MouseOps+RECT
    [MouseOps]::GetWindowRect($hwnd, [ref]$rect)
    
    Write-Host "Window rect: Left=$($rect.Left), Top=$($rect.Top), Right=$($rect.Right), Bottom=$($rect.Bottom)"
    
    # Calculate center of window (offset slightly to the right where the 3D view is)
    $centerX = [int](($rect.Left + $rect.Right) / 2) + 100
    $centerY = [int](($rect.Top + $rect.Bottom) / 2)
    
    Write-Host "Clicking at: $centerX, $centerY"
    
    # Click
    [MouseOps]::Click($centerX, $centerY)
    
    Write-Host "Click sent!"
} else {
    Write-Host "Could not find Inventor window"
}
