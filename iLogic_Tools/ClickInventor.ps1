# ClickInventor.ps1 - Click on Inventor window properly
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class MouseHelper {
    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int X, int Y);
    
    [DllImport("user32.dll")]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, int dwExtraInfo);
    
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
    
    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    
    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
    
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left, Top, Right, Bottom;
    }
    
    public const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
    public const uint MOUSEEVENTF_LEFTUP = 0x0004;
    public const int SW_RESTORE = 9;
    
    public static void Click(int x, int y) {
        SetCursorPos(x, y);
        System.Threading.Thread.Sleep(100);
        mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
        System.Threading.Thread.Sleep(50);
        mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
    }
}
"@

# Find Inventor process
$inventorProc = Get-Process | Where-Object { $_.ProcessName -like "*Inventor*" -and $_.MainWindowHandle -ne 0 } | Select-Object -First 1

if ($inventorProc) {
    $hwnd = $inventorProc.MainWindowHandle
    Write-Host "Found Inventor: $($inventorProc.MainWindowTitle)"
    Write-Host "Handle: $hwnd"
    
    # Restore window if minimized
    [MouseHelper]::ShowWindow($hwnd, [MouseHelper]::SW_RESTORE)
    Start-Sleep -Milliseconds 300
    
    # Bring to foreground
    [MouseHelper]::SetForegroundWindow($hwnd)
    Start-Sleep -Milliseconds 300
    
    # Get window rect after restore
    $rect = New-Object MouseHelper+RECT
    [MouseHelper]::GetWindowRect($hwnd, [ref]$rect)
    
    Write-Host "Window: L=$($rect.Left), T=$($rect.Top), R=$($rect.Right), B=$($rect.Bottom)"
    
    # Check if valid
    if ($rect.Left -lt -1000) {
        Write-Host "Window appears minimized/hidden, using screen center..."
        # Use middle of primary screen
        $centerX = 960
        $centerY = 540
    } else {
        # Calculate center, offset toward the 3D view area (right side)
        $width = $rect.Right - $rect.Left
        $height = $rect.Bottom - $rect.Top
        $centerX = $rect.Left + [int]($width * 0.6)  # 60% from left
        $centerY = $rect.Top + [int]($height * 0.5)  # 50% from top
    }
    
    Write-Host "Clicking at: $centerX, $centerY"
    [MouseHelper]::Click($centerX, $centerY)
    Write-Host "Done!"
} else {
    Write-Host "Inventor not found!"
}
