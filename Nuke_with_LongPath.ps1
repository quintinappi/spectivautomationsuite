# NUKE USING LONG PATH PREFIX (Bypasses 260 char limit)

$ErrorActionPreference = "SilentlyContinue"

# Convert to absolute path with \\?\ prefix
$path = "c:\Users\Quintin\Documents\Spectiv\3. Working\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\Backup Working"
$longPath = "\\?\$path"

Write-Host "=== LONG PATH NUKE ===" -ForegroundColor Cyan
Write-Host "Target: $longPath" -ForegroundColor Yellow
Write-Host ""

# Use .NET Framework for long path support
Write-Host "[1/2] Using .NET DirectoryInfo..." -ForegroundColor Cyan
try {
    $dir = [System.IO.DirectoryInfo]::new($longPath)
    $dir.Delete($true)
    Write-Host "Success!" -ForegroundColor Green

    if (-not (Test-Path $path)) {
        Write-Host ""
        Write-Host "=== FOLDER DELETED ===" -ForegroundColor Green
        exit 0
    }
} catch {
    Write-Host "Failed: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "[2/2] Using kernel32.dll RmDir..." -ForegroundColor Cyan
try {
    Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    public class Win32 {
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        public static extern bool RemoveDirectory(string lpPathName);
    }
"@

    $result = [Win32]::RemoveDirectory($longPath)
    if ($result) {
        Write-Host "Success!" -ForegroundColor Green
        Write-Host ""
        Write-Host "=== FOLDER DELETED ===" -ForegroundColor Green
    } else {
        Write-Host "Failed with Win32 error" -ForegroundColor Red
        Write-Host ""
        Write-Host "=== MANUAL INTERVENTION NEEDED ===" -ForegroundColor Red
        Write-Host ""
        Write-Host "Download this tool:" -ForegroundColor Yellow
        Write-Host "https://github.com/staticdev/vlongpath/releases" -ForegroundColor White
        Write-Host "Or reboot into Safe Mode and try again" -ForegroundColor White
    }
} catch {
    Write-Host "Failed: $($_.Exception.Message)" -ForegroundColor Red
}
