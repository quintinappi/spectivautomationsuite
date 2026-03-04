# NUKE RECURSIVE FOLDER SCRIPT
# Use this to delete folders with infinite nesting that Windows refuses to delete

Write-Host "=== RECURSIVE FOLDER NUKE ===" -ForegroundColor Cyan
Write-Host ""

$targetFolder = "c:\Users\Quintin\Documents\Spectiv\3. Working\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\Archives\Backups\2026-01-14_PRE_CLEANUP"

Write-Host "Target: " -NoNewline
Write-Host $targetFolder -ForegroundColor Yellow

if (-not (Test-Path $targetFolder)) {
    Write-Host "Folder not found!" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "WARNING: This will permanently delete this folder!" -ForegroundColor Red
Write-Host "Press Ctrl+C to cancel, or any key to continue..." -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

Write-Host ""
Write-Host "METHOD 1: Robocopy Mirror (Nuke contents first)" -ForegroundColor Cyan
$emptyFolder = "C:\temp_empty_folder_" + (Get-Random -Maximum 9999)
New-Item -Path $emptyFolder -ItemType Directory -Force | Out-Null

& robocopy $emptyFolder $targetFolder /MIR /R:1 /W:1 | Out-Null
Remove-Item $emptyFolder -Force -Recurse -ErrorAction SilentlyContinue

Write-Host "Robocopy mirror completed." -ForegroundColor Green

Write-Host ""
Write-Host "METHOD 2: PowerShell Remove-Item (Force recursive delete)" -ForegroundColor Cyan
try {
    Remove-Item $targetFolder -Force -Recurse -ErrorAction Stop
    Write-Host "PowerShell delete successful!" -ForegroundColor Green
} catch {
    Write-Host "PowerShell failed: " -NoNewline
    Write-Host $_.Exception.Message -ForegroundColor Red
}

Write-Host ""
Write-Host "METHOD 3: CMD rmdir (Force delete)" -ForegroundColor Cyan
if (Test-Path $targetFolder) {
    cmd.exe /c "rmdir /s /q `"$targetFolder`"" 2>$null
    if (-not (Test-Path $targetFolder)) {
        Write-Host "CMD rmdir successful!" -ForegroundColor Green
    } else {
        Write-Host "CMD rmdir failed." -ForegroundColor Red
    }
}

Write-Host ""
if (Test-Path $targetFolder) {
    Write-Host "STATUS: Folder STILL EXISTS - Manual intervention needed" -ForegroundColor Red
    Write-Host ""
    Write-Host "Try these manual steps:" -ForegroundColor Yellow
    Write-Host "1. Reboot computer and try again" -ForegroundColor White
    Write-Host "2. Use Unlocker tool: https://www.iobit.com/en/iobit-unlocker.php" -ForegroundColor White
    Write-Host "3. Boot into Safe Mode and delete" -ForegroundColor White
    Write-Host "4. Use Linux Live USB to delete (mounts NTFS read-write)" -ForegroundColor White
} else {
    Write-Host "STATUS: Folder successfully deleted!" -ForegroundColor Green
}

Write-Host ""
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
