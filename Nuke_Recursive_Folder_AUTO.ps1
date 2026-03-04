# AUTOMATIC RECURSIVE FOLDER NUKE (No prompts)
# Use this to delete folders with infinite nesting

$ErrorActionPreference = "SilentlyContinue"

Write-Host "=== AUTO NUKE: Recursive Folder ===" -ForegroundColor Cyan
Write-Host ""

$targetFolder = "c:\Users\Quintin\Documents\Spectiv\3. Working\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\Archives\Backups\2026-01-14_PRE_CLEANUP"

Write-Host "Target: $targetFolder" -ForegroundColor Yellow

if (-not (Test-Path $targetFolder)) {
    Write-Host "Folder not found!" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "[1/3] Robocopy mirror (empty folder trick)..." -ForegroundColor Cyan
$emptyFolder = "C:\temp_empty_folder_" + (Get-Random -Maximum 9999)
New-Item -Path $emptyFolder -ItemType Directory -Force | Out-Null
& robocopy $emptyFolder $targetFolder /MIR /R:1 /W:1 /NFL /NDL /NJH /NJS | Out-Null
Remove-Item $emptyFolder -Force -Recurse -ErrorAction SilentlyContinue
Write-Host "Done." -ForegroundColor Green

Write-Host "[2/3] PowerShell Force-Remove..." -ForegroundColor Cyan
Remove-Item $targetFolder -Force -Recurse -ErrorAction SilentlyContinue
if (-not (Test-Path $targetFolder)) {
    Write-Host "Success!" -ForegroundColor Green
    Write-Host ""
    Write-Host "=== FOLDER DELETED ===" -ForegroundColor Green
    exit 0
}

Write-Host "Failed." -ForegroundColor Red

Write-Host "[3/3] CMD rmdir..." -ForegroundColor Cyan
cmd.exe /c "rmdir /s /q `"$targetFolder`"" 2>$null
Start-Sleep -Seconds 2

if (-not (Test-Path $targetFolder)) {
    Write-Host "Success!" -ForegroundColor Green
    Write-Host ""
    Write-Host "=== FOLDER DELETED ===" -ForegroundColor Green
} else {
    Write-Host "Failed." -ForegroundColor Red
    Write-Host ""
    Write-Host "=== FOLDER STILL EXISTS ===" -ForegroundColor Red
    Write-Host ""
    Write-Host "Manual options:" -ForegroundColor Yellow
    Write-Host "1. Reboot and try this script again" -ForegroundColor White
    Write-Host "2. Use LongPathTool: https://pathdistcheck.sourceforge.net/" -ForegroundColor White
    Write-Host "3. Use Linux Live USB (mounts NTFS read-write)" -ForegroundColor White
    Write-Host "4. Use 'Unlocker' or 'FileAssassin' tools" -ForegroundColor White
}

exit 0
