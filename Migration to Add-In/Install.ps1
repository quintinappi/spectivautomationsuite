# Spectiv Inventor Suite - Installation Script
# Run: Right-click → Run with PowerShell

Write-Host "========================================" -ForegroundColor Cyan
Write-Host " Spectiv Inventor Suite - Installer" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Define paths
$buildDll = "C:\Users\Quintin\source\repos\SpectivInventorSuite\SpectivInventorSuite\bin\Debug\SpectivInventorSuite.dll"
$addinSource = "C:\Users\Quintin\Documents\Spectiv\3. Working\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\Migration to Add-In\SpectivInventorSuite.addin"
$pluginDir = "$env:APPDATA\Autodesk\ApplicationPlugins\SpectivInventorSuite"

# Check if build exists
Write-Host "Checking build output..." -ForegroundColor Yellow
if (-not (Test-Path $buildDll)) {
    Write-Host "ERROR: DLL not found!" -ForegroundColor Red
    Write-Host "Expected: $buildDll" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please build the project first:" -ForegroundColor Yellow
    Write-Host "1. Open SpectivInventorSuite.sln in Visual Studio" -ForegroundColor White
    Write-Host "2. Build → Build Solution" -ForegroundColor White
    Write-Host "3. Run this installer again" -ForegroundColor White
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "Build found: OK" -ForegroundColor Green
Write-Host ""

# Create plugin directory
Write-Host "Creating plugin directory..." -ForegroundColor Yellow
if (-not (Test-Path $pluginDir)) {
    New-Item -ItemType Directory -Path $pluginDir -Force | Out-Null
    Write-Host "Created: $pluginDir" -ForegroundColor Green
} else {
    Write-Host "Exists: $pluginDir" -ForegroundColor Gray
}

# Copy DLL
Write-Host ""
Write-Host "Copying DLL..." -ForegroundColor Yellow
try {
    Copy-Item -Path $buildDll -Destination $pluginDir -Force
    Write-Host "Copied: SpectivInventorSuite.dll" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Failed to copy DLL" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Copy .addin file
Write-Host ""
Write-Host "Copying .addin manifest..." -ForegroundColor Yellow
if (Test-Path $addinSource) {
    try {
        Copy-Item -Path $addinSource -Destination $pluginDir -Force
        Write-Host "Copied: SpectivInventorSuite.addin" -ForegroundColor Green
    } catch {
        Write-Host "ERROR: Failed to copy .addin file" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
} else {
    Write-Host "ERROR: .addin source not found!" -ForegroundColor Red
    Write-Host "Expected: $addinSource" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Verify installation
Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host " Installation Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Installed to: $pluginDir" -ForegroundColor Cyan
Write-Host ""
Write-Host "Files in plugin folder:" -ForegroundColor Yellow
Get-ChildItem -Path $pluginDir | ForEach-Object {
    Write-Host "  - $($_.Name)" -ForegroundColor White
}
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " Next Steps:" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "1. Close Inventor if open" -ForegroundColor White
Write-Host "2. Restart Inventor" -ForegroundColor White
Write-Host "3. Open any assembly file (.iam)" -ForegroundColor White
Write-Host "4. Look for 'Assembly Cloner' button in Assembly tab" -ForegroundColor White
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Open plugin folder
Start-Process explorer.exe -ArgumentList $pluginDir

Write-Host "Plugin folder opened." -ForegroundColor Gray
Write-Host ""
Read-Host "Press Enter to exit"
