@echo off
REM ========================================================================
REM Spectiv Inventor Suite - Deployment Script
REM ========================================================================
echo ========================================
echo Spectiv Inventor Suite - Installer
echo ========================================
echo.

REM Get paths
set "BUILD_DLL=C:\Users\Quintin\source\repos\SpectivInventorSuite\SpectivInventorSuite\bin\Debug\SpectivInventorSuite.dll"
set "ADDIN_SOURCE=C:\Users\Quintin\Documents\Spectiv\3. Working\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\Migration to Add-In\SpectivInventorSuite.addin"
set "PLUGIN_DIR=%APPDATA%\Autodesk\ApplicationPlugins\SpectivInventorSuite"

REM Check if build exists
if not exist "%BUILD_DLL%" (
    echo ERROR: DLL not found at:
    echo   %BUILD_DLL%
    echo.
    echo Please build the project first in Visual Studio.
    echo.
    pause
    exit /b 1
)

REM Create plugin directory
echo Creating plugin directory: %PLUGIN_DIR%
if not exist "%PLUGIN_DIR%" (
    mkdir "%PLUGIN_DIR%"
)

REM Copy DLL
echo.
echo Copying DLL...
copy /Y "%BUILD_DLL%" "%PLUGIN_DIR%\"

if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Failed to copy DLL
    pause
    exit /b 1
)

REM Copy .addin file
echo.
echo Copying .addin manifest...
copy /Y "%ADDIN_SOURCE%" "%PLUGIN_DIR%\"

if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Failed to copy .addin file
    pause
    exit /b 1
)

REM Verify installation
echo.
echo ========================================
echo Installation Complete!
echo ========================================
echo.
echo Installed to: %PLUGIN_DIR%
echo.
echo Files installed:
dir /B "%PLUGIN_DIR%"
echo.
echo ========================================
echo Next Steps:
echo ========================================
echo 1. Close Inventor if open
echo 2. Restart Inventor
echo 3. Open any assembly file (.iam)
echo 4. Click "Assembly Cloner" button in Assembly tab
echo ========================================
echo.

REM Open plugin folder
explorer "%PLUGIN_DIR%"

echo Installer completed. Press any key to exit...
pause >nul
