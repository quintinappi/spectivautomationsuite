@echo off
setlocal

echo ============================================
echo  Assembly Cloner Add-In Deployment
echo ============================================
echo.

set "ADDIN_NAME=AssemblyClonerAddIn"
set "TARGET_DIR=C:\ProgramData\Autodesk\Inventor 2026\Addins"
set "SCRIPT_DIR=%~dp0"
set "PROJECT_DIR=%SCRIPT_DIR%..\InventorAddIn\AssemblyClonerAddIn"
set "BUILD_DIR=%PROJECT_DIR%\bin\x64\Release"

echo Script Dir: %SCRIPT_DIR%
echo Project Dir: %PROJECT_DIR%
echo Build Dir: %BUILD_DIR%
echo Target: %TARGET_DIR%
echo.

:: Check if DLL exists
if not exist "%BUILD_DIR%\%ADDIN_NAME%.dll" (
    echo ERROR: DLL not found at %BUILD_DIR%\%ADDIN_NAME%.dll
    echo Please build the project first in Visual Studio.
    echo.
    echo Build steps:
    echo   1. Open AssemblyClonerAddIn.sln in Visual Studio
    echo   2. Set configuration to Release ^| x64
    echo   3. Build ^> Build Solution
    pause
    exit /b 1
)

:: Create target directory if needed
if not exist "%TARGET_DIR%" (
    echo Creating add-in directory...
    mkdir "%TARGET_DIR%"
)

:: Copy files
echo Copying files...
copy /Y "%BUILD_DIR%\%ADDIN_NAME%.dll" "%TARGET_DIR%\"
copy /Y "%PROJECT_DIR%\%ADDIN_NAME%.addin" "%TARGET_DIR%\"

if errorlevel 1 (
    echo.
    echo ERROR: Failed to copy files. Try running as Administrator.
    pause
    exit /b 1
)

echo.
echo ============================================
echo  Deployment Complete!
echo ============================================
echo.
echo Files deployed to:
echo   %TARGET_DIR%
echo.
echo IMPORTANT - Next steps:
echo   1. Close Inventor COMPLETELY (check Task Manager)
echo   2. Start Inventor fresh
echo   3. Go to Tools ^> Add-Ins
echo   4. Look for "Assembly Cloner with iLogic Patcher"
echo   5. Check the "Loaded" checkbox if not already loaded
echo.
echo If still not showing:
echo   - Check Windows Event Viewer for .NET errors
echo   - Verify Inventor 2026 is installed correctly
echo.
pause
