@echo off
echo ========================================
echo   DEPLOYING INVENTOR ADD-IN
echo ========================================
echo.

set "SOURCE_DLL="
set "SOURCE_ADDIN=%~dp0AssemblyClonerAddIn\AssemblyClonerAddIn.addin"
set "TARGET_ALL_USERS=C:\ProgramData\Autodesk\Inventor 2026\Addins"
set "TARGET_PER_USER=%APPDATA%\Autodesk\Inventor Addins"

if exist "%~dp0AssemblyClonerAddIn\bin\x64\Debug\AssemblyClonerAddIn.dll" set "SOURCE_DLL=%~dp0AssemblyClonerAddIn\bin\x64\Debug\AssemblyClonerAddIn.dll"
if "%SOURCE_DLL%"=="" if exist "%~dp0AssemblyClonerAddIn\bin\Debug\AssemblyClonerAddIn.dll" set "SOURCE_DLL=%~dp0AssemblyClonerAddIn\bin\Debug\AssemblyClonerAddIn.dll"
if "%SOURCE_DLL%"=="" if exist "%~dp0AssemblyClonerAddIn\bin\x64\Release\AssemblyClonerAddIn.dll" set "SOURCE_DLL=%~dp0AssemblyClonerAddIn\bin\x64\Release\AssemblyClonerAddIn.dll"
if "%SOURCE_DLL%"=="" if exist "%~dp0AssemblyClonerAddIn\bin\Release\AssemblyClonerAddIn.dll" set "SOURCE_DLL=%~dp0AssemblyClonerAddIn\bin\Release\AssemblyClonerAddIn.dll"

echo Source DLL: %SOURCE_DLL%
echo Source ADDIN: %SOURCE_ADDIN%
echo Target (All-Users): %TARGET_ALL_USERS%
echo Target (Per-User): %TARGET_PER_USER%
echo.
echo DEPLOYMENT LOCATIONS:
echo - All-Users: C:\ProgramData\Autodesk\Inventor 2026\Addins\ (REQUIRES ADMIN)
echo - Per-User:  %%APPDATA%%\Autodesk\Inventor Addins\ (used as fallback and sync)
echo Build output priority: x64\Debug -> Debug -> x64\Release -> Release
echo.

if "%SOURCE_DLL%"=="" (
    echo ERROR: DLL not found!
    echo Checked these outputs:
    echo - %~dp0AssemblyClonerAddIn\bin\x64\Release\AssemblyClonerAddIn.dll
    echo - %~dp0AssemblyClonerAddIn\bin\x64\Debug\AssemblyClonerAddIn.dll
    echo - %~dp0AssemblyClonerAddIn\bin\Release\AssemblyClonerAddIn.dll
    echo - %~dp0AssemblyClonerAddIn\bin\Debug\AssemblyClonerAddIn.dll
    echo.
    echo Make sure to build the project first:
    echo 1. Open AssemblyClonerAddIn.sln in Visual Studio
    echo 2. Build Debug or Release and Any CPU or x64; script picks best output automatically
    echo 3. Build -^> Build Solution
    pause
    exit /b 1
)

if not exist "%SOURCE_ADDIN%" (
    echo ERROR: .addin manifest not found!
    echo Expected at: %SOURCE_ADDIN%
    pause
    exit /b 1
)

if not exist "%TARGET_PER_USER%" (
    mkdir "%TARGET_PER_USER%" > nul 2>&1
)

echo Copying DLL to per-user location...
copy /Y "%SOURCE_DLL%" "%TARGET_PER_USER%\" > nul
if errorlevel 1 (
    echo ERROR: Failed to copy DLL to per-user location.
    pause
    exit /b 1
)

echo Copying .addin manifest to per-user location...
copy /Y "%SOURCE_ADDIN%" "%TARGET_PER_USER%\" > nul
if errorlevel 1 (
    echo ERROR: Failed to copy .addin to per-user location.
    pause
    exit /b 1
)

echo.
echo Syncing all-users location (admin may be required)...
copy /Y "%SOURCE_DLL%" "%TARGET_ALL_USERS%\" > nul
if errorlevel 1 (
    echo WARNING: Could not copy DLL to all-users location; likely needs admin rights.
    echo          Per-user deployment succeeded and should still load in Inventor.
) else (
    echo [OK] DLL copied to all-users location.
)

copy /Y "%SOURCE_ADDIN%" "%TARGET_ALL_USERS%\" > nul
if errorlevel 1 (
    echo WARNING: Could not copy .addin to all-users location; likely needs admin rights.
    echo          Per-user deployment succeeded and should still load in Inventor.
) else (
    echo [OK] .addin copied to all-users location.
)

echo.
echo ========================================
echo   DEPLOYMENT SUCCESSFUL!
echo ========================================
echo.
echo Files copied to per-user location: %TARGET_PER_USER%
echo Attempted sync to all-users location: %TARGET_ALL_USERS%
echo.
echo ADD-IN FEATURES:
echo - Clone Assembly: Clone assemblies with iLogic patching
echo - Scan iLogic: View iLogic rules in current document
echo - Document Info: View iProperties and mass properties
echo - Part Renamer: Rename parts with heritage method
echo - Part Cloner: Clone individual parts
echo - Smart Inspector: Inspect assembly structure and parameters
echo - Beam Generator: Create SANS steel section beams
echo - Update Doc Settings: Fix BOM decimal precision
echo - Place Parts in IDW: NEW! Scan for PL/S355JR parts and place in IDW
echo.
echo NEXT STEPS:
echo 1. CLOSE Inventor completely (check Task Manager!)
echo 2. START Inventor
echo 3. Go to Tools -^> Add-Ins
echo 4. Look for "Assembly Cloner with iLogic Patcher"
echo 5. Check "Loaded" if not already checked
echo.
echo If still not showing, check:
echo - Windows Event Viewer for .NET errors
echo - Inventor version is 2026
echo - .NET Framework 4.8 is installed
echo.
echo LOG FILES:
echo Log files are saved to: %%USERPROFILE%%\Documents\InventorAutomationSuite\Logs\
echo.
pause
