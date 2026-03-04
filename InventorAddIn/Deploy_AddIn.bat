@echo off
REM Deploy AssemblyClonerAddIn to Inventor Add-In Folder
REM Run this AFTER building the project in Release x64 mode

set DEPLOY_FAILED=0

echo ========================================
echo   Inventor Add-In Deployment
echo ========================================
echo.

REM Check if built DLL exists
if not exist "AssemblyClonerAddIn\bin\x64\Release\AssemblyClonerAddIn.dll" (
    echo ERROR: DLL not found!
    echo Please build the project in Release x64 mode first.
    echo.
    pause
    exit /b 1
)

REM Active Inventor 2026 machine-wide add-in path (priority)
set INVENTOR_ADDIN_PATH_PRIMARY=%ProgramData%\Autodesk\Inventor 2026\Addins

REM User add-in path (secondary)
set INVENTOR_ADDIN_PATH_SECONDARY=%APPDATA%\Autodesk\Inventor 2026\Addins

REM Create primary directory if it doesn't exist
if not exist "%INVENTOR_ADDIN_PATH_PRIMARY%" (
    echo Creating add-in directory: %INVENTOR_ADDIN_PATH_PRIMARY%
    mkdir "%INVENTOR_ADDIN_PATH_PRIMARY%"
)

REM Create secondary directory if it doesn't exist
if not exist "%INVENTOR_ADDIN_PATH_SECONDARY%" (
    echo Creating add-in directory: %INVENTOR_ADDIN_PATH_SECONDARY%
    mkdir "%INVENTOR_ADDIN_PATH_SECONDARY%"
)

echo Copying files to primary path: %INVENTOR_ADDIN_PATH_PRIMARY%
echo.

REM Copy DLL + manifest to primary path
echo Copying AssemblyClonerAddIn.dll to primary...
copy /Y "AssemblyClonerAddIn\bin\x64\Release\AssemblyClonerAddIn.dll" "%INVENTOR_ADDIN_PATH_PRIMARY%\"
if errorlevel 1 set DEPLOY_FAILED=1
echo Copying AssemblyClonerAddIn.addin to primary...
copy /Y "AssemblyClonerAddIn\AssemblyClonerAddIn.addin" "%INVENTOR_ADDIN_PATH_PRIMARY%\"
if errorlevel 1 set DEPLOY_FAILED=1

echo.
echo Copying files to secondary path: %INVENTOR_ADDIN_PATH_SECONDARY%
echo.

REM Copy DLL + manifest to secondary path
echo Copying AssemblyClonerAddIn.dll to secondary...
copy /Y "AssemblyClonerAddIn\bin\x64\Release\AssemblyClonerAddIn.dll" "%INVENTOR_ADDIN_PATH_SECONDARY%\"
if errorlevel 1 set DEPLOY_FAILED=1
echo Copying AssemblyClonerAddIn.addin to secondary...
copy /Y "AssemblyClonerAddIn\AssemblyClonerAddIn.addin" "%INVENTOR_ADDIN_PATH_SECONDARY%\"
if errorlevel 1 set DEPLOY_FAILED=1

echo.
echo ========================================
echo   Deployment Complete!
echo ========================================
echo.
echo NEXT STEPS:
echo 1. Close Inventor if it's running
echo 2. Start Inventor
echo 3. The add-in will load automatically
echo 4. It will monitor plate parts and fix decimals on save
echo.
echo Primary Add-In Location: %INVENTOR_ADDIN_PATH_PRIMARY%
echo Secondary Add-In Location: %INVENTOR_ADDIN_PATH_SECONDARY%
echo.

if "%DEPLOY_FAILED%"=="1" (
    echo WARNING: One or more copy operations failed.
    echo If ProgramData copy failed with Access denied, run this script as Administrator.
    exit /b 1
)

pause
