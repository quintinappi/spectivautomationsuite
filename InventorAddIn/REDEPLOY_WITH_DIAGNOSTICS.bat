@echo off
echo ========================================
echo REDEPLOY ADD-IN WITH DIAGNOSTIC MESSAGES
echo ========================================
echo.
echo This will copy the updated add-in DLL with diagnostic messages.
echo.
echo CRITICAL: Make sure Inventor is COMPLETELY CLOSED first!
echo Check Task Manager - no Inventor.exe should be running.
echo.
pause

set "SOURCE_DLL=%~dp0AssemblyClonerAddIn\bin\x64\Release\AssemblyClonerAddIn.dll"
set "TARGET_DLL=C:\ProgramData\Autodesk\Inventor 2026\Addins\AssemblyClonerAddIn.dll"

echo.
echo Checking source DLL...
if not exist "%SOURCE_DLL%" (
    echo ERROR: Source DLL not found: %SOURCE_DLL%
    pause
    exit /b 1
)

echo Source DLL found: %SOURCE_DLL%
echo.
echo Copying to: %TARGET_DLL%
copy /Y "%SOURCE_DLL%" "%TARGET_DLL%"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: Copy failed! Make sure:
    echo 1. Inventor is completely closed
    echo 2. This script is run as Administrator
    pause
    exit /b 1
)

echo.
echo ========================================
echo SUCCESS! Add-in deployed with diagnostics
echo ========================================
echo.
echo The add-in now provides detailed feedback:
echo   - SUCCESS: Settings updated and BOM refreshed
echo   - NOT_PART: Document is not a .ipt file
echo   - NOT_SHEET_METAL: Part is not sheet metal
echo   - NOT_PLATE: Part doesn't match plate criteria
echo.
echo Next steps:
echo 1. Start Inventor
echo 2. Open a plate part (.ipt)
echo 3. Click "Update Doc Settings" button
echo 4. Read the diagnostic message carefully
echo.
pause
