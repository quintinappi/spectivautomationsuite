@echo off
echo ========================================
echo   INVENTOR ADD-IN DIAGNOSTIC
echo ========================================
echo.

echo [1] Checking deployment status...
echo.
if exist "C:\ProgramData\Autodesk\Inventor 2026\Addins\AssemblyClonerAddIn.dll" (
    echo [OK] DLL is deployed: C:\ProgramData\Autodesk\Inventor 2026\Addins\AssemblyClonerAddIn.dll
    dir "C:\ProgramData\Autodesk\Inventor 2026\Addins\AssemblyClonerAddIn.dll" | findstr /C:"AssemblyClonerAddIn.dll"
) else (
    echo [FAIL] DLL NOT FOUND in deployment folder!
)

if exist "C:\ProgramData\Autodesk\Inventor 2026\Addins\AssemblyClonerAddIn.addin" (
    echo [OK] Manifest is deployed: C:\ProgramData\Autodesk\Inventor 2026\Addins\AssemblyClonerAddIn.addin
) else (
    echo [FAIL] .addin manifest NOT FOUND!
)
echo.

echo [2] Checking .NET Framework...
echo.
reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" /v Release > nul 2>&1
if %errorlevel% equ 0 (
    echo [OK] .NET Framework 4.x is installed
    for /f "tokens=3" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" /v Release ^| findstr Release') do (
        if %%a geq 528040 (
            echo [OK] .NET Framework 4.8 or higher detected
        ) else (
            echo [WARNING] .NET Framework version may be too old
        )
    )
) else (
    echo [FAIL] .NET Framework 4.x NOT FOUND!
)
echo.

echo [3] Checking Inventor installation...
echo.
if exist "C:\Program Files\Autodesk\Inventor 2026" (
    echo [OK] Inventor 2026 folder found
) else (
    echo [FAIL] Inventor 2026 NOT FOUND!
)

if exist "C:\Program Files\Autodesk\Inventor 2026\Bin\Inventor.exe" (
    echo [OK] Inventor.exe found
) else (
    echo [FAIL] Inventor.exe NOT FOUND!
)
echo.

echo [4] Checking Inventor API assemblies...
echo.
if exist "C:\Program Files\Autodesk\Inventor 2026\Bin\Public Assemblies\Autodesk.Inventor.Interop.dll" (
    echo [OK] Inventor API found
) else (
    echo [WARNING] Inventor API DLL not found in expected location
)
echo.

echo [5] Checking for conflicts...
echo.
reg query "HKCU\Software\Autodesk\Inventor\Addins\{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}" > nul 2>&1
if %errorlevel% equ 0 (
    echo [INFO] Registry entry exists (normal after first load)
    reg query "HKCU\Software\Autodesk\Inventor\Addins\{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}" /v LoadBehavior
) else (
    echo [INFO] No registry entry (normal for first-time deployment)
)
echo.

echo [6] Manifest content check...
echo.
findstr /C:"LoadOnStartUp" "C:\ProgramData\Autodesk\Inventor 2026\Addins\AssemblyClonerAddIn.addin" > nul 2>&1
if %errorlevel% equ 0 (
    echo [OK] LoadOnStartUp setting found in manifest
) else (
    echo [WARNING] LoadOnStartUp not found
)

findstr /C:"AssemblyClonerAddIn.dll" "C:\ProgramData\Autodesk\Inventor 2026\Addins\AssemblyClonerAddIn.addin" > nul 2>&1
if %errorlevel% equ 0 (
    echo [OK] DLL reference found in manifest
) else (
    echo [FAIL] DLL reference missing in manifest!
)
echo.

echo ========================================
echo   DIAGNOSTIC COMPLETE
echo ========================================
echo.
echo NEXT STEPS TO LOAD THE ADD-IN:
echo.
echo 1. CLOSE Inventor completely
echo    - Close all Inventor windows
echo    - Check Task Manager (Ctrl+Shift+Esc)
echo    - End any "Inventor.exe" processes
echo.
echo 2. START Inventor fresh
echo.
echo 3. Check Add-Ins:
echo    - Go to Tools -^> Add-Ins
echo    - Look for "Assembly Cloner with iLogic Patcher"
echo    - Tick "Loaded" checkbox
echo    - Tick "Load on Startup" checkbox (optional)
echo.
echo 4. If still not visible:
echo    - Check Windows Event Viewer:
echo      eventvwr.msc -^> Windows Logs -^> Application
echo      Look for .NET Runtime or Inventor errors
echo.
echo    - Verify DLL architecture:
echo      Must be x64 (64-bit) for Inventor 2026
echo.
echo    - Check manifest GUID matches:
echo      GUID: {B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}
echo.
pause
