@echo off
REM Master build script for AssemblyScannerPatcher
REM This builds both iLogicPatcher.dll and AssemblyScannerPatcher.exe

echo.
echo ========================================
echo   Building AssemblyScannerPatcher
echo   Experimental Tool - 2026-01-15
echo ========================================
echo.

REM Check if vbc.exe exists
where vbc >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: vbc.exe not found in PATH
    echo.
    echo Please add .NET Framework SDK to your PATH:
    echo   C:\Windows\Microsoft.NET\Framework64\v4.0.30319
    pause
    exit /b 1
)

REM Step 1: Build iLogicPatcher.dll
echo.
echo [1/3] Building iLogicPatcher.dll...
call "%~dp0Build_iLogicPatcher.cmd" -nopause
if %errorlevel% neq 0 (
    echo ERROR: Failed to build iLogicPatcher.dll
    pause
    exit /b 1
)

REM Step 2: Check if DLL was created
echo.
echo [2/3] Verifying iLogicPatcher.dll...
if not exist "%~dp0iLogicPatcher.dll" (
    echo ERROR: iLogicPatcher.dll was not created
    echo The EXE will not be able to load iLogic patching functionality
    pause
    exit /b 1
)
echo SUCCESS: iLogicPatcher.dll found

REM Step 3: Build AssemblyScannerPatcher.exe
echo.
echo [3/3] Building AssemblyScannerPatcher.exe...
call "%~dp0Build_Scanner.cmd" -nopause
if %errorlevel% neq 0 (
    echo ERROR: Failed to build AssemblyScannerPatcher.exe
    pause
    exit /b 1
)

REM Summary
echo.
echo ========================================
echo   BUILD SUCCESSFUL!
echo ========================================
echo.
echo Output files:
echo   %~dp0iLogicPatcher.dll
echo   %~dp0AssemblyScannerPatcher.exe
echo.
echo Next steps:
echo   1. Test the scanner on a sample assembly
echo   2. Verify iLogic patching works correctly
echo   3. Run comparison reports
echo.
echo See README.txt for usage instructions
echo.
pause
