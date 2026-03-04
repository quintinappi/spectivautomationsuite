@echo off
REM Build AssemblyScannerPatcher.exe
REM Requires vbc.exe (VB.NET compiler)

if "%1"=="-nopause" goto nopause

echo.
echo === Building AssemblyScannerPatcher ===
echo.

:nopause

REM Check if vbc.exe exists
where vbc >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: vbc.exe not found in PATH
    echo.
    echo Please ensure .NET Framework is installed and vbc.exe is in your PATH
    echo Typical locations:
    echo   C:\Windows\Microsoft.NET\Framework\v4.0.30319\vbc.exe
    echo   C:\Windows\Microsoft.NET\Framework64\v4.0.30319\vbc.exe
    echo.
    pause
    exit /b 1
)

REM Set output path
set OUTPUT_DIR=%~dp0
set SOURCE_FILE=%~dp0AssemblyScannerPatcher.vb
set EXE_FILE=%OUTPUT_DIR%AssemblyScannerPatcher.exe

REM Compile AssemblyScannerPatcher
echo Compiling AssemblyScannerPatcher.vb...
vbc /target:exe /out:"%EXE_FILE%" /reference:System.dll "%SOURCE_FILE%" /optimize+

if %errorlevel% neq 0 (
    echo ERROR: Compilation failed
    pause
    exit /b 1
)

if exist "%EXE_FILE%" (
    echo SUCCESS: %EXE_FILE% created
    echo File size:
    for %%A in ("%EXE_FILE%") do echo   %%~zA bytes
) else (
    echo ERROR: EXE was not created
    pause
    exit /b 1
)

echo.
echo === Build Complete ===
echo.

if "%1" neq "-nopause" pause
exit /b 0
