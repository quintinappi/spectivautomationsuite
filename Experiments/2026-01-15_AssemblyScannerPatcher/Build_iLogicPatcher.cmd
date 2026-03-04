@echo off
REM Build iLogicPatcher.dll
REM Requires vbc.exe (VB.NET compiler)

if "%1"=="-nopause" goto nopause

echo.
echo === Building iLogicPatcher.dll ===
echo.

:nopause

REM Check if vbc.exe exists
where vbc >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: vbc.exe not found in PATH
    echo.
    echo Please ensure .NET Framework is installed and vbc.exe is in your PATH
    pause
    exit /b 1
)

REM Set paths
set OUTPUT_DIR=%~dp0
set SOURCE_FILE=%~dp0iLogicPatcher.vb
set DLL_FILE=%OUTPUT_DIR%iLogicPatcher.dll

REM Find Inventor interop assemblies
set INVENTOR_ASSEMBLY=""
set "AUTODESK_INVENTOR_PATH=C:\Program Files\Autodesk\Inventor 2026\Bin"
set "INVENTOR_REFERENCES="

REM Try to find Inventor assemblies in multiple locations
if exist "%AUTODESK_INVENTOR_PATH%\Autodesk.Inventor.Interop.dll" (
    set INVENTOR_REFERENCES=/reference:"%AUTODESK_INVENTOR_PATH%\Autodesk.Inventor.Interop.dll"
) else (
    echo ERROR: Could not find Autodesk.Inventor.Interop.dll
    echo Expected location: %AUTODESK_INVENTOR_PATH%
    echo.
    echo Please modify the AUTODESK_INVENTOR_PATH variable in this batch file
    echo to point to your Inventor installation directory
    pause
    exit /b 1
)

REM Compile as library
echo Compiling iLogicPatcher.vb as DLL...
vbc /target:library /out:"%DLL_FILE%" /reference:System.dll %INVENTOR_REFERENCES% "%SOURCE_FILE%" /optimize+

if %errorlevel% neq 0 (
    echo ERROR: Compilation failed
    pause
    exit /b 1
)

if exist "%DLL_FILE%" (
    echo SUCCESS: %DLL_FILE% created
    echo File size:
    for %%A in ("%DLL_FILE%") do echo   %%~zA bytes
) else (
    echo ERROR: DLL was not created
    pause
    exit /b 1
)

echo.
echo === Build Complete ===
echo.

if "%1" neq "-nopause" pause
exit /b 0
