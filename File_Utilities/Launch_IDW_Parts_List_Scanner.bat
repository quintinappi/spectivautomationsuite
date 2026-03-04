@echo off
REM ==============================================================================
REM IDW PARTS LIST SCANNER - LAUNCHER
REM ==============================================================================
REM
REM This tool:
REM 1. Scans the currently open IDW drawing
REM 2. Extracts all parts from the parts list (BOM)
REM 3. Lists all referenced parts
REM 4. Finds ALL .ipt files in the same folder as the IDW
REM 5. Moves parts NOT in the parts list to "Unrenamed Parts" folder
REM
REM ==============================================================================

pushd "%~dp0"
cscript //nologo "IDW_Parts_List_Scanner.vbs"
popd

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: Script failed with error code %ERRORLEVEL%
    echo.
    pause
)
