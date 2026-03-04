@echo off
REM ==============================================================================
REM IDW-ASSEMBLY SYNCHRONIZER - DYNAMIC REFERENCE FIXER
REM ==============================================================================
REM Syncs IDW references to match whatever the parent assembly uses
REM Handles edge cases where folder structure breaks normal reference updating
REM ==============================================================================

title IDW-Assembly Synchronizer
echo.
echo ==========================================
echo   IDW-ASSEMBLY SYNCHRONIZER
echo ==========================================
echo.
echo This tool syncs IDW references to match
echo whatever the parent assembly uses.
echo.
echo USE THIS WHEN:
echo - STEP 1 worked but IDWs still reference old files
echo - Main assembly in one folder, parts in another
echo - Folder structure prevents normal updating
echo.
echo IMPORTANT:
echo - Inventor must be running
echo - Will process all IDW files in selected folder
echo - Creates log file for troubleshooting
echo.
pause

cscript //nologo "%~dp0IDW_Assembly_Synchronizer.vbs"

echo.
echo ==========================================
echo   Sync complete
echo ==========================================
echo.
pause