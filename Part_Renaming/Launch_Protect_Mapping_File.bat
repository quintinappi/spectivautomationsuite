@echo off
REM ==============================================================================
REM PROTECT MAPPING FILE - SET HIDDEN ATTRIBUTE
REM ==============================================================================
REM Makes STEP_1_MAPPING.txt hidden to protect it from accidental deletion
REM Scripts can still read and update it normally
REM ==============================================================================

title Protect Mapping File
echo.
echo ========================================
echo   PROTECT MAPPING FILE
echo ========================================
echo.
echo This will make STEP_1_MAPPING.txt hidden
echo to protect it from accidental deletion.
echo.
echo The file will still be fully accessible
echo to all renaming scripts.
echo.
pause

cscript //nologo "%~dp0Protect_Mapping_File.vbs"

echo.
echo ========================================
echo   Protection complete
echo ========================================
echo.
pause