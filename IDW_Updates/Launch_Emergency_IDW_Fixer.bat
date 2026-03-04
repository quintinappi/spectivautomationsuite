@echo off
REM ==============================================================================
REM EMERGENCY IDW FIXER LAUNCHER
REM ==============================================================================
REM Use this when STEP 2 misses specific folders like "Launder 2" or "Extension"
REM ==============================================================================

title Emergency IDW Fixer
echo.
echo ========================================
echo   EMERGENCY IDW FIXER
echo ========================================
echo.
echo This tool fixes IDW files in specific folders
echo when STEP 2 misses them.
echo.
echo Make sure Inventor is running!
echo.
pause

cscript //nologo "%~dp0Emergency_IDW_Fixer.vbs"

echo.
echo ========================================
echo   Emergency fixer completed
echo ========================================
echo.
pause