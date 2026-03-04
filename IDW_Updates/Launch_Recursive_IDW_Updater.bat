@echo off
echo ==============================================================================
echo RECURSIVE IDW UPDATER
echo ==============================================================================
echo.
echo This tool will:
echo 1. Recursively scan ALL folders for STEP_1_MAPPING.txt files
echo 2. Aggregate ALL mappings into comprehensive dictionary
echo 3. Recursively find ALL .idw files
echo 4. Update all IDWs using aggregated mappings
echo 5. Generate detailed report
echo.
echo Make sure Inventor is running!
echo.
pause

echo.
echo Starting Recursive IDW Updater...
echo.
echo.

cscript //nologo "%~dp0Recursive_IDW_Updater.vbs"

echo.
echo.
echo ========================================
echo Recursive IDW Updater Complete
echo ========================================
echo.
echo Check the log file in the Logs folder for details.
echo.
pause
