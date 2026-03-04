@echo off
echo ========================================
echo STEP 2: IDW REFERENCE UPDATES
echo ========================================
echo This will update ALL sub-assembly IDW files
echo using assembly-by-assembly processing.
echo.
echo Features:
echo - Finds IDWs dynamically (any naming)
echo - Uses STEP 1's proven method
echo - Assembly-by-assembly (isolated failures)
echo - Safe: doesn't close Structure.iam
echo.
echo Make sure:
echo 1. STEP 1 (Part Renaming) completed
echo 2. Structure.iam is OPEN in Inventor
echo 3. STEP_1_MAPPING.txt exists
echo.
pause
echo.
echo Running IDW Reference Updater...
cscript //nologo "%~dp0IDW_Reference_Updater.vbs"
echo.
echo IDW Reference Updates Complete!
echo Check the log file for details.
echo.
pause