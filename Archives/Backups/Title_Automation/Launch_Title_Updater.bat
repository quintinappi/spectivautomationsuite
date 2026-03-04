@echo off
echo ========================================
echo TITLE AUTOMATION: TITLE UPDATER
echo ========================================
echo This will automatically update IDW
echo drawing titles based on assembly names
echo and view information.
echo.
echo Features:
echo - Automatic title generation
echo - Base view detection
echo - Assembly name integration
echo - Batch processing of IDW files
echo.
echo Make sure:
echo 1. IDW files are accessible
echo 2. Inventor is running
echo 3. Assembly renaming completed
echo.
pause
echo.
echo Running Title Updater...
cscript //nologo "%~dp0Title_Updater.vbs"
echo.
echo Title Updates Complete!
echo Check the log file for details.
echo.
pause