@echo off
echo ========================================
echo PART CLONER
echo ========================================
echo This will copy a single part (.ipt) to a
echo new isolated location.
echo.
echo Features:
echo - Copies individual part file
echo - Reads and displays iProperties
echo - Optional renaming
echo - Creates isolated copy
echo.
echo Perfect for:
echo - Creating part variants
echo - Isolating parts for modification
echo - Backup copies of parts
echo.
echo Make sure Inventor is running with your
echo SOURCE part open!
echo.
pause
echo.
echo Running Part Cloner...
cscript //nologo "Part_Cloner.vbs"
echo.
echo Part Cloning Complete!
echo Check the log file for details.
echo.
pause