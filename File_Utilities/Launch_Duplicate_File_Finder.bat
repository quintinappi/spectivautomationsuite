@echo off
echo ========================================
echo FILE UTILITIES: DUPLICATE FILE FINDER
echo ========================================
echo This will scan directories for duplicate
echo Inventor files that might cause issues
echo during the renaming process.
echo.
echo Features:
echo - Recursive directory scanning
echo - Inventor file type detection (.ipt, .iam, .idw)
echo - Duplicate identification
echo - Detailed reporting
echo.
echo Use this before running the Assembly
echo Renamer to identify potential conflicts.
echo.
pause
echo.
echo Running Duplicate File Finder...
cscript //nologo "Duplicate_File_Finder.vbs"
echo.
echo Duplicate Scan Complete!
echo Check the results for any issues.
echo.
pause