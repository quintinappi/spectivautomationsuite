@echo off
echo ========================================
echo FILE UTILITIES: UNUSED PART FINDER
echo ========================================
echo This tool will scan your project and
echo identify IPT files that are NOT used
echo in your main assembly.
echo.
echo Features:
echo - Scans open assembly for all parts
echo - Scans folders for ALL IPT files
echo - Identifies unused/orphaned files
echo - Moves unused files to backup folder
echo.
echo Use this to clean up your project
echo folder after renaming parts.
echo.
echo IMPORTANT: Make sure your main
echo assembly is OPEN in Inventor!
echo.
pause
echo.
echo Running Unused Part Finder...
cscript //nologo "Unused_Part_Finder.vbs"
echo.
echo Operation Complete!
echo Check the log file for details.
echo.
pause
