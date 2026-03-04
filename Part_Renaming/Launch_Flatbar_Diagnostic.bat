@echo off
echo ========================================
echo  FLATBAR DIAGNOSTIC TOOL
echo ========================================
echo.
echo This will scan your active assembly and find all potential flatbars,
echo showing you exactly how they're being classified.
echo.
echo Make sure you have:
echo  1. Inventor running
echo  2. An assembly open
echo.
pause

cscript //nologo "Diagnose_Flatbars.vbs"

echo.
pause
