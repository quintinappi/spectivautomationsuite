@echo off
echo === BOM SIMPLE INVESTIGATION ===
echo.
echo This diagnostic will investigate BOM display formatting.
echo Make sure you have an ASSEMBLY open in Inventor (not a part).
echo.
echo Press any key to continue...
pause > nul
echo.
echo Running simplified diagnostic...
cscript //nologo "BOM_Simple_Investigation.vbs"
echo.
echo Diagnostic complete. Check BOM_NUCLEAR_DIAGNOSTIC_REPORT.txt for results.
echo.
pause