@echo off
echo === BOM NUCLEAR DEEP INVESTIGATION ===
echo.
echo This diagnostic will investigate BOM display formatting in depth.
echo Make sure you have an ASSEMBLY open in Inventor (not a part).
echo.
echo Press any key to continue...
pause > nul
echo.
echo Running diagnostic...
cscript //nologo "BOM_Nuclear_Deep_Investigation.vbs"
echo.
echo Diagnostic complete. Check BOM_NUCLEAR_DIAGNOSTIC_REPORT.txt for results.
echo.
pause