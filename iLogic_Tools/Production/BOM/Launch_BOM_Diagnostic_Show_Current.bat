@echo off
echo === BOM DIAGNOSTIC - SHOW CURRENT STATE ===
echo.
echo This will show current parameter values and iProperty formulas.
echo No modifications - just investigation.
echo.
echo Make sure you have an ASSEMBLY open in Inventor.
echo.
echo Press any key to continue...
pause > nul
echo.
echo Running diagnostic...
cscript //nologo "BOM_Diagnostic_Show_Current.vbs"
echo.
echo Diagnostic complete. Check output above.
echo.
pause