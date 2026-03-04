@echo off
echo === BOM IPROPERTY FORMULA INVESTIGATION ===
echo.
echo This will investigate how Stock Number iProperty formulas work.
echo Make sure you have an ASSEMBLY open in Inventor.
echo.
echo Press any key to continue...
pause > nul
echo.
echo Running iProperty formula investigation...
cscript //nologo "BOM_iProperty_Formula_Investigation.vbs"
echo.
echo Investigation complete. Check output for formula details.
echo.
pause