@echo off
echo === FIX BOM IPROPERTY FORMULAS ===
echo.
echo This will update Stock Number formulas to use ROUND() for whole numbers.
echo Make sure you have an ASSEMBLY open in Inventor.
echo.
echo WARNING: This will modify iProperty formulas in your plate parts!
echo Make sure you have backups.
echo.
echo Press any key to continue...
pause > nul
echo.
echo Running iProperty formula fix...
cscript //nologo "Fix_BOM_iProperty_Formulas.vbs"
echo.
echo Fix complete. Check BOM for whole number display.
echo.
pause