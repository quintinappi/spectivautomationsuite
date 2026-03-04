@echo off
echo === BOM iPROPERTIES INVESTIGATION ===
echo.
echo This will check if BOM uses iProperties instead of parameters.
echo Make sure you have an ASSEMBLY open in Inventor.
echo.
echo Press any key to continue...
pause > nul
echo.
echo Running iProperties investigation...
cscript //nologo "BOM_iProperties_Investigation.vbs"
echo.
echo Investigation complete. Check output for iProperty mappings.
echo.
pause