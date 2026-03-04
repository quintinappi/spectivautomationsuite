@echo off
echo === FIX iLOGIC FORMULA WITH FORMAT ===
echo.
echo This will modify the iProperty formula to use Format() functions.
echo Changes: <parameter> -> Format(<parameter>, "0")
echo.
echo Make sure you have an ASSEMBLY open in Inventor.
echo.
echo WARNING: This will modify iProperty formulas in your plate parts!
echo Make sure you have backups.
echo.
echo Press any key to continue...
pause > nul
echo.
echo Running iLogic formula format fix...
cscript //nologo "Fix_iLogic_Formula_With_Format.vbs"
echo.
echo Format fix complete. Check BOM for whole number display.
echo.
pause