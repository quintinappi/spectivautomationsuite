@echo off
echo === FIX iLOGIC STOCK NUMBER FORMULA ===
echo.
echo This will modify the iProperty formula to use ROUND() functions.
echo Changes: <parameter> -> Round(<parameter>)
echo.
echo Make sure you have an ASSEMBLY open in Inventor.
echo.
echo WARNING: This will modify iProperty formulas in your plate parts!
echo Make sure you have backups.
echo.
echo Press any key to continue...
pause > nul
echo.
echo Running iLogic formula fix...
cscript //nologo "Fix_iLogic_Stock_Number_Formula.vbs"
echo.
echo Formula fix complete. Check BOM for whole number display.
echo.
pause