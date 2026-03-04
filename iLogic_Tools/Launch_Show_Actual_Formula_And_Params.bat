@echo off
echo === SHOW ACTUAL FORMULA AND PARAMETERS ===
echo.
echo This will show the exact iProperty formula and parameter names.
echo No modifications - just investigation.
echo.
echo Make sure you have an ASSEMBLY open in Inventor.
echo.
echo Press any key to continue...
pause > nul
echo.
echo Running diagnostic...
cscript //nologo "Show_Actual_Formula_And_Params.vbs"
echo.
echo Check output above for actual formula and parameter names.
echo.
pause