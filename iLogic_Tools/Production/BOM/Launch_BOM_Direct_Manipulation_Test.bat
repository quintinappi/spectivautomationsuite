@echo off
echo === BOM DIRECT MANIPULATION TEST ===
echo.
echo This will try to set BOM display properties directly.
echo Make sure you have an ASSEMBLY open in Inventor.
echo.
echo Press any key to continue...
pause > nul
echo.
echo Running direct manipulation test...
cscript //nologo "BOM_Direct_Manipulation_Test.vbs"
echo.
echo Test complete. Check BOM in Inventor for results.
echo.
pause