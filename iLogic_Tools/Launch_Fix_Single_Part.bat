@echo off
echo ================================================
echo FIX SINGLE PART - ADD LENGTH2 PARAMETER
echo ================================================
echo.
echo INSTRUCTIONS:
echo 1. Open FL25 part in Inventor (NSCR05-780-FL25.ipt)
echo 2. Press any key to run the fix
echo.
pause

cscript.exe "%~dp0Fix_Single_Part_Length2.vbs"

echo.
echo ================================================
echo Check the Parameters dialog in Inventor to verify
echo Save the part if everything looks correct
echo ================================================
pause
