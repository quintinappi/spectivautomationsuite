@echo off
cls
echo ===============================================
echo TEST: Sheet Metal Converter on Single Part
echo ===============================================
echo.
echo This will convert the currently open part
echo to sheet metal with correct flat pattern orientation.
echo.
echo Make sure Part3 DM-UP.ipt is open in Inventor!
echo.
pause

cscript //nologo "TEST_New_Convert_Single_Part.vbs"

echo.
echo ===============================================
echo Test complete! Check Inventor for results.
echo ===============================================
pause
