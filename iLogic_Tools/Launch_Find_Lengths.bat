@echo off
echo Please make sure the ASSEMBLY is open in Inventor (not a part)
echo Press any key to continue...
pause >nul
cscript.exe "%~dp0Find_Length_By_Max_Value.vbs"
pause
