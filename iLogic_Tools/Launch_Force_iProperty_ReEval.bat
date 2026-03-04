@echo off
echo === FORCE IPROPERTY RE-EVALUATION ===
echo.
echo This mimics manual Document Settings precision toggle.
echo Changes precision 0 -^> 3 -^> 0, then saves (no save between changes).
echo.
echo Make sure you have an ASSEMBLY open in Inventor.
echo.
pause
echo.
cscript //nologo "Force_iProperty_ReEval.vbs"
echo.
pause