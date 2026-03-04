@echo off
REM Nuclear Reopen Cycle Launcher
REM Author: Quintin de Bruin © 2026

echo.
echo ================================================
echo   NUCLEAR REOPEN CYCLE
echo ================================================
echo.
echo WARNING: This will CLOSE and REOPEN your assembly!
echo.
echo This is the LAST RESORT option when all other
echo methods fail to refresh BOM formulas.
echo.
echo Make sure:
echo   - Inventor is running
echo   - Assembly is open
echo   - ALL CHANGES ARE SAVED!
echo.
pause

cscript //nologo "%~dp0Nuclear_Reopen_Cycle.vbs"

echo.
echo ================================================
echo   PROCESS COMPLETE
echo ================================================
echo.
pause
