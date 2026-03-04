@echo off
echo =========================================================
echo iLOGIC SCANNER
echo =========================================================
echo.
echo This tool scans your Inventor document for iLogic rules.
echo.
echo Features:
echo   - Detect iLogic rules in current document
echo   - Display rule names and source code
echo   - Export rules to external text files
echo   - Scan referenced parts in assemblies
echo.
echo Make sure your assembly/part is open in Inventor!
echo.
pause

cd /d "%~dp0"
cscript //nologo "iLogic_Scanner.vbs"

echo.
echo =========================================================
echo Script completed. Press any key to exit.
pause >nul
