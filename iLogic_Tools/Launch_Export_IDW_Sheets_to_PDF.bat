@echo off
REM Launcher for Export IDW Sheets to PDF
REM Exports each sheet of the open IDW to a separate PDF with numbered file names

echo =========================================================
echo EXPORT IDW SHEETS TO PDF
echo =========================================================
echo.
echo This script will:
echo   1. Check the open IDW file name
echo   2. Export each sheet as a separate PDF
echo   3. Name them as [BaseName]-1.pdf, [BaseName]-2.pdf, etc.
echo.
echo REQUIREMENTS:
echo   - Inventor must be running
echo   - A drawing (IDW) must be open
echo.
pause

cscript.exe "Export_IDW_Sheets_to_PDF.vbs"

echo.
pause