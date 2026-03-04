@echo off
cls
echo ===============================================
echo SHEET METAL CONVERTER - SINGLE PART
echo ===============================================
echo.
echo This tool converts a SINGLE PART to sheet metal
echo with correct flat pattern orientation.
echo.
echo REQUIREMENTS:
echo - Inventor must be running
echo - A PART document (.ipt) must be open
echo - Part should be a plate (solid body)
echo.
echo PROCESS:
echo 1. Finds the largest face (top/bottom, not edge)
echo 2. Converts part to sheet metal
echo 3. Creates flat pattern with CORRECT orientation
echo 4. Adds PLATE LENGTH and PLATE WIDTH formulas
echo 5. Saves the part
echo.
echo RESULT:
echo - Flat pattern shows large face (not 6mm edge)
echo - Custom iProperties with formulas:
echo   PLATE LENGTH = =^<SHEET METAL LENGTH^>
echo   PLATE WIDTH = =^<SHEET METAL WIDTH^>
echo.
echo ===============================================
echo.
pause

echo.
echo Starting conversion...
echo.

cscript //nologo "%~dp0Sheet_Metal_Part_Converter.vbs"

if errorlevel 1 (
    echo.
    echo ===============================================
    echo ERROR: Conversion failed!
    echo Check the error message above.
    echo ===============================================
    pause
    exit /b 1
)

echo.
echo ===============================================
echo Conversion complete!
echo Check Inventor for the converted part.
echo ===============================================
pause
