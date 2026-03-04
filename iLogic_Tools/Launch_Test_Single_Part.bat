@echo off
cls
echo =========================================================
echo SHEET METAL CONVERSION - SINGLE PART TEST
echo =========================================================
echo.
echo This script tests the sheet metal conversion workflow on ONE part.
echo.
echo REQUIREMENTS:
echo - Inventor must be running
echo - A PART document (.ipt) must be open (not assembly)
echo - Part should be a plate with "PL" or "S355JR" in description
echo.
echo WHAT THIS TEST DOES:
echo 1. Checks if part is already sheet metal
echo 2. Converts to sheet metal (if needed)
echo 3. Sets correct thickness from description
echo 4. Creates flat pattern
echo 5. Extracts and displays flat pattern dimensions
echo.
echo This test will MODIFY the open part document!
echo Make sure you have a backup if needed.
echo.
pause

echo.
echo Running test script...

REM Check if Inventor is running
tasklist /FI "IMAGENAME eq Inventor.exe" 2>NUL | find /I /N "Inventor.exe">NUL
if "%ERRORLEVEL%"=="1" (
    echo ERROR: Inventor is not running. Please start Inventor first.
    echo.
    pause
    exit /b 1
)

echo Inventor is running. Executing test...
echo.

REM Run the test script
cscript //nologo "TEST_Single_Part_Conversion.vbs"

echo.
echo Test completed. Check the log file for detailed results.
echo.
pause
