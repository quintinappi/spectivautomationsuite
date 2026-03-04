@echo off
REM ==============================================================================
REM SMART PREFIX SCANNER LAUNCHER
REM ==============================================================================
REM Use this BEFORE running STEP 1 on new assemblies to prevent duplicate numbers
REM ==============================================================================

title Smart Prefix Scanner
echo.
echo ========================================
echo   SMART PREFIX SCANNER
echo ========================================
echo.
echo This tool scans your existing model to detect
echo the highest part numbers used, then updates
echo the Registry so STEP 1 continues correctly.
echo.
echo WHEN TO USE:
echo - Before adding new assemblies like Access Walkway
echo - When Registry has been cleared but files renamed
echo - To prevent duplicate part numbers
echo.
echo Make sure your main assembly is open in Inventor!
echo.
pause

cscript //nologo "%~dp0Smart_Prefix_Scanner.vbs"

echo.
echo ========================================
echo   Scanner completed
echo ========================================
echo.
pause