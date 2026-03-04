@echo off
REM ==================================================================================
REM IDW PART PLACER LAUNCHER
REM ==================================================================================
REM Automatically creates base views for ALL parts from the assembly on Sheet 1
REM Places them on Sheet 2 (and Sheet 3 if needed) in a grid layout
REM ==================================================================================

echo ========================================
echo IDW PART PLACER
echo ========================================
echo.
echo This will place ALL parts from the assembly
echo on Sheet 1 onto Sheet 2 (and Sheet 3 if needed).
echo.
echo Make sure Inventor is running with your
echo IDW drawing open (with assembly on Sheet 1)!
echo.
echo Press any key to continue...
pause >nul

echo.
echo Running IDW Part Placer...
echo.

cscript //nologo "%~dp0IDW_Part_Placer.vbs"

echo.
echo Press any key to exit...
pause >nul
