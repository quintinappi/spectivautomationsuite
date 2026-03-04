@echo off
REM ******************************************************************************
REM PART GROUPING CORRECTION TOOL - LAUNCHER
REM ******************************************************************************
REM Purpose: Rename parts from one grouping to another (e.g., IPE -> B)
REM
REM This tool:
REM 1. Scans assembly for all parts and their groupings
REM 2. Lists unique groupings found (PL, B, CH, A, IPE, etc.)
REM 3. Lets you select which grouping to change
REM 4. Lets you select the target grouping
REM 5. Checks registry for next available numbers
REM 6. Renames parts using heritage method
REM 7. Updates all assembly references
REM 8. Updates IDW drawing references
REM
REM Use Case: When you need to reclassify parts after cloning
REM Example: Change all IPE parts to B grouping
REM
REM Date: January 21, 2026
REM ******************************************************************************

echo.
echo ============================================================
echo   PART GROUPING CORRECTION TOOL
echo ============================================================
echo.
echo This tool will:
echo   1. Scan your assembly and list all part groupings
echo   2. Let you select which grouping to change
echo   3. Let you select the new target grouping
echo   4. Rename parts using sequential numbering
echo   5. Update all assembly references
echo   6. Update IDW drawing references
echo.
echo IMPORTANT: Make sure Inventor is running with your assembly open!
echo ============================================================
echo.

pause

echo.
echo Starting Part Grouping Correction Tool...
echo.

cscript //nologo "%~dp0Part_Grouping_Correction.vbs"

if errorlevel 1 (
    echo.
    echo ERROR: Script failed. Check Grouping_Correction_Log.txt for details.
    pause
) else (
    echo.
    echo ============================================================
    echo   COMPLETED
    echo ============================================================
    echo.
    echo Check Grouping_Correction_Log.txt for full details.
    echo.
    pause
)
