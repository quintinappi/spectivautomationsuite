@echo off
cls
echo =========================================================
echo FIX NON-PLATE PARTS - ADD LENGTH2 PARAMETER
echo =========================================================
echo.
echo This tool will:
echo - Scan the open assembly for all non-plate parts
echo - Find parts without Length or Length2 parameters
echo - Add Length2 parameter linked to the longest dimension
echo.
echo REQUIREMENTS:
echo - Open the assembly (.iam) in Inventor first
echo - Parts must be saved and accessible
echo.
echo NOTE: You must manually enable the Export checkbox
echo       for Length2 in each part's Parameters dialog after.
echo.
pause

echo.
echo Starting Fix Non-Plate Parts tool...
cscript.exe //nologo "%~dp0Fix_All_Non_Plate_Parts.vbs"

echo.
echo =========================================================
echo.
pause
