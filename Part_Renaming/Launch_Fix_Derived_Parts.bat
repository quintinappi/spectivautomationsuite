@echo off
echo ========================================
echo FIX DERIVED PARTS (Post-Clone)
echo ========================================
echo This script fixes cloned assemblies that
echo have derived parts pointing to external
echo base files.
echo.
echo What it does:
echo - Scans open assembly for derived parts
echo - Finds base files outside assembly folder
echo - Copies base files locally with prefix
echo - Updates all derived references to local
echo.
echo When to use:
echo - After Assembly Cloner on models with
echo   derived parts (derived IPTs)
echo - When Check Derived Parts shows external
echo   base file locations
echo.
echo IMPORTANT:
echo - Open your CLONED assembly in Inventor
echo - Run this script
echo - Re-run if chained derivations exist
echo.
pause
echo.
echo Running Derived Parts Fixer...
cscript //nologo "Fix_Derived_Parts.vbs"
echo.
echo Derived Parts Fix Complete!
echo Check Fix_Derived_Log.txt for details.
echo.
pause
