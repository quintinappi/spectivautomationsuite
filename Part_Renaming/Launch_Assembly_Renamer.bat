@echo off
echo ========================================
echo PART RENAMING: ASSEMBLY RENAMER
echo ========================================
echo This will rename all parts in the entire
echo assembly hierarchy using dynamic prefixes
echo and create mapping file for IDW updates.
echo.
echo Features:
echo - User-defined project prefix support
echo - Intelligent part grouping (PL, B, CH, A, etc.)
echo - Registry-based counter persistence
echo - Full assembly hierarchy processing
echo.
echo Make sure Inventor is running with your
echo main assembly open!
echo.
pause
echo.
echo Running Assembly Renamer...
cscript //nologo "Assembly_Renamer.vbs"
echo.
echo Assembly Renaming Complete!
echo Check the log file for details.
echo Mapping file created: STEP_1_MAPPING.txt
echo.
pause