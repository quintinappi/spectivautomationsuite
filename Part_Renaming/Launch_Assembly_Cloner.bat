@echo off
echo ========================================
echo ASSEMBLY CLONER
echo ========================================
echo This will copy an assembly with ALL its
echo parts to a new isolated location.
echo.
echo Features:
echo - Copies assembly + all referenced parts
echo - Updates references to use local copies
echo - Optional heritage renaming (PL, CH, B, etc.)
echo - Copies and updates IDW drawings
echo - Creates fully isolated clone
echo.
echo Perfect for:
echo - Creating variants without cross-references
echo - Copying assemblies to new projects
echo - Isolating assemblies for modification
echo.
echo Make sure Inventor is running with your
echo SOURCE assembly open!
echo.
pause
echo.
echo Running Assembly Cloner...
cscript //nologo "Assembly_Cloner.vbs"
echo.
echo Assembly Cloning Complete!
echo Check the log file for details.
echo.
pause
