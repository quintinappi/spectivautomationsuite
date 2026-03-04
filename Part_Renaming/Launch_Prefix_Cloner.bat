@echo off
echo ========================================
echo PREFIX CLONER (Prefix Changer Only)
echo ========================================
echo This will copy an assembly with ALL its
echo parts to a new location, replacing ONLY
echo the filename PREFIX (keeping suffixes).
echo.
echo Features:
echo - Copies assembly + all referenced parts
echo - Detects common prefix automatically
echo - Replaces ONLY the prefix in filenames
echo - Keeps part suffixes intact (B1, PL2, etc.)
echo - Updates all assembly and IDW references
echo - Generates STEP_1_MAPPING.txt
echo.
echo Example:
echo   N1SCR04-780-B1.IPT  becomes  N2SCR04-780-B1.IPT
echo   (Only the N1 prefix changes to N2)
echo.
echo Perfect for:
echo - Creating another section in the plant
echo - Duplicating assemblies for different areas
echo - Keeping consistent part naming conventions
echo.
echo Make sure Inventor is running with your
echo SOURCE assembly open!
echo.
pause
echo.
echo Running Prefix Cloner...
cscript //nologo "Prefix_Cloner.vbs"
echo.
echo Prefix Cloning Complete!
echo Check the log file and STEP_1_MAPPING.txt for details.
echo.
pause
