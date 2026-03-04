@echo off
echo ==============================================================================
echo MAPPING RECOVERY TOOL
echo ==============================================================================
echo.
echo This tool will attempt to recover the mapping file for a renamed assembly.
echo.
echo STEPS:
echo 1. Open your renamed assembly in Inventor
echo 2. This script will scan the assembly
echo 3. Create STEP_1_MAPPING.txt based on OldVersions folder
echo 4. Then you can run IDW Reference Updater
echo.
pause

echo.
echo Starting Mapping Recovery...
echo.
echo.

cscript //nologo "%~dp0Recover_Mapping.vbs"

echo.
echo.
echo ========================================
echo Mapping Recovery Complete
echo ========================================
echo.
echo Check assembly directory for STEP_1_MAPPING.txt
echo.
pause
