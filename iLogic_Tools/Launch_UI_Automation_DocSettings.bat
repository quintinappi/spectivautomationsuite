@echo off
echo === UI AUTOMATION - DOCUMENT SETTINGS ===
echo.
echo This will open each plate part and use SendKeys to:
echo 1. Open Document Settings dialog
echo 2. Toggle precision up/down
echo 3. Save
echo.
echo WARNING: Do NOT touch keyboard/mouse while this runs!
echo The script uses UI automation and needs control.
echo.
echo Make sure Inventor is visible on screen.
echo.
pause
echo.
cscript //nologo "UI_Automation_DocSettings.vbs"
echo.
pause