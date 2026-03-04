@echo off
cls
echo =========================================================
echo FIX WIDTH AND LENGTH PARAMETERS ON BOM FOR PLATE PARTS
echo =========================================================
echo.
echo This tool adds WIDTH and LENGTH custom iProperty columns
echo to the BOM and populates them with sheet metal flat pattern
echo dimensions for plate parts only.
echo.
echo REQUIREMENTS:
echo - Inventor must be running
echo - An ASSEMBLY document (.iam) must be open
echo - Plate parts must be converted to sheet metal (have flat patterns)
echo.
echo PLATE DETECTION:
echo Parts are identified as plates if their Description contains:
echo   - "PL" (e.g., PL10, 10PL, etc.)
echo   - "VRN" (Vloer/Roof/N plates)
echo   - "S355JR" (structural steel grade)
echo.
echo PROCESS:
echo 1. Scans assembly for plate parts
echo 2. For each plate part with a flat pattern:
echo    - Creates/updates LENGTH custom iProperty = ^<sheet metal length^>
echo    - Creates/updates WIDTH custom iProperty = ^<sheet metal width^>
echo    - Saves the part
echo 3. Refreshes assembly BOM
echo.
echo AFTER RUNNING:
echo 1. Open BOM in your assembly
echo 2. Right-click column header ^> "Add Custom iProperty Columns"
echo 3. Add "LENGTH" and "WIDTH" columns (Type: Text)
echo 4. Plate parts will show their flat pattern dimensions
echo.
echo =========================================================
echo.
pause

echo.
echo Running Fix BOM Plate Dimensions script...
echo.

cscript //nologo "%~dp0Fix_BOM_Plate_Dimensions.vbs"

echo.
echo =========================================================
echo Script completed!
echo =========================================================
echo.
pause
