@echo off
echo ========================================
echo REGISTRY MANAGEMENT: REGISTRY MANAGER
echo ========================================
echo This will manage the part numbering
echo counter database and mapping files.
echo.
echo Features:
echo - Scan counters for specific prefix
echo - Show all existing counters
echo - Clear entire database (with mapping file)
echo - Dynamic prefix support
echo.
echo Options:
echo - SCAN: Check current numbering state
echo - CLEAR: Reset all counters and mapping
echo.
pause
echo.
echo Running Registry Manager...
cscript //nologo "Registry_Manager.vbs"
echo.
pause