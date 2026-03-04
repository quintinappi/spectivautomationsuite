@echo off
REM ==============================================================================
REM LAYER STYLE MANAGER - DETAILED LAUNCHER
REM ==============================================================================
REM This advanced tool lets you:
REM - Choose specific layers for Visible, Hidden, Center Lines, Center Marks
REM - Apply them to selected views
REM - Each curve type gets routed to its appropriate layer
REM ==============================================================================

echo.
echo ========================================
echo   LAYER STYLE MANAGER - DETAILED
echo ========================================
echo.
echo This tool will help you:
echo - Choose a layer for VISIBLE lines
echo - Choose a layer for HIDDEN lines
echo - Choose a layer for CENTER LINES (optional)
echo - Choose a layer for CENTER MARKS (optional)
echo - Apply to specific views or ALL views
echo.
echo Make sure Inventor is running with an IDW open!
echo.
pause

cscript //nologo "%~dp0Layer_Style_Manager_Detailed.vbs"

echo.
echo ========================================
echo Layer Style Manager Complete
echo ========================================
echo.
pause
