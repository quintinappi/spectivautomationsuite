@echo off
REM ==============================================================================
REM LAYER STYLE MANAGER LAUNCHER
REM ==============================================================================
REM This tool lets you:
REM - See all available layers in your drawing
REM - Choose which layer to apply
REM - Apply it to selected views (or ALL views)
REM ==============================================================================

echo.
echo ========================================
echo   LAYER STYLE MANAGER
echo ========================================
echo.
echo This tool will help you:
echo - List all available layers in your drawing
echo - Select a layer to apply
echo - Apply it to specific views or ALL views
echo.
echo Make sure Inventor is running with an IDW open!
echo.
pause

cscript //nologo "%~dp0Layer_Style_Manager.vbs"

echo.
echo ========================================
echo Layer Style Manager Complete
echo ========================================
echo.
pause
