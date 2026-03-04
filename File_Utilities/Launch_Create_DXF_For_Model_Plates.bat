@echo off
REM ======================================================================
REM CREATE DXF FOR MODEL PLATES
REM ======================================================================

pushd "%~dp0"
cscript //nologo "Create_DXF_For_Model_Plates.vbs"
popd

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: Script failed with error code %ERRORLEVEL%
    echo.
    pause
)
