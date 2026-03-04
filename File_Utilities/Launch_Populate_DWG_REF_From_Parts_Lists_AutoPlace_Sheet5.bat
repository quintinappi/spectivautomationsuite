@echo off
REM ======================================================================
REM POPULATE DWG REF + AUTO-PLACE ALL MISSING PARTS
REM Sheet target is selected interactively from non-DXF sheets.
REM ======================================================================

pushd "%~dp0"
cscript //nologo "Populate_DWG_REF_From_Parts_Lists.vbs" /dwgrefonly:off /autoplace:on
popd

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: Script failed with error code %ERRORLEVEL%
    echo.
    pause
)
