@echo off
REM ==============================================================================
REM POPULATE DWG REF FROM PARTS LISTS - LAUNCHER
REM ==============================================================================
REM
REM This tool:
REM 1. Scans all sheets in the active IDW
REM 2. Reads all parts from all parts lists
REM 3. Builds a unique DWG REF value
REM 4. Writes DWG REF to drawing properties/title block prompts
REM
REM ==============================================================================

pushd "%~dp0"
cscript //nologo "Populate_DWG_REF_From_Parts_Lists.vbs" /dwgrefonly:on /autoplace:off
popd

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: Script failed with error code %ERRORLEVEL%
    echo.
    pause
)
