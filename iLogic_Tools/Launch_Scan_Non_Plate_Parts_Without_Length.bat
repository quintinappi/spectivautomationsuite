@echo off
REM Launcher for Scan_Non_Plate_Parts_Without_Length.vbs
REM Scans open assembly for non-plate parts without Length parameter

setlocal enabledelayedexpansion

REM Get the script directory
set SCRIPT_DIR=%~dp0

REM Run the VBScript
cscript.exe "!SCRIPT_DIR!Scan_Non_Plate_Parts_Without_Length.vbs"

REM Pause to see output
pause
