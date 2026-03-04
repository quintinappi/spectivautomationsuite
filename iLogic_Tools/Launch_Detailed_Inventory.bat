@echo off
REM Launcher for Detailed_Assembly_Inventory.vbs

setlocal enabledelayedexpansion

REM Get the script directory
set SCRIPT_DIR=%~dp0

REM Run the VBScript
cscript.exe "!SCRIPT_DIR!Detailed_Assembly_Inventory.vbs"

REM Pause to see output
pause
