@echo off
copy "Assets\SpectivLauncher.cs" "SpectivLauncher.cs"
"C:\\Program Files\\Microsoft Visual Studio\\18\\Community\\MSBuild\\Current\\Bin\\Roslyn\\csc.exe" /target:winexe /out:SpectivLauncher.exe /win32icon:"Assets\icon.ico" SpectivLauncher.cs
del SpectivLauncher.cs
echo Build complete! New SpectivLauncher.exe created with updated Launch_UI.ps1.
pause
