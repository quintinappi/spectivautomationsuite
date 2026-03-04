# SpectivLauncher - EXE Launcher Instructions

## Overview
You now have an EXE launcher for the Inventor Automation Suite instead of using a batch file!

## Quick Start

### Step 1: Build the EXE
Run `Build_Launcher.bat` to compile SpectivLauncher.exe

**Requirements:**
- Windows with .NET Framework (already installed on most Windows systems)
- OR Visual Studio Developer Command Prompt

**If you get an error:**
1. Open Start Menu
2. Search for "Developer Command Prompt for VS"
3. Right-click and run as Administrator
4. Navigate to this folder
5. Run `Build_Launcher.bat`

### Step 2: Add Your Splash Image
1. Create a folder named `assets` in this directory
2. Place your splash image as `splash.png` in the `assets` folder
3. Recommended size: 1170x520 pixels (will be scaled to fit)

### Step 3: Add Custom Icon (Optional)
1. Place your `.ico` file in this folder as `icon.ico`
2. Run `Build_Launcher.bat` again
3. Your icon will be applied to SpectivLauncher.exe

### Step 4: Launch
Double-click `SpectivLauncher.exe` to launch the UI!

## File Structure
```
FINAL_PRODUCTION_SCRIPTS 1 Oct 2025/
│
├── SpectivLauncher.exe          <-- Your custom EXE launcher (after building)
├── SpectivLauncher.cs           <-- C# source code
├── Build_Launcher.bat           <-- Script to build the EXE
├── Launch_UI.bat                <-- Old batch launcher (can delete if EXE works)
├── Launch_UI.ps1                <-- PowerShell UI script
│
└── assets/                      <-- Create this folder
    └── splash.png               <-- Your splash screen image (optional)
```

## Icon Requirements
- Format: `.ico` file
- Recommended sizes: 256x256, 128x128, 64x64, 48x48, 32x32, 16x16
- Online converters: https://www.icoconverter.com/ or https://convertico.com/

## Troubleshooting

### "csc.exe not found" error
Use Visual Studio Developer Command Prompt (see Step 1 above)

### Splash image not showing
1. Make sure the `assets` folder exists
2. Make sure the image is named exactly `splash.png`
3. Check that the image file is not corrupted

### EXE doesn't launch
1. Ensure `Launch_UI.ps1` is in the same folder as the EXE
2. Try right-clicking and running as Administrator
3. Check Windows Defender isn't blocking the EXE

## Advantages of EXE Launcher
- Customizable icon
- No console window (fully hidden)
- Professional appearance
- Easy to distribute

## Rebuilding
Any time you make changes to SpectivLauncher.cs, simply run `Build_Launcher.bat` again to recreate the EXE.
