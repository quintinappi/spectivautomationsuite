# Spectiv Inventor Suite - Build & Test Guide

**Status:** Ready for Visual Studio Compilation
**Date:** 2025-01-21

---

## 📋 Prerequisites

### Required Software
1. **Visual Studio 2019 or later** (VB.NET support required)
2. **Inventor 2024, 2025, or 2026** - Any version for testing
3. **Inventor SDK** - Optional but recommended

### Inventor API References
You need to reference these DLLs from Inventor:
- `Inventor.Interop.dll` (from Inventor bin folder)
- Location: `C:\Program Files\Autodesk\Inventor 2026\Bin\`

**Note:** If you don't have the SDK, you can copy the interop DLL from Inventor's bin folder.

---

## 🏗️ Creating the Visual Studio Project

### Step 1: Create New Project
```
1. Open Visual Studio
2. Create New Project
3. Template: Class Library (.NET Framework)
4. Name: SpectivInventorSuite
5. Framework: .NET Framework 4.8 or later
```

### Step 2: Add References
```
1. Right-click project → Add → Reference
2. Browse to: C:\Program Files\Autodesk\Inventor 2026\Bin\
3. Add: Inventor.Interop.dll

4. Add .NET references:
   - System.Windows.Forms
   - Microsoft.VisualBasic (for InputBox)
```

### Step 3: Add Source Files
```
Copy these files to your project folder from Migration to Add-In/src/:
├── SpectivInventorSuiteAddIn.vb    (Add-In entry point)
├── AssemblyClonerForm.vb           (UI Form)
├── AssemblyCloner.vb               (Main logic)
├── Logger.vb                        (Logging)
├── FileHelper.vb                    (File operations)
├── RegistryManager.vb               (Registry)
└── PartClassifier.vb                (Classification)
```

### Step 4: Configure Build
```
1. Project Properties → Application
2. Output type: Class Library
3. Root namespace: SpectivInventorSuite
4. Build → Configuration: Release

5. Build → Advanced
   - Option Explicit: On
   - Option Strict: On
   - Option Compare: Binary
```

### Step 5: Build
```
Build → Build Solution
Output: SpectivInventorSuite.dll
```

---

## 📦 Installation

### Method 1: Manual Installation (For Testing)
```
1. Create folder: %APPDATA%\Autodesk\ApplicationPlugins\SpectivInventorSuite\
2. Copy files:
   - SpectivInventorSuite.dll (from build output)
   - SpectivInventorSuite.addin (from Migration folder)

3. Restart Inventor
4. Check "Assembly" tab for "Assembly Cloner" button
```

### Method 2: Deploy Script (Automated)
Create `Install.bat`:
```batch
@echo off
echo Installing Spectiv Inventor Suite...

REM Create plugin folder
mkdir "%APPDATA%\Autodesk\ApplicationPlugins\SpectivInventorSuite" 2>nul

REM Copy files
copy /Y SpectivInventorSuite.dll "%APPDATA%\Autodesk\ApplicationPlugins\SpectivInventorSuite\"
copy /Y SpectivInventorSuite.addin "%APPDATA%\Autodesk\ApplicationPlugins\SpectivInventorSuite\"

echo Installation complete!
echo Restart Inventor to use the Add-In.
pause
```

---

## 🧪 Testing Checklist

### Pre-Test Preparation
- [ ] Backup your test assembly
- [ ] Have a test assembly with:
  - [ ] 3+ sub-assemblies
  - [ ] 10+ parts
  - [ ] 5+ IDW drawings in various folders
  - [ ] Mixed part types (PL, B, CH, A, FL)

### Test Cases

| Test | Steps | Expected Result |
|------|-------|-----------------|
| **Load Add-In** | Start Inventor | Ribbon button appears in Assembly tab |
| **Open Form** | Click ribbon button | Form opens with current assembly info |
| **Select Destination** | Click Browse, choose folder | Destination path displayed |
| **Clone (No Rename)** | Click Clone | Files copied, references updated |
| **Clone (With Rename)** | Enable rename, enter prefix, Click Clone | Parts renamed with heritage naming |
| **Progress Update** | During clone operation | Progress bar updates, log scrolls |
| **Completion** | Clone finishes | Success message, folder opens |
| **IDW Updates** | Open cloned IDW | References point to new parts |
| **Mapping File** | Check destination folder | STEP_1_MAPPING.txt exists |

### Validation Tests
```
1. Open cloned main assembly in Inventor
2. Check Model Browser - all parts should resolve (red x = fail)
3. Open each IDW drawing
4. Check each drawing references new parts
5. Verify STEP_1_MAPPING.txt contains correct mappings
```

---

## 🐛 Troubleshooting

### "Button doesn't appear"
**Solution:**
1. Check Add-In is installed in correct folder
2. Check .addin file has correct CLSID matching VB.NET GUID
3. Restart Inventor
4. Check Add-In Manager: Tools → Application Options → Add-In Manager

### "Class not registered"
**Solution:**
1. Ensure Inventor.Interop.dll is referenced
2. Check Copy Local = True for Inventor reference
3. Rebuild solution

### "File not found exceptions"
**Solution:**
1. Check test assembly exists
2. Check destination folder is accessible
3. Run Inventor as Administrator if needed

### "References don't update"
**Solution:**
1. Check SilentOperation is enabled
2. Check all parts are loaded before assemblies
3. Verify paths in copiedFiles dictionary are correct

### "IDW files not found"
**Solution:**
1. Check CollectIDWFiles() ran successfully
2. Verify IDW files exist in source folder
3. Check FindIDWFilesRecursive() scans subfolders

---

## 📊 Expected Test Results

### Successful Clone Output
```
========================================
ASSEMBLY CLONER STARTING
========================================
DETECTED: Structure.iam
DETECTED: Full path: C:\Test\Structure.iam
DETECTED: Occurrences - 50

COLLECT: Scanning assembly for all referenced parts...
COLLECT: PART Column-1.ipt (PL 20mm S355JR) at ROOT>Column-1
COLLECT: PART Beam-17.ipt (UB254x146x31) at ROOT>Beam-17
...
COLLECT: Found 250 unique files

COPY: Starting file copy process...
REGISTRY: Loading existing counters for prefix: CLONE-001-
COPIED: Column-1.ipt -> CLONE-001-PL1.ipt
COPIED: Beam-17.ipt -> CLONE-001-B1.ipt
...

ASM UPDATE: Starting reference update process...
ASM UPDATE: Preloaded 200 parts into memory
ASM UPDATE: Opened 50 sub-assemblies
ASM UPDATE: Updated 250 references in 51 assemblies

IDW UPDATE: Starting IDW reference update process...
IDW UPDATE: Found 63 IDW files to process
IDW UPDATE: Updated 189 references in 63 IDW files

MAPPING: Wrote 250 entries to STEP_1_MAPPING.txt

VALIDATE: Validating clone and generating final report...
VALIDATE: SUCCESS - All copied files verified in destination
========================================
CLONE COMPLETED SUCCESSFULLY
========================================
```

---

## 🚀 Next Steps After Testing

### If Tests Pass:
1. ✅ Package for Autodesk App Store
2. ✅ Create installer
3. ✅ Write user documentation
4. ✅ Port remaining 29 tools

### If Tests Fail:
1. ❌ Debug using Visual Studio debugger
2. ❌ Compare with VBScript version
3. ❌ Check log files for errors
4. ❌ Verify each phase independently

---

## 📝 Notes

### Registry Permissions
If registry operations fail:
- Run Inventor as Administrator
- Or disable UAC temporarily
- Check HKEY_CURRENT_USER\Software\InventorRenamer\

### File Lock Issues
If file copy fails:
- Close Inventor before running
- Check files aren't open in other apps
- Check disk space

### Large Assemblies
For assemblies with 1000+ parts:
- Increase memory in Visual Studio
- Test with smaller assemblies first
- Check progress log for completion

---

**Last Updated:** 2025-01-21
**For questions, refer to HANDOVER.md or VBSCRIPT_TO_VBNET.md**
