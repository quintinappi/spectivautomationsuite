# Inventor Add-In - Quick Start Guide

**Assembly Cloner with iLogic Patcher**
*Last Updated: January 8, 2026*

---

## 🚀 FIRST TIME SETUP (One-Time Only)

### Step 1: Deploy the Add-In

**Option A: Automatic Deployment (Recommended)**
```
1. Navigate to: FINAL_PRODUCTION_SCRIPTS\InventorAddIn\
2. Right-click DEPLOY_NOW.bat
3. Select "Run as administrator"
4. Wait for "DEPLOYMENT SUCCESSFUL!" message
```

**Option B: Manual Deployment**
```
Copy these 2 files to: C:\ProgramData\Autodesk\Inventor 2026\Addins\
- AssemblyClonerAddIn\bin\x64\Release\AssemblyClonerAddIn.dll
- AssemblyClonerAddIn\AssemblyClonerAddIn.addin
```

### Step 2: Load in Inventor

```
1. CLOSE Inventor completely (check Task Manager - no Inventor.exe)
2. START Inventor
3. Go to: Tools → Add-Ins
4. Find: "Assembly Cloner with iLogic Patcher"
5. Check ✓ "Loaded"
6. Check ✓ "Load on Startup" (optional, but recommended)
```

**Done!** The add-in is now active.

---

## 📋 WHAT THE ADD-IN DOES

**Automatic Features (Always Active):**
- ✅ **Plate Part Monitor** - Automatically removes decimals from BOM quantities when you save plate parts
- ✅ **iLogic Integration** - Preserves iLogic rules when cloning assemblies

**Manual Tools (Use from Ribbon):**
- 🔧 **Clone Assembly** - Copy assemblies with automatic reference updates
- 🔍 **Document Info Scanner** - Analyze assembly structure
- ✏️ **Part Renamer** - Heritage-based batch renaming
- 📊 **Smart Inspector** - Assembly validation and checking
- 🏗️ **Beam Generator** - Create steel beam assemblies
- ⚙️ **Plate Settings** - Configure plate document settings

---

## 🎯 DAILY USAGE

### After Inventor Starts:
**The add-in loads automatically** - you'll see the tools in the Inventor ribbon under the "Add-Ins" tab.

### When Working with Plate Parts:
1. Open your plate part (.ipt file)
2. Make your changes
3. **Save the file** (Ctrl+S)
4. ✅ **BOM decimals automatically fixed to 0** on save

### When Cloning Assemblies:
1. Open the assembly you want to clone
2. Go to: Add-Ins tab → Clone Assembly
3. Select destination folder
4. ✅ **Parts copied, references updated, iLogic rules patched**

---

## ⚠️ TROUBLESHOOTING

### Add-In Not Showing in Inventor?

**Run Diagnostic:**
```
InventorAddIn\DIAGNOSE_ADDIN.bat
```

**Common Fixes:**
1. ✅ Verify deployment: `C:\ProgramData\Autodesk\Inventor 2026\Addins\AssemblyClonerAddIn.dll` exists
2. ✅ Check .NET Framework: Must be 4.8 or higher
3. ✅ Verify Inventor version: Must be 2026
4. ✅ Check Event Viewer: `eventvwr.msc` → Windows Logs → Application (look for .NET errors)

### BOM Decimals Not Updating?

**Known Issue:** BOM quantities don't refresh until manual document settings toggle
- 🔧 **Fix in progress** - investigating Inventor API refresh methods
- **Workaround:** Open document settings, change units → change back (no save needed)

### DLL Missing or Wrong Architecture?

**Verify Build:**
```
1. The DLL MUST be from: bin\x64\Release\ (144 KB)
2. NOT from: bin\Debug\ (26 KB) ❌
3. NOT from: bin\Release\ (37 KB, 32-bit) ❌
```

**Rebuild if Needed:**
```
1. Open: InventorAddIn\AssemblyClonerAddIn.sln in Visual Studio
2. Set configuration: Release | x64
3. Build → Build Solution
4. Re-run DEPLOY_NOW.bat
```

---

## 🔄 UPDATING THE ADD-IN

**After Making Code Changes:**

```
1. Open AssemblyClonerAddIn.sln in Visual Studio
2. Make your changes to .vb files
3. Set configuration: Release | x64
4. Build → Build Solution
5. Run: DEPLOY_NOW.bat (as administrator)
6. CLOSE Inventor completely
7. RESTART Inventor
```

**The add-in will reload with your changes.**

---

## 📁 FOLDER STRUCTURE

```
InventorAddIn/
├── DEPLOY_NOW.bat              ← Deploy script (run as admin)
├── DIAGNOSE_ADDIN.bat          ← Diagnostic tool
├── QUICK_START_GUIDE.md        ← This file
├── README.md                   ← Technical documentation
├── AssemblyClonerAddIn.sln     ← Visual Studio solution
└── AssemblyClonerAddIn/
    ├── StandardAddInServer.vb   ← Entry point (creates ribbon)
    ├── AssemblyCloner.vb        ← Cloning logic
    ├── PartRenamer.vb           ← Heritage renaming
    ├── iLogicPatcher.vb         ← iLogic rule updates
    ├── PlateDocumentSettings.vb ← BOM decimal fixer
    ├── [other modules...]
    ├── AssemblyClonerAddIn.addin ← Manifest (XML config)
    └── bin/x64/Release/
        └── AssemblyClonerAddIn.dll ← Compiled add-in (144 KB)
```

---

## 🛠️ DEVELOPMENT NOTES

### Key Files to Modify:

**Add-In Activation:**
- `StandardAddInServer.vb` - Entry point, creates ribbon buttons

**Core Features:**
- `AssemblyCloner.vb` - Assembly copying logic
- `PartRenamer.vb` - Heritage-based renaming with client classification
- `iLogicPatcher.vb` - iLogic rule text replacement
- `PlateDocumentSettings.vb` - **BOM decimal fixing** (currently under investigation)

**UI Forms:**
- `BeamGeneratorForm.vb` - Beam creation UI
- `AssemblyInspectorForm.vb` - Inspector UI

**Data:**
- `SteelSectionData.vb` - Steel section database

### GUID (Don't Change!)
```
ClassId: {B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}
ClientId: {B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}
```

If you change the GUID, Inventor will treat it as a NEW add-in.

---

## 📞 SUPPORT

**Diagnostic Tool:**
```
InventorAddIn\DIAGNOSE_ADDIN.bat
```

**Event Viewer (for .NET errors):**
```
Win+R → eventvwr.msc
Windows Logs → Application
Filter: .NET Runtime, Inventor
```

**Registry Check:**
```
HKEY_CURRENT_USER\Software\Autodesk\Inventor\Addins\{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}
LoadBehavior = 1 (auto-created by Inventor after first load)
```

---

## ✅ CHECKLIST

**First-Time Setup:**
- [ ] Run DEPLOY_NOW.bat (as admin)
- [ ] Close Inventor completely
- [ ] Restart Inventor
- [ ] Tools → Add-Ins → Check "Loaded"

**After Code Changes:**
- [ ] Build: Release | x64
- [ ] Run DEPLOY_NOW.bat
- [ ] Restart Inventor

**Verify Working:**
- [ ] Add-in appears in Tools → Add-Ins
- [ ] Ribbon buttons visible in Add-Ins tab
- [ ] Save a plate part → BOM decimals = 0

---

**Assembly Cloner with iLogic Patcher**
*By Quintin de Bruin © 2025*
