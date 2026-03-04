# QUICK REFERENCE - Going Forward

**Last Updated:** January 9, 2026

---

## 🚀 INVENTOR ADD-IN - QUICK START

### First-Time Setup (One Time Only)
```
1. Run: InventorAddIn\DEPLOY_NOW.bat (as Administrator)
2. Close Inventor completely
3. Start Inventor
4. Go to: Tools → Add-Ins
5. Check: "Assembly Cloner with iLogic Patcher"
6. Check: "Load on Startup" (optional)
```

### After Code Changes (VB.NET Add-In)
```
1. Open: InventorAddIn\AssemblyClonerAddIn.sln
2. Set: Release | x64
3. Build → Build Solution
4. Run: InventorAddIn\DEPLOY_NOW.bat (as admin)
5. Restart Inventor
```

### Troubleshooting
```
Run: InventorAddIn\DIAGNOSE_ADDIN.bat
```

**Add-In Location:**
```
C:\ProgramData\Autodesk\Inventor 2026\Addins\
├── AssemblyClonerAddIn.dll (144 KB, x64)
└── AssemblyClonerAddIn.addin (668 bytes)
```

---

## ✅ BOM DECIMAL FIX - NOW WORKING!

### The Problem (SOLVED ✅)
- Scripts update document settings ✅
- BOM doesn't refresh ❌ → **FIXED!**

### The Solution
Added `ForceUnitsRefreshEvent()` function that toggles units (mm → cm → mm) to trigger BOM cache invalidation.

### Updated Files (3 Scripts)
1. ✅ `iLogic_Tools\Update_Decimal_Precision.vbs`
2. ✅ `iLogic_Tools\Force_BOM_Refresh.vbs`
3. ✅ `InventorAddIn\AssemblyClonerAddIn\PlateDocumentSettings.vb` (requires rebuild!)

### How to Test
```
1. Open assembly with plate parts
2. Run: iLogic_Tools\Launch_Decimal_Precision_Updater.bat
3. Check BOM - should show 0 decimals immediately
4. NO manual toggle needed!
```

**Details:** See `BOM_DECIMAL_FIX_SUMMARY.md`

---

## 📁 CLEAN FOLDER STRUCTURE

### Backup Cleanup Recommendations
| Folder | Size | Action | Savings |
|--------|------|--------|---------|
| **Backup 2/** | 158 MB | ❌ DELETE (nested duplicates) | 158 MB |
| **backup working/FINAL_PRODUCTION_SCRIPTS...** | 174 MB | ❌ DELETE (nested folder) | 174 MB |
| **Build artifacts (.vs, bin, obj)** | 15 MB | ❌ DELETE from backups | 15 MB |
| **Total Savings** | - | - | **~340 MB** |

**Keep:**
- ✅ `Backup/` - Dec 9 snapshot (clean out build artifacts)
- ✅ Main working directory

**Delete:**
- ❌ `Backup 2/` entirely
- ❌ `backup working/FINAL_PRODUCTION_SCRIPTS 1 Oct 2025/` nested folder
- ❌ All `.vs/`, `bin/Debug/`, `bin/Release/`, `obj/` in backups

---

## 📋 PRODUCTION WORKFLOW

### STEP 1: Part Renaming (Heritage Method)
```
Main_Launcher.bat → Option 1
```
- Renames all parts with client classification (PL, B, CH, A, FL, etc.)
- Global numbering across entire assembly hierarchy
- Creates STEP_1_MAPPING.txt (original → heritage names)
- Updates assembly references to new names

### STEP 2: IDW Updates
```
Main_Launcher.bat → Option 2
```
- Updates all IDW drawings to reference heritage part names
- Uses Design Assistant method (`ReplaceReference`)
- Processes entire assembly hierarchy dynamically
- No hardcoded file names or assumptions

### STEP 3: Title Automation
```
Main_Launcher.bat → Option 3
```
- Updates drawing view titles with professional formatting
- Base views: Part Number + Scale + Quantity
- Non-base views: View Name (bold)
- Uses dynamic parameters (auto-updates)

### Rescue Tools (If Needed)
- **Option 6:** Smart Prefix Scanner (prevent duplicates)
- **Option 7:** Emergency IDW Fixer (fix missed folders)

---

## 🔧 MAIN TOOLS OVERVIEW

| Tool | Purpose | When to Use |
|------|---------|-------------|
| **Part Renaming** | Heritage method, client classification | First step, new assemblies |
| **IDW Updates** | Update drawing references | After STEP 1 |
| **Title Automation** | Professional drawing titles | After STEP 2 |
| **Smart Prefix Scanner** | Detect existing numbering | Before adding new assemblies |
| **Emergency IDW Fixer** | Fix individual folders | When STEP 2 misses specific folders |
| **Registry Manager** | View numbering counters | Diagnostics only (read-only) |
| **Decimal Precision Updater** | Set plate BOM to 0 decimals | Anytime, works with fix! |
| **Export IDW Sheets to PDF** | Export each sheet as separate PDF | After drawing completion |

---

## 🎯 CRITICAL SUCCESS FACTORS

### ✅ DO:
- ✅ Use dynamic file scanning (no hardcoding)
- ✅ Process assembly-by-assembly (proven method)
- ✅ Close only IDW documents, NOT assemblies during iteration
- ✅ Run Smart Prefix Scanner BEFORE adding new assemblies
- ✅ Follow STEP 1 → STEP 2 → STEP 3 workflow
- ✅ Test on backup copies first
- ✅ Use ForceUnitsRefreshEvent() for BOM decimal fixes

### ❌ DON'T:
- ❌ Hardcode file names or paths
- ❌ Assume naming conventions (IDW ≠ assembly name)
- ❌ Close Structure.iam during sub-assembly iteration
- ❌ Skip validation steps in testing checklist
- ❌ Forget to backup before major operations
- ❌ Deploy add-in without rebuilding after code changes

---

## 📊 KEY FILE LOCATIONS

### Production Scripts
```
FINAL_PRODUCTION_SCRIPTS 1 Oct 2025/
├── Main_Launcher.bat                  # Main menu (15 options)
├── Part_Renaming/                     # STEP 1 scripts
│   ├── STEP_1_Heritage_Renaming.vbs
│   └── Smart_Prefix_Scanner.vbs
├── IDW_Updates/                       # STEP 2 scripts
│   ├── STEP_2_IDW_Updates.vbs
│   └── Emergency_IDW_Fixer.vbs
├── Title_Automation/                  # STEP 3 scripts
│   └── Title_Updater.vbs
├── iLogic_Tools/                      # BOM decimal fixes
│   ├── Update_Decimal_Precision.vbs   # ✅ FIXED
│   └── Force_BOM_Refresh.vbs          # ✅ FIXED
└── InventorAddIn/                     # Add-in development
    ├── DEPLOY_NOW.bat
    ├── DIAGNOSE_ADDIN.bat
    ├── QUICK_START_GUIDE.md
    └── AssemblyClonerAddIn/
        └── PlateDocumentSettings.vb   # ✅ FIXED (rebuild needed)
```

### Documentation
```
├── CLAUDE.md                          # Master memory/lessons learned
├── CRITICAL_LESSONS_LEARNED.md        # The two mistakes that broke everything
├── TESTING_CHECKLIST_NEW_PLANT.md     # Step-by-step test procedure
├── BOM_DECIMAL_FIX_SUMMARY.md         # BOM fix technical details
├── QUICK_REFERENCE_GOING_FORWARD.md   # This file
└── InventorAddIn/QUICK_START_GUIDE.md # Add-in deployment guide
```

---

## 🧪 TESTING ON NEW PLANT

**Follow:** `TESTING_CHECKLIST_NEW_PLANT.md`

**Golden Rules:**
1. ✅ **BACKUP FIRST** - Full project copy
2. ✅ **Follow checklist exactly** - All validation steps
3. ✅ **Document results** - Fill out ALL sections
4. ✅ **Verify assembly count** - STEP 2 should process 50+, not 1!
5. ✅ **Spot check 5 IDWs** - Random drawings for heritage references

---

## 💡 TROUBLESHOOTING QUICK GUIDE

### Add-In Not Showing in Inventor?
```
1. Run: InventorAddIn\DIAGNOSE_ADDIN.bat
2. Check: C:\ProgramData\Autodesk\Inventor 2026\Addins\ for files
3. Verify: .NET Framework 4.8 installed
4. Event Viewer: eventvwr.msc → Application → .NET errors
```

### BOM Still Showing Decimals?
```
1. Verify: Scripts updated with ForceUnitsRefreshEvent()
2. Run: Force_BOM_Refresh.vbs (includes fix)
3. Check: Document Settings → Units → Precision = 0
4. See: BOM_DECIMAL_FIX_SUMMARY.md
```

### STEP 2 Only Processing 1 Assembly?
```
Root cause: Closing Structure.iam during iteration
Fix: Only close IDW documents (type 12294), NOT assemblies
See: CRITICAL_LESSONS_LEARNED.md
```

### IDW References Not Updating?
```
Root cause: Full-path mapping required, not filename-only
Fix: Use Emergency_IDW_Fixer.vbs for specific folders
See: CLAUDE.md → Design Assistant Method
```

---

## 📞 QUICK HELP COMMANDS

| Issue | Command |
|-------|---------|
| Check add-in deployment | `InventorAddIn\DIAGNOSE_ADDIN.bat` |
| View numbering counters | `Registry_Management\Launch_Registry_Manager.bat` |
| Force BOM refresh | `iLogic_Tools\Launch_Force_BOM_Refresh.bat` |
| Fix missed IDW folder | `IDW_Updates\Launch_Emergency_IDW_Fixer.bat` |
| Scan existing numbers | `Part_Renaming\Launch_Smart_Prefix_Scanner.bat` |

---

## 🎉 SUCCESS CHECKLIST

**Project is ready when:**
- ✅ All parts renamed with heritage numbers
- ✅ All IDW drawings updated to reference heritage names
- ✅ All view titles professionally formatted
- ✅ BOM shows 0 decimals for plate quantities
- ✅ Smart Prefix Scanner run before adding new assemblies
- ✅ Testing checklist completed and documented
- ✅ Spot checks verify correctness

---

**REMEMBER:**
- 📖 Check CRITICAL_LESSONS_LEARNED.md when in doubt
- 🧪 Always test on backup copies first
- 📝 Document all issues in CLAUDE.md
- 🔧 Don't hardcode - use dynamic detection
- ✅ Follow the proven workflow: STEP 1 → STEP 2 → STEP 3

---

**Last Updated:** January 9, 2026
**Status:** Production ready with BOM decimal fix implemented
