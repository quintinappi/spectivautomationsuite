# 🔍 COMPLETE TOOL INDEX - Inventor Automation Suite
**Comprehensive Reference for All Tools, Scripts, and Utilities**

*Last Updated: January 23, 2026*

---

## 📖 HOW TO USE THIS INDEX

This document provides a complete catalog of all tools in the Inventor Automation Suite with:
- **Exact file paths** for quick location
- **Purpose and function** descriptions
- **Category organization** for logical grouping
- **Search keywords** for AI/human reference

**For AI assistants:** Use this index to quickly locate tools without scanning the entire workspace. Search for keywords like "unused", "IDW", "rename", "clone", etc.

---

## 🎯 CORE PRODUCTION WORKFLOW

### Assembly_Renamer.vbs
**Location:** `Part_Renaming\Assembly_Renamer.vbs`  
**Launcher:** `Part_Renaming\Launch_Assembly_Renamer.bat`  
**Purpose:** **STEP 1** - Heritage renaming of parts in existing assembly with group-based sequential numbering  
**Keywords:** rename, heritage, grouping, CH, PL, B, A, P, SQ, FL, IPE, FLG, R, PART, mapping, STEP_1_MAPPING.txt  
**Registry Keys:** `HKCU\Software\InventorRenamer\[prefix][groupcode]`  
**Input:** Existing assembly file  
**Output:** Renamed parts + STEP_1_MAPPING.txt  
**Notes:** Creates mapping file automatically. Fixed to update Part Number iProperty. IPE parts go to B grouping.

### IDW_Reference_Updater.vbs
**Location:** `IDW_Updates\IDW_Reference_Updater.vbs`  
**Launcher:** `IDW_Updates\Launch_IDW_Reference_Updater.bat`  
**Purpose:** **STEP 2** - Update ALL IDW drawing references using mapping file  
**Keywords:** IDW, drawing, reference, update, mapping, STEP_1_MAPPING.txt, silent operation  
**Input:** STEP_1_MAPPING.txt  
**Output:** Updated IDW files  
**Notes:** Uses SilentOperation mode. Searches recursively for IDW files. Handles IDW→IPT references.

---

## 🔄 CLONING TOOLS

### Assembly_Cloner.vbs
**Location:** `Assembly_Cloner.vbs` (root) OR `Part_Renaming\Assembly_Cloner.vbs`  
**Launcher:** `Launch_Assembly_Cloner.bat` (root) OR `Part_Renaming\Launch_Assembly_Cloner.bat`  
**Purpose:** Clone assembly with optional heritage renaming (group-based numbering)  
**Keywords:** clone, duplicate, copy, heritage, renaming, grouping, iProperties, Part Number, IDW, recursive scan  
**Registry Keys:** `HKCU\Software\InventorRenamer\[prefix][groupcode]`  
**Input:** Source assembly + new prefix  
**Output:** Cloned assembly with renamed parts  
**Notes:** Fixed to scan IDW files recursively using ApprenticeServer. Updates Part Number iProperty. IPE→B grouping.

### Prefix_Cloner.vbs
**Location:** `Part_Renaming\Prefix_Cloner.vbs`  
**Launcher:** `Part_Renaming\Launch_Prefix_Cloner.bat`  
**Purpose:** Clone assembly with prefix replacement ONLY (keeps original suffixes)  
**Keywords:** clone, prefix, replace, suffix, maintain, iProperties, Part Number  
**Input:** Source assembly + new prefix  
**Output:** Cloned assembly with prefix-replaced parts  
**Notes:** Does NOT do heritage renaming. Only replaces prefix. Updates Part Number iProperty.

### Part_Cloner.vbs
**Location:** `Part_Renaming\Part_Cloner.vbs`  
**Launcher:** `Part_Renaming\Launch_Part_Cloner.bat`  
**Purpose:** Clone individual parts (not assemblies)  
**Keywords:** clone, part, IPT, single file  
**Input:** Individual IPT file  
**Output:** Cloned part file  

---

## 🗑️ FILE UTILITIES

### Unused_Part_Finder.vbs
**Location:** `File_Utilities\Unused_Part_Finder.vbs`  
**Launcher:** `File_Utilities\Launch_Unused_Part_Finder.bat`  
**Purpose:** Find and move unused IPT files to backup folder  
**Keywords:** unused, orphan, cleanup, backup, move, IPT, assembly scan, file management  
**Input:** Assembly file path  
**Output:** Unused_Parts_Backup_[datetime] folder with moved files + Backup_Summary.txt  
**Log:** `Logs\Unused_Part_Finder_[datetime].log`  
**Notes:** Skips folders starting with "unused_parts_backup" and "oldversions". Scans assemblies recursively.

### Duplicate_File_Finder.vbs
**Location:** `File_Utilities\Duplicate_File_Finder.vbs`  
**Launcher:** `File_Utilities\Launch_Duplicate_File_Finder.bat`  
**Purpose:** Find duplicate files and naming conflicts  
**Keywords:** duplicate, conflict, find, scan, file management  
**Input:** Folder path  
**Output:** Report of duplicate files  

### IDW_Parts_List_Scanner.vbs
**Location:** `File_Utilities\IDW_Parts_List_Scanner.vbs`  
**Launcher:** `File_Utilities\Launch_IDW_Parts_List_Scanner.bat`  
**Purpose:** Scan IDW files for parts list data  
**Keywords:** IDW, parts list, BOM, scan, analyze  
**Input:** IDW folder path  
**Output:** Parts list report  
**Log:** `File_Utilities\IDW_Parts_List_Scanner_[datetime].log`  

---

## 📐 IDW UPDATES & RESCUE TOOLS

### Emergency_IDW_Fixer.vbs
**Location:** `IDW_Updates\Emergency_IDW_Fixer.vbs`  
**Launcher:** `IDW_Updates\Launch_Emergency_IDW_Fixer.bat`  
**Purpose:** **TROUBLESHOOTING** - Fix specific IDW folders missed by standard updater  
**Keywords:** emergency, rescue, IDW, fix, troubleshoot, folder-specific, silent operation  
**Input:** Specific folder path + mapping file  
**Output:** Updated IDW files  
**Notes:** Uses SilentOperation mode. For problematic folders that standard updater missed.

### IDW_Assembly_Synchronizer.vbs
**Location:** `IDW_Updates\IDW_Assembly_Synchronizer.vbs`  
**Launcher:** `IDW_Updates\Launch_IDW_Assembly_Synchronizer.bat`  
**Purpose:** **TROUBLESHOOTING** - Sync scattered IDWs when they reference renamed assemblies  
**Keywords:** IDW, assembly, sync, scattered, reference update  
**Input:** IDW files + renamed assembly  
**Output:** Synchronized IDW files  

### Recursive_IDW_Updater.vbs
**Location:** `IDW_Updates\Recursive_IDW_Updater.vbs`  
**Launcher:** `IDW_Updates\Launch_Recursive_IDW_Updater.bat`  
**Purpose:** Recursively update IDW files in nested folders  
**Keywords:** IDW, recursive, nested, update  
**Input:** Root folder path  
**Output:** Updated IDW files  

---

## 🎨 IDW UTILITIES

### IDW_Part_Placer.vbs
**Location:** `IDW_Utilities\IDW_Part_Placer.vbs`  
**Launcher:** `IDW_Utilities\Launch_IDW_Part_Placer.bat`  
**Purpose:** Place parts on IDW sheets  
**Keywords:** IDW, place, parts, sheet, layout  
**Log:** `IDW_Utilities\IDW_Part_Placer.log`  

### Copy_Views_Sheet1_to_Sheet2.vbs
**Location:** `IDW_Utilities\Copy_Views_Sheet1_to_Sheet2.vbs`  
**Purpose:** Copy views between IDW sheets  
**Keywords:** IDW, copy, views, sheets, transfer  

---

## 🏷️ TITLE AUTOMATION

### Title_Updater.vbs
**Location:** `Title_Automation\Title_Updater.vbs`  
**Launcher:** `Title_Automation\Launch_Title_Updater.bat`  
**Purpose:** Update IDW drawing title blocks  
**Keywords:** title, title block, IDW, update, automation  
**Output:** `Title_Automation\title_output.txt`  

### Auto_Balloon_Views.vbs
**Location:** `Title_Automation\Auto_Balloon_Views.vbs`  
**Launcher:** `Title_Automation\Launch_Auto_Balloon_Views.bat`  
**Purpose:** Automatically add balloons to IDW views  
**Keywords:** balloon, auto, IDW, annotation, views  
**Log:** `Title_Automation\AutoBalloonLog.txt`  
**Documentation:** `Title_Automation\Auto_Balloon_Views_README.md`  

### Change_Balloon_Style.vbs
**Location:** `Title_Automation\Change_Balloon_Style.vbs`  
**Launcher:** `Title_Automation\Launch_Change_Balloon_Style.bat`  
**Purpose:** Change balloon style across IDW drawings  
**Keywords:** balloon, style, change, IDW, format  
**Log:** `Title_Automation\ChangeBalloonStyleLog.txt`  

### Change_Dimension_Style.vbs
**Location:** `Title_Automation\Change_Dimension_Style.vbs`  
**Launcher:** `Title_Automation\Launch_Change_Dimension_Style.bat`  
**Purpose:** Change dimension style across IDW drawings  
**Keywords:** dimension, style, change, IDW, format  

### Quantity_Detector.vbs
**Location:** `Title_Automation\Quantity_Detector.vbs`  
**Purpose:** Detect quantities in parts lists  
**Keywords:** quantity, detect, parts list, count  

### View_Summary_Generator.vbs
**Location:** `Title_Automation\View_Summary_Generator.vbs`  
**Purpose:** Generate summary of views in IDW files  
**Keywords:** view, summary, report, IDW, analysis  

---

## 🎯 VIEW STYLE MANAGER

### View_Style_Manager.vbs
**Location:** `View_Style_Manager\View_Style_Manager.vbs`  
**Launcher:** `View_Style_Manager\Launch_View_Style_Manager.bat`  
**Purpose:** Manage and apply view styles across IDW drawings  
**Keywords:** view style, standard, apply, manage, IDW  
**Log:** `View_Style_Manager\ViewStyleManager_[datetime].log`  
**Documentation:** `View_Style_Manager\README.md`, `View_Style_Manager\QUICK_START.md`  

### Apply_Standard_To_Views.vbs
**Location:** `View_Style_Manager\Apply_Standard_To_Views.vbs`  
**Purpose:** Apply standard view styles to views  
**Keywords:** standard, view style, apply  

### Change_Document_Standard.vbs
**Location:** `View_Style_Manager\Change_Document_Standard.vbs`  
**Purpose:** Change document standard settings  
**Keywords:** document, standard, change, settings  

### Force_View_Style_Update.vbs
**Location:** `View_Style_Manager\Force_View_Style_Update.vbs`  
**Purpose:** Force update of view styles  
**Keywords:** force, update, view style, refresh  

### View_Style_Diagnostic.vbs
**Location:** `View_Style_Manager\View_Style_Diagnostic.vbs`  
**Purpose:** Diagnose view style issues  
**Keywords:** diagnostic, troubleshoot, view style, analyze  

### Deep_View_Analysis.vbs
**Location:** `View_Style_Manager\Deep_View_Analysis.vbs`  
**Purpose:** Deep analysis of view properties  
**Keywords:** analysis, deep, view, properties, inspect  

### Test_View_Style_Properties.vbs
**Location:** `View_Style_Manager\Test_View_Style_Properties.vbs`  
**Purpose:** Test view style property changes  
**Keywords:** test, view style, properties  

---

## 🗄️ REGISTRY MANAGEMENT

### Registry_Manager.vbs
**Location:** `Registry_Management\Registry_Manager.vbs`  
**Launcher:** `Registry_Management\Launch_Registry_Manager.bat`  
**Purpose:** Manage numbering counters in Windows Registry  
**Keywords:** registry, counter, numbering, manage, scan, update, clear  
**Registry Keys:** `HKCU\Software\InventorRenamer\[prefix][groupcode]`  
**Functions:** Scan, Update, Clear registry values  

### Clear_Renamer_Registry.vbs
**Location:** `Registry_Management\Clear_Renamer_Registry.vbs`  
**Purpose:** Clear all renaming registry counters (utility script)  
**Keywords:** clear, reset, registry, counter, clean  
**Registry Keys:** `HKCU\Software\InventorRenamer\`  
**Notes:** Utility for resetting all group counters to start fresh  

---

## 🔧 DERIVED PARTS & CROSS-REFERENCE TOOLS

### Fix_Derived_Parts.vbs
**Location:** `Part_Renaming\Fix_Derived_Parts.vbs`  
**Launcher:** `Part_Renaming\Launch_Fix_Derived_Parts.bat`  
**Purpose:** Fix derived part references after renaming  
**Keywords:** derived, parts, fix, reference, update  
**Log:** `Part_Renaming\Fix_Derived_Log.txt`  
**Documentation:** `Part_Renaming\DERIVED_PARTS_HANDLING.md`  

### Update_SameFolder_Derived_Parts.vbs
**Location:** `Part_Renaming\Update_SameFolder_Derived_Parts.vbs`  
**Launcher:** `Part_Renaming\Launch_Update_SameFolder_Derived_Parts.bat`  
**Purpose:** **STEP 2 (OPTIONAL)** - Fix derived parts in same folder  
**Keywords:** derived, same folder, update, local  
**Log:** `Part_Renaming\Update_Derived_Log.txt`  

### Check_Derived_Parts.vbs
**Location:** `Part_Renaming\Check_Derived_Parts.vbs`  
**Purpose:** Check for derived parts in assembly  
**Keywords:** check, derived, scan, detect  
**Report:** `Part_Renaming\DerivedParts_Report.txt`  

### Fix_Structure_CrossRefs.vbs
**Location:** `Part_Renaming\Fix_Structure_CrossRefs.vbs`  
**Purpose:** Fix cross-references in structure  
**Keywords:** cross-reference, structure, fix  
**Log:** `Part_Renaming\Structure_CrossRef_Fix_Log.txt`  

### Scan_Structure_CrossRefs.vbs
**Location:** `Part_Renaming\Scan_Structure_CrossRefs.vbs`  
**Purpose:** Scan for cross-references in structure  
**Keywords:** scan, cross-reference, structure, detect  
**Report:** `Part_Renaming\Structure_CrossRef_Report.txt`  

### Scan_IDW_CrossRefs.vbs
**Location:** `Part_Renaming\Scan_IDW_CrossRefs.vbs`  
**Purpose:** Scan for cross-references in IDW files  
**Keywords:** scan, cross-reference, IDW, detect  
**Report:** `Part_Renaming\CrossRef_Report.txt`  

---

## 🧩 iLOGIC & DIAGNOSTIC TOOLS

### iLogic_Patcher.vbs
**Location:** `Part_Renaming\iLogic_Patcher.vbs`  
**Launcher:** `Part_Renaming\Launch_iLogic_Patcher.bat`  
**Purpose:** **STEP 3 (OPTIONAL)** - Fix iLogic rules after renaming  
**Keywords:** iLogic, patch, fix, rules, update  

### Diagnose_AttributeSets.vbs
**Location:** `Part_Renaming\Diagnose_AttributeSets.vbs`  
**Purpose:** Diagnose attribute set issues  
**Keywords:** diagnose, attribute, attribute set, inspect  

### Diagnose_Flatbars.vbs
**Location:** `Part_Renaming\Diagnose_Flatbars.vbs`  
**Launcher:** `Part_Renaming\Launch_Flatbar_Diagnostic.bat`  
**Purpose:** Diagnose flatbar part issues  
**Keywords:** diagnose, flatbar, inspect, troubleshoot  

---

## 📊 MAPPING FILE TOOLS

### Generate_Mapping.vbs
**Location:** `Part_Renaming\Generate_Mapping.vbs`  
**Purpose:** Generate mapping file for renaming operations  
**Keywords:** mapping, generate, create, STEP_1_MAPPING.txt  
**Output:** Mapping file  

### Create_Filename_Mapping.vbs
**Location:** `Part_Renaming\Create_Filename_Mapping.vbs`  
**Purpose:** Create filename-based mapping  
**Keywords:** filename, mapping, create  

### Create_Proper_Mapping.vbs
**Location:** `Part_Renaming\Create_Proper_Mapping.vbs`  
**Purpose:** Create properly formatted mapping file  
**Keywords:** proper, mapping, create, format  

### Create_Complete_Mapping.vbs
**Location:** `Part_Renaming\Create_Complete_Mapping.vbs`  
**Purpose:** Create complete mapping with all references  
**Keywords:** complete, mapping, create, comprehensive  

### Protect_Mapping_File.vbs
**Location:** `Part_Renaming\Protect_Mapping_File.vbs`  
**Launcher:** `Part_Renaming\Launch_Protect_Mapping_File.bat`  
**Purpose:** Protect mapping file from accidental edits  
**Keywords:** protect, mapping, lock, read-only  

### Recover_Mapping.vbs
**Location:** `Part_Renaming\Recover_Mapping.vbs`  
**Launcher:** `Part_Renaming\Launch_Recover_Mapping.bat`  
**Purpose:** Recover lost or corrupted mapping file  
**Keywords:** recover, mapping, restore, backup  

---

## 🔍 SCANNING & ANALYSIS TOOLS

### Smart_Prefix_Scanner.vbs
**Location:** `Part_Renaming\Smart_Prefix_Scanner.vbs`  
**Launcher:** `Part_Renaming\Launch_Smart_Prefix_Scanner.bat`  
**Purpose:** **BEFORE RENAMING** - Scan model to detect duplicate part numbers  
**Keywords:** scan, prefix, duplicate, detect, prevent, pre-check  

### Scan_Assembly_Refs.vbs
**Location:** `Scan_Assembly_Refs.vbs` (root)  
**Purpose:** Scan assembly references  
**Keywords:** scan, assembly, references, analyze  

---

## 🔄 ROOT-LEVEL UTILITIES

### Fix_Mapping_Paths.vbs
**Location:** `Fix_Mapping_Paths.vbs` (root)  
**Purpose:** Fix paths in mapping files  
**Keywords:** fix, mapping, paths, correct  

### Replace_Mapping_File.vbs
**Location:** `Replace_Mapping_File.vbs` (root)  
**Purpose:** Replace mapping file  
**Keywords:** replace, mapping, swap  

### Part_Grouping_Correction.vbs
**Location:** `Part_Grouping_Correction.vbs` (root)  
**Launcher:** `Launch_Part_Grouping_Correction.bat` (root)  
**Purpose:** Correct part grouping assignments  
**Keywords:** grouping, correction, fix, reassign  
**Log:** `Grouping_Correction_Log.txt`  

---

## 🧹 CLEANUP & MAINTENANCE TOOLS

### Nuke_BackupWorking.ps1
**Location:** `Nuke_BackupWorking.ps1` (root)  
**Purpose:** Delete backup working folders  
**Keywords:** nuke, delete, backup, clean, cleanup, PowerShell  

### Nuke_Recursive_Folder.ps1
**Location:** `Nuke_Recursive_Folder.ps1` (root)  
**Purpose:** Recursively delete folders  
**Keywords:** nuke, delete, recursive, clean, PowerShell  

### Nuke_Recursive_Folder_AUTO.ps1
**Location:** `Nuke_Recursive_Folder_AUTO.ps1` (root)  
**Purpose:** Automatically delete recursive folders (no prompts)  
**Keywords:** nuke, delete, recursive, auto, automatic, PowerShell  

### Nuke_with_LongPath.ps1
**Location:** `Nuke_with_LongPath.ps1` (root)  
**Purpose:** Delete folders with long path names  
**Keywords:** nuke, delete, long path, Windows path limit, PowerShell  

---

## 🧪 TEST & EXPERIMENTAL TOOLS

### temp_test.vbs
**Location:** `temp_test.vbs` (root)  
**Purpose:** Temporary testing script  
**Keywords:** test, temporary, experimental  

### test_apprentice.vbs
**Location:** `test_apprentice.vbs` (root)  
**Purpose:** Test ApprenticeServer functionality  
**Keywords:** test, apprentice, ApprenticeServer  

### Test_Single_IDW_Refs.vbs
**Location:** `Part_Renaming\Test_Single_IDW_Refs.vbs`  
**Purpose:** Test single IDW reference updates  
**Keywords:** test, IDW, reference, single file  

---

## 🚀 LAUNCHERS & UI

### Launch_UI.ps1
**Location:** `Launch_UI.ps1` (root)  
**Purpose:** Launch PowerShell-based UI for tool selection  
**Keywords:** launcher, UI, menu, PowerShell, main  

### SpectivLauncher.exe
**Location:** `SpectivLauncher.exe` (root)  
**Purpose:** Compiled launcher executable  
**Keywords:** launcher, executable, EXE, compiled  

### Inventor_Automation_Suite.hta
**Location:** `Inventor_Automation_Suite.hta` (root)  
**Purpose:** HTA-based launcher (legacy - not store-compliant)  
**Keywords:** launcher, HTA, legacy  

### Build_Launcher.bat
**Location:** `Build_Launcher.bat` (root)  
**Purpose:** Build/compile the launcher executable  
**Keywords:** build, compile, launcher  

---

## 📦 ADD-IN INTEGRATION

### InventorAutomationSuiteAddIn (Folder)
**Location:** `InventorAutomationSuiteAddIn\`  
**Purpose:** VB.NET Inventor Add-In project for Autodesk App Store  
**Keywords:** add-in, VB.NET, Inventor, ribbon, commercial  
**Contents:** MainForm.vb, AssemblyClonerForm.vb, AssemblyCloner.vb, etc.  
**Documentation:** `INVENTOR_ADDIN_INTEGRATION_PLAN.md`, `BUILD_INSTRUCTIONS.md`  

### InventorAutomationSuiteAddIn.xml
**Location:** `InventorAutomationSuiteAddIn.xml` (root)  
**Purpose:** Add-in manifest for Inventor registration  
**Keywords:** manifest, XML, add-in, registration  

---

## 📚 DOCUMENTATION FILES

### PROJECT_SUMMARY.md
**Location:** `PROJECT_SUMMARY.md` (root)  
**Purpose:** Complete project status and accomplishments  

### QUICKSTART.md
**Location:** `QUICKSTART.md` (root)  
**Purpose:** Quick start guide for new users  

### README_SUITE.md
**Location:** `README_SUITE.md` (root)  
**Purpose:** Suite overview and usage guide  

### LAUNCHER_SCRIPTS_LIST.md
**Location:** `LAUNCHER_SCRIPTS_LIST.md` (root)  
**Purpose:** Complete list of launcher scripts organized by category  

### BUILD_INSTRUCTIONS.md
**Location:** `BUILD_INSTRUCTIONS.md` (root)  
**Purpose:** Instructions for building the add-in  

### VISUAL_STUDIO_SETUP_GUIDE.md
**Location:** `VISUAL_STUDIO_SETUP_GUIDE.md` (root)  
**Purpose:** Visual Studio setup for add-in development  

### INSTALLER_CREATION_GUIDE.md
**Location:** `INSTALLER_CREATION_GUIDE.md` (root)  
**Purpose:** Guide for creating MSI installer  

### AUTODESK_STORE_COMPLIANCE.md
**Location:** `AUTODESK_STORE_COMPLIANCE.md` (root)  
**Purpose:** Autodesk App Store compliance requirements  

### COMMERCIALIZATION_ROADMAP.md
**Location:** `COMMERCIALIZATION_ROADMAP.md` (root)  
**Purpose:** Roadmap for commercial release  

### AUTODESK_APP_STORE_COMMERCIALIZATION_PLAN.md
**Location:** `AUTODESK_APP_STORE_COMMERCIALIZATION_PLAN.md` (root)  
**Purpose:** Complete commercialization plan  

### CLEANUP_SUMMARY.md
**Location:** `CLEANUP_SUMMARY.md` (root)  
**Purpose:** Summary of cleanup operations performed  

### RENAME_WORKFLOW_GUIDE.txt
**Location:** `RENAME_WORKFLOW_GUIDE.txt` (root)  
**Purpose:** Step-by-step renaming workflow guide  

---

## 📂 FOLDER STRUCTURE REFERENCE

```
FINAL_PRODUCTION_SCRIPTS 1 Oct 2025/
├── Part_Renaming/           # Main renaming tools and cloners
├── IDW_Updates/             # IDW reference update and rescue tools
├── IDW_Utilities/           # IDW manipulation utilities
├── File_Utilities/          # File management (unused finder, duplicates)
├── Title_Automation/        # Title block and balloon automation
├── View_Style_Manager/      # View style management and diagnostics
├── Registry_Management/     # Registry counter management
├── InventorAutomationSuiteAddIn/  # VB.NET Add-In project
├── InventorAddIn/           # Legacy add-in
├── InventorAddIn_CS/        # C# add-in experiments
├── iLogic_Tools/            # iLogic-related tools
├── Documentation/           # Additional documentation
├── Archives/                # Archived/backup files
├── Experiments/             # Experimental scripts
├── Logs/                    # Log files from tool execution
├── Todos/                   # Task lists and notes
└── APPDATA/                 # Application data
```

---

## 🔑 KEY CONCEPTS & KEYWORDS

### Group Codes (Heritage Renaming)
**CH** - Channel  
**PL** - Plate  
**B** - Beam (includes IPE beams)  
**A** - Angle  
**P** - Pipe  
**SQ** - Square/Rectangular Tube  
**FL** - Flat  
**LPL** - Large Plate  
**FLG** - Flange  
**R** - Round Bar  
**OTHER** - Unclassified parts  
**PART** - Generic parts  

### Registry Structure
- **Base Path:** `HKCU\Software\InventorRenamer\`
- **Key Format:** `[prefix][groupcode]` (e.g., `N1SCR05-001-CH`, `N1SCR05-001-PL`)
- **Value:** Highest used number for that group
- **Management:** Use Registry_Manager.vbs or Clear_Renamer_Registry.vbs

### Mapping Files
- **STEP_1_MAPPING.txt** - Created by Assembly_Renamer, used by IDW_Reference_Updater
- **Format:** `OldFilename.ipt,NewFilename.ipt` (one per line)
- **Location:** Same folder as assembly being renamed
- **Protection:** Can be locked with Protect_Mapping_File.vbs

### Log Files
- **Location:** `Logs\[ToolName]_[datetime].log`
- **Format:** Timestamped entries with operation details
- **Retention:** Kept for troubleshooting and audit trail

---

## 🆘 COMMON SEARCH SCENARIOS

**Finding the unused file cleaner:**
→ Search for "unused" → Find `Unused_Part_Finder.vbs` in `File_Utilities\`

**Finding IDW update tools:**
→ Search for "IDW" → Multiple tools in `IDW_Updates\` and `IDW_Utilities\`

**Finding registry management:**
→ Search for "registry" or "counter" → Find tools in `Registry_Management\`

**Finding cloning tools:**
→ Search for "clone" → `Assembly_Cloner.vbs`, `Prefix_Cloner.vbs`, `Part_Cloner.vbs` in root and `Part_Renaming\`

**Finding renaming tools:**
→ Search for "rename" or "heritage" → `Assembly_Renamer.vbs` in `Part_Renaming\`

**Finding troubleshooting/rescue tools:**
→ Search for "emergency", "fix", "diagnose" → Various diagnostic and fix scripts

**Finding derived parts tools:**
→ Search for "derived" → Tools in `Part_Renaming\` for derived part handling

**Finding mapping file tools:**
→ Search for "mapping" → Tools in `Part_Renaming\` for mapping file operations

---

## 📞 RELATED DOCUMENTATION

For detailed workflow instructions, see:
- **QUICKSTART.md** - Getting started guide
- **LAUNCHER_SCRIPTS_LIST.md** - Organized tool list by category
- **RENAME_WORKFLOW_GUIDE.txt** - Step-by-step renaming process
- **View_Style_Manager\README.md** - View style management details
- **Part_Renaming\DERIVED_PARTS_HANDLING.md** - Derived parts workflow

---

**End of Index** - Use Ctrl+F to search this document for specific tools or keywords
