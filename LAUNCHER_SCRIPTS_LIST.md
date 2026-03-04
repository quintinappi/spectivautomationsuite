# Spectiv Inventor Automation Suite - Complete Scripts List

## Main Launcher: `Launch_UI.ps1` (or SpectivLauncher.exe)

Last Updated: January 20, 2026

---

## 📋 ORGANIZATION BY CATEGORY

### 1️⃣ **CORE PRODUCTION WORKFLOW** (Blue)
*Main renaming and IDW update workflow*

| # | Script Name | Function | Description |
|---|-------------|----------|-------------|
| 1 | **Assembly Renamer** | `Part_Renaming\Launch_Assembly_Renamer.bat` | **STEP 1**: Rename parts with heritage method. Creates STEP_1_MAPPING.txt automatically. **RUN FIRST** |
| 2 | **Update Same-Folder Derived Parts** | `Part_Renaming\Launch_Update_SameFolder_Derived_Parts.bat` | **STEP 2**: Fix local derived parts. ONLY if you have derived parts in same folder. **OPTIONAL** |
| 3 | **iLogic Patcher** | `Part_Renaming\Launch_iLogic_Patcher.bat` | **STEP 3**: Fix iLogic rules. ONLY if assembly has iLogic rules. **OPTIONAL** |
| 4 | **IDW Updates** | `IDW_Updates\Launch_IDW_Reference_Updater.bat` | **STEP 4**: Update ALL IDW drawings to new part names using mapping file. **ALWAYS REQUIRED** |
| 5 | **Title Automation** | `Title_Automation\Launch_Title_Updater.bat` | Update IDW drawing titles. Run after IDW updates complete. **OPTIONAL** |

---

### 2️⃣ **MANAGEMENT & UTILITIES** (Dark Blue)
*Database and file management tools*

| # | Script Name | Function | Description |
|---|-------------|----------|-------------|
| 6 | **Registry Management** | `Registry_Management\Launch_Registry_Manager.bat` | Manage numbering counters and database (scan, update, clear) |
| 7 | **File Utilities** | `File_Utilities\Launch_Duplicate_File_Finder.bat` | Find duplicate files and conflicts |
| 8 | **Deploy Inventor Add-In** | `InventorAddIn\Deploy_AddIn.bat` | Install/Update Assembly Cloner Add-In for Inventor (shows installed status) |

---

### 3️⃣ **RESCUE & SYNCHRONIZATION** (Deep Blue)
*Troubleshooting and rescue tools*

| # | Script Name | Function | Description |
|---|-------------|----------|-------------|
| 9 | **Smart Prefix Scanner** | `Part_Renaming\Launch_Smart_Prefix_Scanner.bat` | **BEFORE renaming**: Scan model to detect and prevent duplicate part numbers |
| 10 | **Emergency IDW Fixer** | `IDW_Updates\Launch_Emergency_IDW_Fixer.bat` | **TROUBLESHOOTING**: Fix specific IDW folders missed by Step 4. Use when IDW Updates didn't catch all drawings |
| 11 | **IDW-Assembly Synchronizer** | `IDW_Updates\Launch_IDW_Assembly_Synchronizer.bat` | **TROUBLESHOOTING**: Sync scattered IDWs when they use old assembly names. Use when IDWs reference renamed assembly |

---

### 4️⃣ **CLONING TOOLS** (Gray)
*Assembly and part cloning workflow*

| # | Script Name | Function | Description |
|---|-------------|----------|-------------|
| 12 | **Assembly Cloner** | `Part_Renaming\Launch_Assembly_Cloner.bat` | Clone assembly + parts to new folder. **CLONING WORKFLOW STEP 1** |
| 13 | **Cloner (Prefix Changer Only)** | `Part_Renaming\Launch_Prefix_Cloner.bat` | Clone assembly replacing ONLY the prefix (e.g., NCRH01-000- to NCRH02-000-). Keeps part suffixes intact |
| 14 | **Part Cloner** | `Part_Renaming\Launch_Part_Cloner.bat` | Clone individual part to new folder |
| 15 | **Fix Derived Parts (Post-Clone)** | `Part_Renaming\Launch_Fix_Derived_Parts.bat` | **[POST-CLONE]**: Fix EXTERNAL derived part references after cloning. **CLONING WORKFLOW STEP 2** |

---

### 5️⃣ **ILOGIC & ANALYSIS** (Dark Blue)
*iLogic rules and analysis tools*

| # | Script Name | Function | Description |
|---|-------------|----------|-------------|
| 16 | **iLogic Scanner** | `iLogic_Tools\Launch_iLogic_Scanner.bat` | Scan/export iLogic rules from document |
| 17 | **iLogic Patcher** | `Part_Renaming\Launch_iLogic_Patcher.bat` | **[POST-RENAME]**: Rename component references in iLogic rules after part renaming. Use after Assembly_Renamer to fix broken iLogic references |
| 18 | **Find Missing Detailed Parts** | `iLogic_Tools\Launch_Find_Missing_Detailed_Parts.bat` | Check which assembly parts haven't been detailed |

---

### 6️⃣ **SHEET METAL CONVERSION** (Orange)
*Sheet metal conversion tools*

| # | Script Name | Function | Description |
|---|-------------|----------|-------------|
| 19 | **Sheet Metal Converter (Assembly)** | `iLogic_Tools\Launch_Sheet_Metal_Converter.bat` | Convert all plate parts in assembly |
| 20 | **Sheet Metal Converter (Part)** | `iLogic_Tools\Launch_Sheet_Metal_Part_Converter.bat` | Convert single part to sheet metal |

---

### 7️⃣ **DRAWING CUSTOMIZATION** (Gray)
*Drawing and style customization tools*

| # | Script Name | Function | Description |
|---|-------------|----------|-------------|
| 21 | **Change Balloon Style** | `Title_Automation\Launch_Change_Balloon_Style.bat` | Replace all balloons with selected style |
| 22 | **Change Dimension Style** | `Title_Automation\Launch_Change_Dimension_Style.bat` | Replace all dimensions with selected style |
| 23 | **Export IDW Sheets to PDF** | `iLogic_Tools\Launch_Export_IDW_Sheets_to_PDF.bat` | Export each sheet as separate PDF |
| 24 | **Master Style Replicator** | `View_Style_Manager\Launch_Master_Style_Replicator.bat` | Copy view styling from Master View to other views |

---

### 8️⃣ **VIEW MANAGEMENT** (Teal)
*View style management tools*

| # | Script Name | Function | Description |
|---|-------------|----------|-------------|
| 25 | **Master Style Replicator** | `View_Style_Manager\Launch_Master_Style_Replicator.bat` | Copy view styling from Master View to other views (duplicate from Drawing Customization) |

---

### 9️⃣ **PARTS LIST AND BOM** (Blue)
*Parts list and cleanup tools*

| # | Script Name | Function | Description |
|---|-------------|----------|-------------|
| 26 | **Create Sheet Parts List** | `iLogic_Tools\Launch_Create_Sheet_Parts_List.bat` | Create parts list for components visible on current sheet |
| 27 | **Clean Up Unused Files** | `File_Utilities\Launch_IDW_Parts_List_Scanner.bat` | **CLEANUP**: After renaming parts, move old/unreferenced IPT files to Unrenamed Parts folder |

---

### 🔟 **PARAMETER MANAGEMENT** (Orange)
*Parameter export and fixing tools*

| # | Script Name | Function | Description |
|---|-------------|----------|-------------|
| 28 | **Length Parameter Exporter** | `iLogic_Tools\Launch_Length_Parameter_Exporter.bat` | Enable export for Length params on non-plate parts |
| 29 | **Fix Non-Plate Parts** | `iLogic_Tools\Launch_Fix_Non_Plate_Parts.bat` | Add Length2 parameter to parts missing Length |
| 30 | **Fix Single Part Length2** | `iLogic_Tools\Launch_Fix_Non_Plate_Parts.bat` | Add Length2 parameter to active part (longest dimension) |

---

## 🎯 TYPICAL WORKFLOWS

### **RENAME WORKFLOW** (For existing assemblies)
1. Smart Prefix Scanner (prevent duplicates)
2. **Assembly Renamer** (STEP 1 - REQUIRED)
3. Update Same-Folder Derived Parts (STEP 2 - OPTIONAL)
4. iLogic Patcher (STEP 3 - OPTIONAL)
5. **IDW Updates** (STEP 4 - REQUIRED)
6. Title Automation (OPTIONAL)

### **CLONE WORKFLOW** (For new assemblies)
1. **Assembly Cloner** (STEP 1 - REQUIRED)
2. Fix Derived Parts Post-Clone (STEP 2 - REQUIRED if external derived parts exist)
3. IDW Updates (if IDWs exist)
4. Title Automation (OPTIONAL)

---

## 📊 STATISTICS

- **Total Scripts**: 30 functions
- **Total Categories**: 10 categories
- **Production-Ready**: ✅ All scripts tested and working
- **UI Launcher**: SpectivLauncher.exe (C# compiled) or Launch_UI.ps1 (PowerShell)
- **Search Functionality**: Built-in search/filter
- **Status Tracking**: Progress bar and log window

---

## 🔧 ADDITIONAL FUNCTIONS BUILT INTO LAUNCHER

### UI Functions:
- **Search/Filter**: Real-time filtering of scripts by name
- **Category Tree View**: Browse by category
- **Progress Bar**: Visual feedback during script execution
- **Log Window**: Timestamped log of all script runs
- **Status Bar**: Current status display
- **About Dialog**: Version and feature information
- **Add-In Detection**: Shows installation status for Inventor Add-In

### Launcher Features:
- **Professional UI**: Clean, modern Windows Forms interface
- **Color-Coded Categories**: Each category has unique color
- **Descriptions**: Every script has detailed description
- **Error Handling**: Try/catch with user-friendly error messages
- **Working Directory Management**: Automatically restores working directory after script execution

---

## 📁 FILE LOCATIONS

### Main Scripts:
- **Core**: `Part_Renaming/`, `IDW_Updates/`, `Title_Automation/`
- **Tools**: `Registry_Management/`, `File_Utilities/`, `InventorAddIn/`
- **View**: `View_Style_Manager/`
- **iLogic**: `iLogic_Tools/`

### Launchers:
- **PowerShell**: `Launch_UI.ps1`
- **Compiled EXE**: `SpectivLauncher.exe`
- **Builder**: `Build_Launcher.bat`

### Assets:
- **Source Code**: `Assets/SpectivLauncher.cs`
- **Icon**: `Assets/icon.ico`
- **Splash**: `Assets/splash.png` (if exists)
