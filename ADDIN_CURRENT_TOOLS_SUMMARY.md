# Current Inventor Add-In Tools Summary

## Assembly Cloner Add-In with iLogic Patcher
**Version:** 1.0
**Author:** Quintin de Bruin © 2025
**Platform:** Inventor 2026 (.NET Framework 4.8)

---

## 🎯 9 RIBBON TOOLS (Currently Installed)

### **Assembly Ribbon** (Tools Tab → Cloner Tools Panel)

| # | Button | Icon/Color | Module | Function | Status |
|---|--------|------------|--------|----------|--------|
| 1 | **Clone Assembly** | C / Blue | AssemblyCloner.vb | Clone assembly with all parts to new folder, patch iLogic rules automatically | ✅ Working |
| 2 | **Scan iLogic** | I / Green | iLogicPatcher.vb | Scan current document for iLogic rules and display details | ✅ Working |
| 3 | **View Document Info** | D / Orange | DocumentInfoScanner.vb | Display iProperties, Mass, and iLogic rules for current document | ✅ Working |
| 4 | **Part Renamer** | R / Red | PartRenamer.vb | Rename assembly parts with heritage method (uses EXACT same logic as Assembly_Renamer.vbs) | ✅ Working |
| 5 | **Smart Inspector** | S / Dark Gray | AssemblySmartInspector.vb | Inspect current assembly, list parts, show iLogic parameters/forms (experimental) | ✅ Working |
| 6 | **Beam Assembly Generator** | B / Teal | BeamAssemblyGenerator.vb | Generate parametric beam assembly with endplates (SANS steel sections) | ✅ Working |
| 7 | **Update Document Settings** | U / Pink | PlateDocumentSettings.vb | Apply document settings to fix BOM decimal precision for plate parts | ✅ Working (enhanced diagnostics) |
| 8 | **Place Parts in IDW** | V / Cyan | PartPlacer.vb | Scan assembly for PL/S355JR parts and place them in a new IDW at 1:1 scale | ✅ NEW |

### **Part Ribbon** (Tools Tab → Cloner Tools Panel)

| # | Button | Icon/Color | Module | Function | Status |
|---|--------|------------|--------|----------|--------|
| 9 | **Part Cloner** | P / Purple | Built-in handler | Clone individual part to new folder, update iProperties and iLogic | ✅ Working |

*Note: Scan iLogic, View Document Info, Smart Inspector, Beam Generator, Update Document Settings, and Part Placer also appear in Part ribbon.*

---

## 🤖 AUTOMATIC FEATURES (Always Active)

### **Plate Part Monitor**
- **Module:** PlateDocumentSettings.vb
- **Trigger:** OnDocumentSave event
- **Function:** Automatically removes decimals from BOM quantities when saving plate parts
- **Status:** ✅ Active with refresh-cycle diagnostics for precision/dirty/BOM stages
- **Workaround:** Use **Update Document Settings** command on assembly when BOM display is stale

---

## 📂 ADD-IN MODULES (Source Code)

### **Core Modules** (Keep - Working)

| File | Purpose | Key Functions | Lines of Code |
|------|---------|---------------|---------------|
| `StandardAddInServer.vb` | Entry point, creates ribbon buttons | Activate(), CreateUserInterface(), button handlers | ~630 lines |
| `AssemblyCloner.vb` | Assembly copying with iLogic patching | CloneAssembly(), CopyAssemblyRecursively(), PatchAllRules() | ~500 lines |
| `PartRenamer.vb` | Heritage-based renaming with client classification | RenameAssemblyParts(), GetPartGrouping(), CreateHeritagePart() | ~800 lines |
| `iLogicPatcher.vb` | iLogic rule text replacement | ScanAndDisplayRules(), PatchRules(), GetRuleText() | ~300 lines |
| `DocumentInfoScanner.vb` | Document analysis | ScanAndDisplayInfo(), GetPropertySets() | ~200 lines |
| `AssemblySmartInspector.vb` | Assembly validation | InspectActiveAssembly(), ListOccurrences() | ~400 lines |
| `BeamAssemblyGenerator.vb` | Parametric beam creation | ShowGeneratorForm(), CreateBeamAssembly() | ~600 lines |
| `PlateDocumentSettings.vb` | BOM decimal fixing | ProcessCurrentDocument(), FixBOMDecimals() | ~350 lines |
| `PartPlacer.vb` | Place parts in IDW | Execute(), ScanAssemblyForMatchingParts(), CreateIDWWithParts() | ~800 lines |

### **UI Forms** (Keep - If used)

| File | Purpose | Status |
|------|---------|--------|
| `BeamGeneratorForm.vb` | Beam creation UI | ✅ Used by Beam Generator |
| `AssemblyInspectorForm.vb` | Inspector UI | ✅ Used by Smart Inspector |

### **Data Files** (Check before removing)

| File | Purpose | Status |
|------|---------|--------|
| `SteelSectionData.vb` | Steel section database | ⚠️ Check if used by Beam Generator |
| `InspectorModels.vb` | Inspector data models | ⚠️ Check if used by Smart Inspector |

---

## 🔧 TECHNICAL DETAILS

### **Add-In Identification**
- **ClassId:** `{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}`
- **ClientId:** `{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}`
- **DisplayName:** "Assembly Cloner with iLogic Patcher"
- **Description:** "Clone assemblies, rename parts, patch iLogic rules"

### **Deployment Paths**
- **User Add-Ins:** `%APPDATA%\Autodesk\Inventor 2026\Addins\AssemblyClonerAddIn\`
- **All Users:** `C:\ProgramData\Autodesk\Inventor 2026\Addins\AssemblyClonerAddIn\`
- **Compiled DLL:** `bin\x64\Release\AssemblyClonerAddIn.dll` (144 KB)
- **Manifest:** `AssemblyClonerAddIn.addin`

### **Dependencies**
- **Autodesk.Inventor.Interop.dll** - Inventor 2026 API
- **Autodesk.iLogic.Interfaces.dll** - iLogic API
- **System.Windows.Forms** - UI components
- **System.Drawing** - Graphics/icons

### **Build Configuration**
- **Framework:** .NET Framework 4.8
- **Platform:** x64 (64-bit)
- **Configuration:** Release
- **IDE:** Visual Studio 2022

---

## 🎨 ICON SYSTEM

Icons are dynamically generated using `CreateGlyphPicture()` function:

| Button | Letter | Color | RGB |
|--------|--------|-------|-----|
| Clone Assembly | C | Blue | 33, 150, 243 |
| Scan iLogic | I | Green | 76, 175, 80 |
| Document Info | D | Orange | 255, 152, 0 |
| Part Renamer | R | Red | 255, 87, 34 |
| Smart Inspector | S | Dark Gray | 33, 33, 33 |
| Beam Generator | B | Teal | 0, 150, 136 |
| Update Doc Settings | U | Pink | 233, 30, 99 |
| Part Cloner | P | Purple | 156, 39, 176 |
| Place Parts in IDW | V | Cyan | 0, 188, 212 |

Icons created at 16x16 and 32x32 pixels with colored backgrounds and white letters.

---

## 📊 USAGE STATISTICS

- **Total Tools:** 9 ribbon buttons
- **Automatic Features:** 1 (Plate Part Monitor)
- **Source Modules:** 9 VB.NET files
- **Total Lines of Code:** ~4,580 lines
- **Development Time:** ~40+ hours (estimated)
- **Status:** Production ready (1 feature under investigation)

---

## 🚀 VS. EXTERNAL VBSCRIPT TOOLS

### **Add-In Advantages:**
- ✅ Native Inventor ribbon integration
- ✅ Direct access to Inventor API (no COM late binding)
- ✅ Can read/write iLogic rules (VBScript cannot)
- ✅ Automatic event monitoring (OnDocumentSave)
- ✅ Professional UI with icons
- ✅ Faster execution (compiled vs interpreted)

### **VBScript Advantages:**
- ✅ Easy to edit (text files)
- ✅ No compilation required
- ✅ 30 tools already working
- ✅ Battle-tested on production projects
- ✅ Flexible file operations

### **Complementary Use:**
- **Add-In** = Core workflows + iLogic integration + automation
- **VBScript** = Complex batch operations + file management + utilities

---

## 📝 DEVELOPMENT NOTES

### **What Works Well:**
1. ✅ **Assembly Cloner** - Copies entire hierarchies with reference updates
2. ✅ **iLogic Patcher** - Reads and modifies rule source code (unique feature!)
3. ✅ **Part Renamer** - Client-based classification (PL, B, CH, A, FL, etc.)
4. ✅ **Beam Generator** - Creates parametric assemblies with SANS steel sections
5. ✅ **Smart Inspector** - Lists parts, shows parameters, validates assembly

### **What Needs Work:**
1. ⚠️ **Edge-case BOM refresh validation** - precision refresh path now logs detailed trace points
   - **Current approach:** Precision/display/expression toggle + units toggle + dirty/save + BOM refresh
   - **Diagnostics:** Review local add-in log for per-part outcomes and refresh stage details
   - **Next check:** Confirm behavior across large assemblies and mixed modifiable/read-only parts

### **Future Enhancements:**
- [ ] Preview dialog before cloning
- [ ] Undo/backup functionality
- [ ] Batch processing multiple assemblies
- [ ] Integration with registry management
- [ ] Better error handling and logging

---

## 📖 REFERENCE

**Quick Start Guide:** `QUICK_START_GUIDE.md`
**Technical Documentation:** `README.md`
**Deployment Script:** `DEPLOY_NOW.bat`
**Diagnostic Tool:** `DIAGNOSE_ADDIN.bat`

**Integration Plan:** `INVENTOR_ADDIN_INTEGRATION_PLAN.md` (this document)

---

**Last Updated:** March 4, 2026
**Version:** 1.2
**Status:** Production Ready (with targeted diagnostics enabled)
