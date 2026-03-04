# Autodesk App Store Submission Checklist (Pre-Filled)

Date: 2026-03-04  
Product: Spectiv Inventor Automation Suite 2026 (AssemblyClonerAddIn path)

Status legend:
- ✅ Done
- 🟡 Partial / needs verification
- ❌ Missing

## 1) Packaging & Technical Readiness

| Item | Status | Evidence / Notes |
|---|---|---|
| Native Inventor add-in project exists | ✅ | `InventorAddIn/AssemblyClonerAddIn/AssemblyClonerAddIn.vbproj` |
| Add-in manifest present | ✅ | `InventorAddIn/AssemblyClonerAddIn/AssemblyClonerAddIn.addin` |
| Add-in set to load on startup | ✅ | `<LoadOnStartUp>1</LoadOnStartUp>` in manifest |
| Release x64 payload built | ✅ | `InventorAddIn/AssemblyClonerAddIn/bin/x64/Release/` contains DLL + dependencies |
| MSI/EXE installer artifact produced | ❌ | No installer artifact found in active Release output |
| Installer project source in active path | ❌ | No active `.vdproj`/`.wxs` installer project found in checked active paths |
| Code signing (installer + binaries) | 🟡 | Required/recommended; no signed artifact evidence captured |
| Deployment workflow documented | ✅ | `.github/instructions/add-in builds.instructions.md`, `INSTALLER_CREATION_GUIDE.md` |

## 2) Store Listing Content

| Item | Status | Evidence / Notes |
|---|---|---|
| Product name finalized | ✅ | `Spectiv Inventor Automation Suite 2026` in `.addin` |
| Short + long description finalized | 🟡 | Draft/commercialization material exists; final store copy not identified |
| Compatibility statement finalized | 🟡 | Draft references exist; final App Store compatibility entry not assembled |
| Category selection finalized | ❌ | No final category selection artifact identified |
| Pricing model finalized | 🟡 | Pricing strategy draft exists in commercialization plan |
| Screenshots set (5-10) prepared | ❌ | No curated screenshot set found in active asset paths |
| App icon prepared | 🟡 | `Assets/icon.ico`, `Assets/spectiv-icon.png` present; store dimension compliance not verified |
| Demo video prepared (recommended) | ❌ | No demo video artifact identified |

## 3) Legal & Support

| Item | Status | Evidence / Notes |
|---|---|---|
| Privacy policy (publishable URL) | ❌ | Not identified in active repo artifacts |
| EULA/license agreement | ❌ | Not identified in active repo artifacts |
| Support contact + response policy | ❌ | Not identified as finalized submission artifact |
| Installation/user documentation | 🟡 | Multiple docs exist; no final app-store submission pack identified |
| Trademark/attribution review | 🟡 | Mentioned in planning docs; final legal review artifact not identified |

## 4) QA & Release Evidence

| Item | Status | Evidence / Notes |
|---|---|---|
| Fresh install validation | ❌ | No completed test report attached |
| Uninstall validation | ❌ | No completed test report attached |
| Upgrade-over-previous validation | ❌ | No completed test report attached |
| Hash verification captured | ❌ | Workflow exists, but no captured run evidence included |
| Submission bundle assembled | ❌ | No final zip/package checklist execution artifact identified |

## 5) Immediate Actions (Execution Order)

1. Build installer project and produce signed MSI/EXE.
2. Run install/uninstall/upgrade validation and save results.
3. Finalize legal package: Privacy Policy URL, EULA, Support policy/contact.
4. Prepare final App Store assets: screenshots, final icon validation, listing copy.
5. Assemble and review final submission bundle against this checklist.

## 6) Local Evidence References

- `InventorAddIn/AssemblyClonerAddIn/AssemblyClonerAddIn.vbproj`
- `InventorAddIn/AssemblyClonerAddIn/AssemblyClonerAddIn.addin`
- `InventorAddIn/AssemblyClonerAddIn/bin/x64/Release/`
- `AUTODESK_STORE_COMPLIANCE.md`
- `AUTODESK_APP_STORE_COMMERCIALIZATION_PLAN.md`
- `.github/instructions/add-in builds.instructions.md`
- `INSTALLER_CREATION_GUIDE.md`

## 7) Inventor Add-In Feature Inventory (Store-Ready)

### 7.1 Enabled in current add-in build (19 commands)

1. **Clone Assembly** — Clones top assembly + subassemblies + parts, updates references, patches iLogic, updates IDW refs, updates iProperties.
2. **Assembly Renamer** — Heritage-based part renaming with grouping/classification and mapping support.
3. **Registry Management** — Manages renamer counters/prefix registry data.
4. **Title Automation (IDW)** — Standardizes IDW view title formatting for active sheet or all sheets.
5. **Auto Ballooner** — Auto-places balloons on active sheet with nearest-edge leader placement and spacing logic.
6. **Populate DWG REF from Parts Lists** — Scans non-DXF sheets/parts lists and writes DWG REF values to models and parts lists.
7. **Populate DWG REF + Auto-place Missing Parts** — Runs DWG REF update and auto-places missing part views.
8. **CREATE DXF FOR MODEL PLATES** — Builds DXF sheet(s), places plate views, and creates filtered parts list output.
9. **Place parts from open Assembly** — Auto-builds ISO/BASE/PARTS sheet sets from an open assembly.
10. **Create Sheet Parts List** — Creates parts list for active sheet, filters to visible sheet parts, and renumbers rows.
11. **Create GA Parts List (Top Level)** — Generates top-level GA parts list from selected assembly source.
12. **Clean Up Unused Files** — Finds unreferenced IPT files in drawing folder and moves them to backup folder.
13. **Length Parameter Exporter** — Enables export of `Length` for non-plate parts.
14. **Length2 Parameter Exporter** — Enables export of `Length2` for non-plate parts.
15. **Thickness Parameter Exporter** — Enables export of `Thickness` for plate parts.
16. **Fix Non-Plate Parts** — Adds/updates `Length2` for non-plate parts across active assembly.
17. **Fix Single Part Length2** — Adds/updates `Length2` on active part using longest model dimension.
18. **Fix BOM Plate Dimensions** — Sets plate `LENGTH`/`WIDTH` BOM properties from sheet metal dimensions.
19. **Apply Plate Desc/Stock Formula** — Final-step formula injection for plate Description and Stock Number fields.

### 7.2 Additional add-in commands in phased migration (present in code, hidden in UI)

- Part Renamer (legacy variant), Part Cloner, Scan iLogic, Document Info, Smart Inspector, Beam Generator, Update Document Settings, BOM Precision, Auto Detail IDW, Update Same Folder Derived Parts, IDW Updates, File Utilities, Unused Part Finder, Deploy Inventor Add-In, Smart Prefix Scanner, Prefix Changer Only Cloner, Fix Derived Parts Post Clone, Sheet Metal Converter (Assembly), Sheet Metal Converter (Part), Change Balloon Style, Change Dimension Style, Export IDW Sheets to PDF, Master Style Replicator.

## 8) Pros, Value, and Time-Saving (Inventor Add-In only)

### 8.1 Key pros

- **Native Inventor ribbon workflow**: no context switching for core detailing/automation commands.
- **Consistency at scale**: standard naming, titles, parts lists, parameters, and DWG REF formatting.
- **Error reduction**: fewer manual reference edits and fewer repetitive sheet-by-sheet updates.
- **Traceability**: tool-level logging and deterministic workflows improve troubleshooting and QA.
- **Coverage across workflow**: supports model prep, renaming, drawing creation/cleanup, and BOM/property normalization.

### 8.2 Practical time-saving statement (safe for listing)

- Internal workflow guidance indicates typical automated phase durations of:
	- **Part renaming**: ~20–30 minutes
	- **Drawing reference updates**: ~10–15 minutes
	- **Title updates**: ~5–10 minutes
- On medium-to-large assembly/drawing packages, this generally replaces several manual operations and can save **hours of repetitive effort per project cycle** (actual savings vary by project size/quality).

### 8.3 Suggested App Store copy block

**Short description**

Native Autodesk Inventor add-in automation suite for assembly cloning, renaming, drawing updates, parts list generation, plate DXF workflows, and parameter/BOM preparation.

**Key benefit bullets**

- Clone assemblies with automatic reference and iLogic patch updates.
- Automate IDW title, balloon, DWG REF, and parts-list workflows.
- Generate plate DXF sheet sets directly from assembly context.
- Standardize parameter exports and BOM-related plate properties.
- Reduce repetitive detailing effort while improving output consistency.
