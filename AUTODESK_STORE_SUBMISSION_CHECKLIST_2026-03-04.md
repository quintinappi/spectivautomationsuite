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
