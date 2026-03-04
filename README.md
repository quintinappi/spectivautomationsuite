# Spectiv Inventor Automation Suite 2026

Comprehensive Autodesk Inventor automation workspace containing:

- Production VBScript utilities for drawing, assembly, naming, and registry workflows
- A VB.NET Inventor Add-In codebase (`InventorAddIn/AssemblyClonerAddIn`) used as the active modernization path
- Documentation, migration plans, and commercialization/App Store readiness artifacts

---

## Table of Contents

- [What This Repository Is](#what-this-repository-is)
- [Current State (March 2026)](#current-state-march-2026)
- [Repository Layout](#repository-layout)
- [Quick Start](#quick-start)
- [Build the Active Add-In](#build-the-active-add-in)
- [Deploy the Add-In (Inventor 2026)](#deploy-the-add-in-inventor-2026)
- [Validation Checklist (Post-Deploy)](#validation-checklist-post-deploy)
- [Autodesk App Store Readiness](#autodesk-app-store-readiness)
- [Key Documentation](#key-documentation)
- [Troubleshooting](#troubleshooting)
- [Operational Notes](#operational-notes)

---

## What This Repository Is

This workspace is the single source for Spectiv Inventor automation workstreams:

1. **Legacy/production script tooling** for day-to-day engineering automation.
2. **Modern add-in implementation** to support a compiled, commercial-ready distribution path.
3. **Commercialization and compliance planning** for Autodesk App Store submission.

The result is a hybrid repo where script workflows and add-in workflows coexist during migration.

---

## Current State (March 2026)

- Active add-in project exists at `InventorAddIn/AssemblyClonerAddIn`.
- Add-in manifest + DLL build output are present in `bin/x64/Release`.
- GitHub remote is connected and repository is now pushed.
- A pre-filled submission checklist exists at `AUTODESK_STORE_SUBMISSION_CHECKLIST_2026-03-04.md`.
- Installer packaging + legal/store assets remain the primary commercialization gaps.

---

## Repository Layout

High-level structure:

```text
INVENTOR_AUTOMATION_SUITE_2026/
├─ InventorAddIn/                          # Active VB.NET add-in path
│  └─ AssemblyClonerAddIn/
├─ InventorAutomationSuiteAddIn/           # Additional add-in experimentation/migration area
├─ IDW_Updates/                            # Drawing reference automation scripts
├─ IDW_Utilities/                          # Drawing utility scripts
├─ Part_Renaming/                          # Part numbering/renaming workflows
├─ Registry_Management/                    # Counter/registry tooling
├─ Title_Automation/                       # Title formatting automation
├─ Documentation/                          # Guides, references, audit notes
├─ Assets/                                 # Icons/branding resources
├─ Logs/                                   # Execution logs
└─ *.md                                    # Build/compliance/commercialization docs
```

---

## Quick Start

### Option A: Use existing script workflow

- Launch the suite with `Launch_Suite.bat`.
- Run individual tools via dedicated launchers (for example `Launch_Assembly_Cloner.bat`).

### Option B: Work on compiled add-in workflow

- Focus on `InventorAddIn/AssemblyClonerAddIn`.
- Build and deploy as an Inventor Add-In for 2026.

---

## Build the Active Add-In

From repository root:

```powershell
dotnet build "InventorAddIn\AssemblyClonerAddIn\AssemblyClonerAddIn.vbproj" --configuration Release /p:Platform=x64
```

Expected output:

- `InventorAddIn/AssemblyClonerAddIn/bin/x64/Release/AssemblyClonerAddIn.dll`
- `InventorAddIn/AssemblyClonerAddIn/bin/x64/Release/AssemblyClonerAddIn.addin`

If `dotnet` build is unavailable for your VS/VB setup, use Visual Studio build with:

- Configuration: `Release`
- Platform: `x64`

---

## Deploy the Add-In (Inventor 2026)

Preferred deployment target is roaming profile (no elevation required):

```powershell
Copy-Item "InventorAddIn\AssemblyClonerAddIn\bin\x64\Release\AssemblyClonerAddIn.dll" "$env:APPDATA\Autodesk\Inventor 2026\Addins\AssemblyClonerAddIn.dll" -Force
Copy-Item "InventorAddIn\AssemblyClonerAddIn\AssemblyClonerAddIn.addin" "$env:APPDATA\Autodesk\Inventor 2026\Addins\AssemblyClonerAddIn.addin" -Force
```

Secondary path (requires admin rights):

- `C:\ProgramData\Autodesk\Inventor 2026\Addins`

---

## Validation Checklist (Post-Deploy)

Verify deployment integrity using hashes:

```powershell
$srcHash = (Get-FileHash "InventorAddIn\AssemblyClonerAddIn\bin\x64\Release\AssemblyClonerAddIn.dll" -Algorithm SHA256).Hash
$dstHash = (Get-FileHash "$env:APPDATA\Autodesk\Inventor 2026\Addins\AssemblyClonerAddIn.dll" -Algorithm SHA256).Hash
"Match: $($srcHash -eq $dstHash)"
```

Then:

1. Restart Inventor.
2. Confirm add-in loads.
3. Run target workflow and check logs.

---

## Autodesk App Store Readiness

Use these as the current authoritative repo-side readiness references:

- `AUTODESK_STORE_SUBMISSION_CHECKLIST_2026-03-04.md` (current status matrix)
- `AUTODESK_STORE_COMPLIANCE.md` (technical/store constraints)
- `AUTODESK_APP_STORE_COMMERCIALIZATION_PLAN.md` (commercial + packaging roadmap)
- `INSTALLER_CREATION_GUIDE.md` (MSI packaging guidance)

### Current top blockers

1. Final installer project/artifact (MSI/EXE)
2. Legal/support submission package (privacy policy, EULA, support terms)
3. Final store media/listing assets (screenshots, listing copy, compatibility declaration)

---

## Key Documentation

Core docs to start with:

- `QUICKSTART.md`
- `BUILD_INSTRUCTIONS.md`
- `INSTALLER_CREATION_GUIDE.md`
- `INVENTOR_ADDIN_INTEGRATION_PLAN.md`
- `PROJECT_SUMMARY.md`
- `README_SUITE.md`

Add-in workflow instruction (used for engineering changes under `InventorAddIn/**`):

- `.github/instructions/add-in builds.instructions.md`

---

## Troubleshooting

### Add-in not loading in Inventor

- Confirm `.addin` and `.dll` are both in the expected Addins path.
- Confirm build target is `Release/x64`.
- Confirm version compatibility in manifest file.

### Build succeeds but behavior is stale

- Re-copy `AssemblyClonerAddIn.dll` and `.addin` to roaming Addins path.
- Re-verify SHA256 hashes.
- Restart Inventor fully.

### Git add errors on Windows (`nul`, long paths)

- This repo contains deep backup trees and reserved filename artifacts in some historical paths.
- `.gitignore` includes protective excludes for invalid/unindexable entries.

---

## Operational Notes

- This repository is intentionally broad and includes historical material, experiments, and production scripts.
- Avoid broad recursive scans over backup trees when doing tooling/search operations; prefer targeted paths.
- For add-in work, keep behavior aligned with proven script logic before refactoring.

---

For support inside this workspace, start with the checklists and guides above, then run the add-in build/deploy validation flow end-to-end.
