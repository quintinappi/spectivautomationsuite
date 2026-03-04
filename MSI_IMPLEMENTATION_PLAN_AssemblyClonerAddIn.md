# MSI Implementation Plan (AssemblyClonerAddIn)

**Date:** 2026-03-03  
**Target app:** Spectiv Inventor Automation Suite 2026 (`AssemblyClonerAddIn`)  
**Project:** `InventorAddIn/AssemblyClonerAddIn/AssemblyClonerAddIn.vbproj`

## 1) Packaging Decision

## Recommended now: Visual Studio Installer Project (`.vdproj`)
- Fastest path to a valid MSI for Autodesk App Store submission.
- Team already uses Visual Studio and .NET Framework 4.8 tooling.
- Easier to maintain in the short term than introducing WiX migration during submission crunch.

## Optional later: WiX migration
- Move to WiX only after first approved store version if you need stronger CI/CD automation and finer upgrade control.

## 2) Build + Payload Source

## Build command (Release x64)
From workspace root:
- `dotnet build "InventorAddIn\AssemblyClonerAddIn\AssemblyClonerAddIn.vbproj" -c Release -p:Platform=x64`

Expected payload from:
- `InventorAddIn/AssemblyClonerAddIn/bin/x64/Release/AssemblyClonerAddIn.dll`
- `InventorAddIn/AssemblyClonerAddIn/AssemblyClonerAddIn.addin`

## 3) Installer Scope + Destination

## Scope (recommended)
- **Per-user install** for reliability and no admin prompt.

## Destination
- `%APPDATA%\Autodesk\Inventor 2026\Addins\`

Install these files directly into that folder:
- `AssemblyClonerAddIn.dll`
- `AssemblyClonerAddIn.addin`

Reason:
- Matches the add-in workflow guidance and avoids ProgramData elevation failures.

## 4) Visual Studio Setup Project Tasks

1. Add **Setup Project** to `AssemblyClonerAddIn.sln` (name: `AssemblyClonerAddIn.Setup`).
2. Add content files:
   - Primary output from `AssemblyClonerAddIn` (Release/x64)
   - `AssemblyClonerAddIn.addin` (explicit file include)
3. Configure target folder to user profile:
   - `[AppDataFolder]Autodesk\Inventor 2026\Addins\`
4. Product metadata:
   - ProductName: `Spectiv Inventor Automation Suite 2026`
   - Manufacturer: `Spectiv`
   - Version: semantic (`1.0.0` etc.)
   - UpgradeCode: fixed GUID (do not change across versions)
5. Add uninstall support (default in MSI).
6. Add license/EULA dialog if required by your commercial policy.
7. Build outputs:
   - `AssemblyClonerAddIn.Setup.msi`
   - optional `setup.exe` bootstrapper

## 5) Versioning + Upgrade Rules

- Keep `UpgradeCode` constant across releases.
- Increment `ProductVersion` each release.
- Ensure major upgrade behavior removes old product cleanly.
- Keep assembly/file versions aligned with MSI version policy.

## 6) Signing

- Sign both:
  - `AssemblyClonerAddIn.dll`
  - MSI (`.msi`) and bootstrapper (`setup.exe`, if shipped)
- Use Authenticode timestamping in signing pipeline.

## 7) Prerequisites

## Required
- Autodesk Inventor 2026 (or supported listed versions).
- .NET Framework 4.8 runtime (normally present on Win10/11; still verify).

## Not required for this add-in package
- **Inventor Apprentice** as a standalone prerequisite.

Notes:
- This is an in-process Inventor add-in. Apprentice is typically only required for separate headless tooling workflows.

## 8) Validation Checklist (release gate)

1. Clean VM test (no dev environment assumptions).
2. Install MSI as non-admin user.
3. Confirm files exist in `%APPDATA%\Autodesk\Inventor 2026\Addins\`.
4. Launch Inventor → Add-In loads at startup.
5. Smoke test critical commands:
   - Assembly clone flow
   - Parameter export flow
   - Registry/tool UI entry points
6. Uninstall and verify:
   - Files removed
   - Add-in no longer loads
7. Reinstall upgrade build and confirm upgrade path works.

## 9) Submission Artifacts Produced by This Plan

- Store-uploadable installer package (`.msi`, optional `setup.exe`)
- Deterministic install path and uninstall behavior
- Documented prerequisites and compatibility statement
- Signed binaries/package ready for trust/review

## 10) Immediate Next Steps

1. Create `AssemblyClonerAddIn.Setup` project in Visual Studio.
2. Wire payload files and AppData add-ins destination.
3. Build MSI and run clean-machine install/uninstall tests.
4. Sign package and generate SHA256 hashes for release evidence.
