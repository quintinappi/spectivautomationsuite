---
description: Required workflow for Inventor Add-In changes: reference-script alignment, VS Code build, PowerShell deployment to ProgramData, hash verification, and guided user testing/log triage.
applyTo: **/InventorAddIn/**
---

# Inventor Add-In Agent Workflow (Required)

## 1) Source-of-truth behavior before coding
- If the user asks to migrate or fix a tool and does not specify the exact legacy script, first search for likely reference scripts in relevant folders (for example `File_Utilities`, `IDW_Updates`, `iLogic_Tools`, `Migration to Add-In`, and add-in source folders).
- If multiple plausible scripts are found, summarize candidates and ask the user which one is the intended behavior source.
- If one clear script is found, proceed and explicitly mirror its behavior in the add-in.
- Do not assume behavior from memory when a reference script exists.

## 2) UX requirements for add-in tools
- Keep UI intuitive and low-friction.
- No typing for selection tasks when avoidable.
- Use dropdowns, checkboxes, or pickers for user selections.
- Preserve existing UX patterns used by current add-in forms.

## 3) Build workflow (always from VS Code terminal)
- Build the add-in from the workspace using solution/project build commands (Release x64 unless user asks otherwise).
- Use: `dotnet build --configuration Release /p:Platform=x64` to build to `bin\x64\Release\`
- Report build result and warnings/errors.
- Do not skip build after code changes.

## 4) Deployment workflow (PowerShell + Roaming priority for reliability)
- **Primary target**: `%APPDATA%\Autodesk\Inventor 2026\Addins` (roaming) - works without elevation
- **Secondary target**: `C:\ProgramData\Autodesk\Inventor 2026\Addins` (requires elevation - may not work)
- Deploy using PowerShell copy commands (no elevation needed for roaming):
  ```powershell
  Copy-Item "InventorAddIn\AssemblyClonerAddIn\bin\x64\Release\AssemblyClonerAddIn.dll" "$env:APPDATA\Autodesk\Inventor 2026\Addins\AssemblyClonerAddIn.dll" -Force
  Copy-Item "InventorAddIn\AssemblyClonerAddIn\AssemblyClonerAddIn.addin" "$env:APPDATA\Autodesk\Inventor 2026\Addins\AssemblyClonerAddIn.addin" -Force
  ```
- ProgramData deployment requires administrator privileges and may fail even with elevation attempts.
- Inventor loads from roaming location when ProgramData deployment fails.

## 5) Post-deploy verification (required)
- Verify hashes after deploy using PowerShell:
  ```powershell
  $srcHash = (Get-FileHash "InventorAddIn\AssemblyClonerAddIn\bin\x64\Release\AssemblyClonerAddIn.dll" -Algorithm SHA256).Hash
  $deployedHash = (Get-FileHash "$env:APPDATA\Autodesk\Inventor 2026\Addins\AssemblyClonerAddIn.dll" -Algorithm SHA256).Hash
  "Match: $($srcHash -eq $deployedHash)"
  ```
- Confirm deployed hash matches source before asking user to test.
- If ProgramData deployment is attempted, verify those hashes too.

## 6) User testing handoff
- After verified deployment, instruct the user to open/reopen Inventor and test the specific workflow.
- Ask for exact failing behavior (sheet, command, field, expected vs actual) if issues remain.

## 7) Logging and troubleshooting
- Ensure tool errors are logged clearly.
- When user reports failure, inspect relevant logs and map evidence to code paths.
- Provide concise, actionable troubleshooting steps.

## 8) Communication expectations for future sessions
- Do not make the user restate this workflow.
- Reuse these steps automatically for all add-in tasks under `InventorAddIn`.
- If unsure which legacy logic to mirror, ask once with concrete options.