# Inventor Add-In Migration Plan (Selected Tools)

## Goal
Migrate selected working tools into the Inventor Add-In with minimal behavior drift and controlled exposure.

## Principles
- Keep current working scripts as source of truth until each add-in command is validated.
- Migrate tool-by-tool with feature flags.
- Preserve existing input/output contracts (file names, mapping files, logs, prompts).
- No UI exposure for incomplete commands.
- No external script execution from add-in commands (native Inventor add-in execution only).
- Autodesk Store target: every sellable tool must run entirely inside the add-in.

## Current Exposure
Enabled in add-in:
- Clone Assembly
- Part Renamer

Hidden/disabled for migration:
- All other commands

## Phase 1: Foundation (in progress)
1. Command catalog (stable IDs for tools)
2. Feature flags (enable/disable per tool)
3. Execution guard (blocked commands show migration message)
4. Shared service interfaces:
   - Document/context service
   - File system + path policy service
   - User prompt/dialog service
   - Logging + telemetry service

## Phase 2: Core Production Workflow
- Assembly Renamer
- Update Same-Folder Derived Parts
- iLogic Patcher
- IDW Updates
- Title Automation (IDW only)

### Acceptance Criteria
- Output files and naming conventions match current scripts.
- Side effects on assemblies/IDWs are identical or intentionally documented.

## Phase 3: Cloning + Rescue
- Assembly Cloner
- Cloner (Prefix Changer Only)
- Part Cloner
- Fix Derived Parts (Post-Clone)
- Smart Prefix Scanner

### Acceptance Criteria
- Clone trees and reference updates match existing behavior.
- No new broken references in sampled large assemblies.

## Phase 4: Drawing Customization + Parts List/BOM + View
- Change Balloon Style
- Change Dimension Style
- Export IDW Sheets to PDF
- Master Style Replicator
- Create Sheet Parts List
- Clean Up unused Files
- Populate DWG REF from Parts Lists
- Populate DWG REF + Auto-place Missing Parts
- CREATE DXF FOR MODEL PLATES
- View Management: Master Style Replicator (shared command if same)

### Acceptance Criteria
- View placement and dimension behavior match existing outputs.
- BOM/parts-list fields preserve formatting and precision expectations.

## Phase 5: Parameter + Sheet Metal + Utilities
- Length Parameter Exporter
- Length2 Parameter Exporter
- Thickness Parameter Exporter
- Fix Non-Plate Parts
- Fix Single Part Length2
- Fix BOM Plate Dimensions
- Sheet Metal Converter (Assembly)
- Sheet Metal Converter (Part)
- Registry Management
- File Utilities
- Unused Part Finder
- Deploy Inventor Add-In

### Acceptance Criteria
- Parameter creation/update rules match current logic.
- Sheet metal conversion is stable across representative part families.

## Test Strategy per Tool
1. Baseline run using current script tool
2. Add-in run on same dataset
3. Diff outputs:
   - model references
   - IDW outputs
   - log/mapping files
4. Sign-off checklist
5. Enable feature flag for production

## Risk Register (known hiccup areas)
1. Document context mismatch (Part vs Assembly vs Drawing) at command start
2. iLogic API/runtime differences between script host and add-in runtime
3. Drawing intent/retrieval behavior differences in IDW APIs
4. File lock/timing differences when replacing references
5. User prompt flow changes causing accidental behavior drift

## Rollout Model
- Internal: tool hidden but callable in debug
- Pilot: enabled for 1-2 users with rollback
- Production: enabled by default after 3 clean project validations
