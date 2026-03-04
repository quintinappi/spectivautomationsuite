# Populate DWG REF from Parts Lists - Implementation Notes

## Purpose
`File_Utilities/Populate_DWG_REF_From_Parts_Lists.vbs` automates two related tasks for an open IDW:
1. Populate `DWG REF` values from Parts Lists and actual view placements.
2. Optionally place missing detailed parts onto a selected non-DXF sheet.

## Current Behavior

### 1) DWG REF population (always available)
- Scans **non-DXF** sheets only.
- Collects parts/assemblies from parts lists.
- Collects actual model placement from drawing views.
- Writes parts-list `DWG REF` column values in format:
  - `IDWNAME-01/02/...`
- Updates model user properties with `DWG REF` aliases:
  - `DWG REF`, `DWG. REF.`, `DWG_REF`, `DWGREF`
- Reports undetailed parts (`.ipt` with no non-DXF placements).

### 2) Missing-part placement mode (optional)
- Prompts user with a **dropdown** to choose a target from available non-DXF sheet names.
- Shows a pre-run notice that placement sheets may be used/created.
- Places **all missing parts**, not just one.
- Uses auto layout in a grid (`3 x 3` slots per sheet).
- If base sheet fills up, creates continuation sheets named:
  - `AUTO-MISSING-1`, `AUTO-MISSING-2`, ...

### 3) Plate handling
- Detects sheet-metal part subtype:
  - `{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}`
- For sheet-metal parts, places two views:
  - Folded (`SheetMetalFoldedModel = True`)
  - Flat pattern (`SheetMetalFoldedModel = False`)
- If no flat pattern exists, attempts to create one via `Unfold`.

## Autodesk API details used

### AddBaseView signature (Inventor 2026)
`DrawingViews.AddBaseView(Model, Position, Scale, ViewOrientation, ViewStyle, [ModelViewName], [ArbitraryCamera], [AdditionalOptions])`

### Enum values in use
- `kFrontViewOrientation = 10764`
- `kHiddenLineRemovedDrawingViewStyle = 32258`

### AdditionalOptions key used for sheet metal
- `SheetMetalFoldedModel` (Boolean)

## Runtime modes and arguments

### `dwgrefonly`
- `on`  -> only DWG REF workflow.
- `off` -> enables placement workflow if `autoplace` is on.

### `autoplace`
- `on`  -> run missing-part placement workflow.
- `off` -> no placement.

### `targetsheet` (optional)
- Can pre-fill a preferred target sheet name.
- Placement mode still validates against non-DXF sheets.

## Launchers

### DWG REF only
- `File_Utilities/Launch_Populate_DWG_REF_From_Parts_Lists.bat`
- Runs with: `/dwgrefonly:on /autoplace:off`

### DWG REF + place missing parts
- `File_Utilities/Launch_Populate_DWG_REF_From_Parts_Lists_AutoPlace_Sheet5.bat`
- Runs with: `/dwgrefonly:off /autoplace:on`
- User is prompted to select target sheet from non-DXF list.

## Related function: CREATE DXF FOR MODEL PLATES

### Script
- `File_Utilities/Create_DXF_For_Model_Plates.vbs`

### Launcher
- `File_Utilities/Launch_Create_DXF_For_Model_Plates.bat`

### UI entry
- Added under **Parts List and BOM**:
  - `CREATE DXF FOR MODEL PLATES`

### Workflow
1. Prompts sheet selection with dropdown (non-DXF sheets).
2. Finds assembly view on selected sheet.
3. Creates new sheet `DXF FOR {MODEL NAME}`.
4. Places all detected plate parts at 1:1 scale.
5. For sheet metal plates, uses flat pattern view (`SheetMetalFoldedModel=False`).
6. Creates a parts list and filters it to plate rows only.

## Logging
- Logs are written to:
  - `File_Utilities/Logs/Populate_DWG_REF_YYYYMMDD_HHMMSS.log`
- Includes mode flags, placement outcomes, per-model update outcomes, and summary.

## Notes / Constraints
- DXF sheets are excluded from scans and placement targets.
- Some models may fail to update iProperties if document/property access is blocked by Inventor state.
- Placement failures are accumulated and reported in summary/log instead of hard-stopping the run.
