# IDW Auto-Dimension Research (Inventor 2026)

Date: 2026-03-02
Goal: Assess if Inventor API can auto-place dimensions on IDW and how to integrate into current PS1/EXE workflow.

## 1) Confirmed Inventor API capability (from local interop XML)

Confirmed in local file:
- `InventorAddIn/AssemblyClonerAddIn/bin/Debug/Autodesk.Inventor.Interop.xml`

### Drawing dimension entry points
- `Sheet.DrawingDimensions`
- `DrawingDimensions.GeneralDimensions`

### Programmatic dimension creation (supported)
- `GeneralDimensions.AddLinear(...)`
- `GeneralDimensions.AddAngular(...)`
- `GeneralDimensions.AddDiameter(...)`
- `GeneralDimensions.AddRadius(...)`

### Dimension sets (supported)
- `BaselineDimensionSets.Add(...)`
- `BaselineDimensionSet.AddMember(...)`
- `ChainDimensionSets.Add(...)`
- `ChainDimensionSets.AddUsingBaseDimension(...)`
- `ChainDimensionSet.AddMembers(...)`
- `OrdinateDimensionSets.Add(...)`
- `OrdinateDimensionSet.AddMember(...)`
- `OrdinateDimensions.Add(...)`

### Geometry intent support (supported)
- `Sheet.CreateGeometryIntent(...)`
- `GeometryIntent` object is present and usable

### Retrieve model annotations to drawing (supported)
- `Sheet.GetRetrievableAnnotations(...)`
- `Sheet.RetrieveAnnotations(...)`
- `DrawingEvents.OnRetrieveDimensions(...)`

## 2) Feasibility conclusion

**Yes — auto-detailing is feasible.**

Inventor API supports both:
1. Creating dimensions from scratch using geometry intents (`AddLinear`, `AddAngular`, etc.).
2. Retrieving model annotations into drawings (`GetRetrievableAnnotations` / `RetrieveAnnotations`).

For your environment, this is best implemented in the add-in first (strong typing, better reliability), then optionally exposed in PS1 launcher UI.

## 3) Practical limits you should expect

- “Perfect” detailing for every part/view is not realistic without rule-based heuristics.
- Geometry intent selection can fail on complex/ambiguous edges (especially projected/foreshortened views).
- Collision-free text/leader placement requires iterative spacing logic.
- Hidden-line style and view orientation significantly affect which curves are usable.
- Sheet metal flat patterns need dedicated rules (different preferred dimensions than folded views).

## 4) Recommended MVP approach

## Phase A (fastest success): semi-automatic retrieval
1. For selected drawing views, call `GetRetrievableAnnotations`.
2. Filter retrieved annotations to dimensions only.
3. Call `RetrieveAnnotations`.
4. Normalize style/layer/precision to your standard.

Pros: quickest path to value, high reliability.

## Phase B (true auto-detail): rule-based generated dimensions
1. For each target `DrawingView`, collect candidate `DrawingCurve`s.
2. Build `GeometryIntent`s via `Sheet.CreateGeometryIntent`.
3. Place baseline outer dims first (overall width/height).
4. Add secondary feature dims by rule set (holes, radii, offsets).
5. Run spacing pass to reduce overlaps/crossing leaders.

## 5) Where to integrate in your current codebase

Strongest existing insertion points:
- `InventorAddIn/AssemblyClonerAddIn/PartPlacer.vb` (after each placed base view)
- `InventorAddIn/AssemblyClonerAddIn/BalloonLeaderManager.vb` (existing geometry-intent logic)
- `InventorAddIn/AssemblyClonerAddIn/StandardAddInServer.vb` (new ribbon command wiring)

Suggested new add-in module:
- `InventorAddIn/AssemblyClonerAddIn/AutoDetailer.vb`

Suggested command name:
- `Cmd_AutoDetailIDW`

## 6) PS1/EXE integration path

Your launcher chain is:
- `Assets/SpectivLauncher.cs` launches `Launch_UI.ps1`
- `Launch_UI.ps1` lists tools by `.bat` script entries

Two integration options:

### Option 1 (recommended): add-in button first
- Add `Cmd_AutoDetailIDW` in add-in ribbon (`StandardAddInServer.vb`).
- Keep launcher unchanged initially.
- Lowest risk, best debugging inside Inventor.

### Option 2: launcher menu item
- Create `File_Utilities/Launch_Auto_Detail_IDW.bat` (or add-in invoker script).
- Add menu entry to `Launch_UI.ps1` under Drawing/IDW category.
- EXE requires no code change if it already launches updated PS1.

## 7) Recommended rollout plan (safe)

1. Build MVP command: **Auto Detail Active Sheet (Outer Dims Only)**.
2. Validate on 10–20 representative drawings.
3. Add optional mode: **Retrieve model dimensions first**.
4. Add sheet-metal-specific dimension rules.
5. Expose in `Launch_UI.ps1` once stable.

## 8) Effort estimate

- MVP (outer dimensions only): ~1–2 dev days
- Robust v1 (multiple feature rules + overlap management): ~4–7 dev days
- Production hardening (edge cases + logging + toggles): +2–4 dev days

## 9) Bottom line

You can build auto-detailing with current Inventor 2026 API. The best implementation path for your stack is:
- Start as add-in command (`AutoDetailer.vb`) using `GeneralDimensions` + `GeometryIntent`
- Then surface it in launcher (`Launch_UI.ps1` + `.bat`) after validation.
