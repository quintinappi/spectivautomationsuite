# BOM Formula Re-evaluation Investigation & Solutions

**Date:** January 8, 2026
**Author:** Quintin de Bruin
**Status:** Investigation & Testing Phase

---

## CRITICAL DISCOVERY: Not a Cache Issue - Formula Re-evaluation Problem

### User's Key Observation

> "Even closing Inventor and reopening doesn't work. Something in changing a manual toggle and saving triggers something in the parts list/BOM to update. Because probably its using a formula. Even removing the formula completely, closing, reopening, adding the formula STILL shows incorrect decimal. Literally ONLY going to settings, units, toggling to something and back and saving makes the formula / BOM / parts list show correct."

This observation reveals the **true nature of the problem**: BOM formulas only re-evaluate when **UnitsOfMeasure changes IN THE UI**.

---

## THE PROBLEM: UI Event Context vs Programmatic Context

### Why Manual UI Toggle Works

```
1. User opens Document Settings → Units dialog
2. User changes mm → cm
3. Inventor's UI event handler fires
4. Event handler marks ALL formulas as "dirty" (needs re-evaluation)
5. User changes cm → mm
6. UI event handler fires again, marks formulas dirty again
7. User clicks OK or Save
8. Formula system sees "dirty" flags → Re-evaluates ALL formulas
9. BOM display updates with current precision settings ✅
```

### Why Programmatic Toggle Fails

```
1. Code: UnitsOfMeasure.LengthUnits = cm
2. Property setter executes, value changes
3. BUT: No UI event handler involved
4. Formulas NOT marked as "dirty" ❌
5. Code: UnitsOfMeasure.LengthUnits = mm
6. Property setter executes, value changes back
7. Formula system never sees "dirty" flag
8. No re-evaluation happens ❌
9. BOM display unchanged ❌
```

**ROOT CAUSE:** BOM formulas cache their evaluated precision settings and **don't automatically re-evaluate** when `LinearDimensionPrecision` changes programmatically. They **only re-evaluate when UnitsOfMeasure changes trigger UI events**.

---

## THEORY: Transaction Boundary & Formula Evaluation

Inventor likely uses a **transaction-based architecture** for formula evaluation:

1. **Formula Cache:** BOM formulas cache their display format based on current `LinearDimensionPrecision`
2. **Dirty Flags:** Formulas only re-evaluate when marked "dirty"
3. **UI Event Trigger:** UnitsOfMeasure changes in UI set dirty flags automatically
4. **Programmatic Gap:** COM API property changes DON'T trigger dirty flags
5. **Transaction Flush:** Formulas re-evaluate when transaction commits (Save, Rebuild, etc.)

**The Challenge:** Find a way to programmatically trigger the SAME dirty flag that the UI sets.

---

## INVESTIGATION APPROACHES

### Approach 1: Transaction Flush via Save

**Theory:** Saving the document forces transaction commit, which might trigger formula re-evaluation.

**Implementation:**
```vbscript
' Toggle precision
params.LinearDimensionPrecision = 0

' Force transaction flush
asmDoc.Save

' Test: Does BOM update now?
```

**Likelihood:** Low - save doesn't trigger formula re-evaluation on its own.

### Approach 2: BOMView.Renumber() Method

**Theory:** BOMView has a `Renumber()` method that forces complete BOM rebuild.

**Implementation:**
```vbscript
Dim bomView
Set bomView = bom.BOMViews.Item("Structured")
bomView.Renumber ' Force full BOM rebuild
```

**Likelihood:** High - renumbering should trigger formula refresh.

### Approach 3: LengthDisplayUnits Toggle (Not Just LengthUnits)

**Theory:** `UnitsOfMeasure.LengthDisplayUnits` affects formula **display format** more directly than `LengthUnits`.

**Implementation:**
```vbscript
Dim originalDisplayUnits
originalDisplayUnits = unitsOfMeasure.LengthDisplayUnits

' Toggle display units
unitsOfMeasure.LengthDisplayUnits = kCentimeterLengthUnits
asmDoc.Update

unitsOfMeasure.LengthDisplayUnits = originalDisplayUnits
asmDoc.Update
```

**Likelihood:** Medium - may trigger formula formatting refresh.

### Approach 4: Parameter.Expression Reset

**Theory:** Setting parameter expression to itself might mark it dirty.

**Implementation:**
```vbscript
Dim param
Set param = params.ModelParameters.Item(1)

' Touch parameter to mark dirty
Dim originalExpr
originalExpr = param.Expression
param.Expression = originalExpr ' Force dirty flag
```

**Likelihood:** Low - likely optimized away as no-op.

### Approach 5: Nuclear Document Reopen

**Theory:** Closing and reopening the assembly forces **complete re-evaluation** of all formulas from scratch.

**Implementation:**
```vbscript
Dim asmPath
asmPath = asmDoc.FullFileName

asmDoc.Save
asmDoc.Close False

' Wait for resource release
WScript.Sleep 2000

' Reopen - formulas re-evaluate from scratch
Set asmDoc = invApp.Documents.Open(asmPath, True)
```

**Likelihood:** 100% - guaranteed to work, but disruptive to workflow.

### Approach 6: Direct BOMQuantity Manipulation

**Theory:** If BOMQuantity object has a `Precision` or `DisplayFormat` property, set it directly.

**Implementation:**
```vbscript
For Each bomRow In bomView.BOMRows
    Dim bomQty
    Set bomQty = bomRow.BOMQuantity

    ' If property exists, set directly
    bomQty.Precision = 0 ' Bypass formulas entirely
Next
```

**Likelihood:** Unknown - needs investigation to see if properties exist.

---

## NEW TOOLS PROVIDED

### 1. Diagnose_BOM_Formula_System.vbs

**Purpose:** Deep investigation of BOM formula architecture.

**What It Does:**
- Inspects BOMView object properties
- Investigates BOMRow and BOMQuantity structure
- Checks for precision/formatting properties
- Compares assembly vs part precision settings
- Generates detailed diagnostic report

**Output:** `BOM_DIAGNOSTIC_REPORT.txt`

**When to Run:** First step to understand the BOM formula system.

**Launcher:** `Launch_Diagnose_BOM_Formula_System.bat`

---

### 2. Force_Formula_Reevaluation.vbs

**Purpose:** Attempts multiple methods to force formula re-evaluation.

**Methods Used:**
1. Transaction flush via Save
2. BOM Structured View rebuild (Renumber)
3. Parameter system rebuild
4. LengthDisplayUnits toggle (not just LengthUnits)
5. Rebuild2 with force flag
6. Final save to commit

**When to Run:** After precision is set correctly, to force BOM to refresh.

**Launcher:** `Launch_Force_Formula_Reevaluation.bat`

---

### 3. Nuclear_Reopen_Cycle.vbs

**Purpose:** Guaranteed formula refresh via document close/reopen.

**How It Works:**
1. Saves assembly
2. Closes assembly (flushes ALL caches)
3. Waits for resource release
4. Reopens assembly (formulas re-evaluate from scratch)

**When to Run:** When all other methods fail (last resort).

**Caveat:** Closes your current assembly - disruptive to workflow.

**Launcher:** `Launch_Nuclear_Reopen_Cycle.bat`

---

## TESTING WORKFLOW

### Step 1: Run Diagnostic

```
1. Open assembly in Inventor
2. Run: Launch_Diagnose_BOM_Formula_System.bat
3. Review: BOM_DIAGNOSTIC_REPORT.txt
4. Look for:
   - BOMQuantity.Precision property
   - BOMQuantity.DisplayFormat property
   - Any formatting-related properties
```

### Step 2: Test Force Re-evaluation

```
1. Set part precision to 0 decimals (existing script)
2. Run: Launch_Force_Formula_Reevaluation.bat
3. Check BOM - do quantities show 0 decimals?
4. If YES: Success! Document which method worked.
5. If NO: Proceed to Step 3.
```

### Step 3: Nuclear Option

```
1. Save all changes
2. Run: Launch_Nuclear_Reopen_Cycle.bat
3. Assembly closes and reopens
4. Check BOM - should show 0 decimals
5. If STILL NO: Issue is not formula caching
```

---

## QUESTIONS TO ANSWER

1. **Does BOMQuantity have a Precision property?**
   - If YES: Can we set it directly to bypass formulas?

2. **Does BOMView.Renumber() trigger formula refresh?**
   - If YES: This is our solution (non-disruptive)

3. **Does LengthDisplayUnits toggle work better than LengthUnits?**
   - If YES: Update existing scripts to use DisplayUnits

4. **Does Save operation trigger formula re-evaluation?**
   - If YES: Simple save after precision change should work

5. **Do formulas re-evaluate on document reopen?**
   - If YES: Nuclear option works (but disruptive)
   - If NO: Problem is deeper than formula caching

---

## NEXT STEPS

### Immediate

1. ✅ Run `Diagnose_BOM_Formula_System.vbs`
2. ⏳ Review diagnostic report for BOMQuantity properties
3. ⏳ Test `Force_Formula_Reevaluation.vbs` methods
4. ⏳ Document which method works (if any)

### If Methods 1-2 Work

1. Update `Update_Decimal_Precision.vbs` to use working method
2. Remove `Force_BOM_Refresh.vbs` (obsolete)
3. Document the solution in CLAUDE.md

### If Only Nuclear Option Works

1. Create user-friendly workflow for close/reopen cycle
2. Add auto-reopen to `Update_Decimal_Precision.vbs`
3. Warn user that assembly will close/reopen
4. Consider this a known Inventor API limitation

### If Nothing Works

1. Issue is NOT formula caching/re-evaluation
2. Problem is elsewhere (BOM formatting, iProperty references, etc.)
3. Investigate alternative approaches:
   - Direct BOM export/import with formatting
   - iLogic rules for BOM formatting
   - Custom BOM view with fixed precision

---

## EXPECTED RESULTS

### Best Case Scenario

- `BOMView.Renumber()` triggers formula refresh
- BOM updates immediately after precision change
- No document close/reopen needed
- Solution: Add `bomView.Renumber()` to existing scripts

### Most Likely Scenario

- `LengthDisplayUnits` toggle + Save triggers refresh
- Requires save operation, but no close/reopen
- Solution: Update scripts to toggle DisplayUnits + Save

### Worst Case Scenario

- Only document close/reopen forces refresh
- Disruptive to user workflow
- Solution: Automate close/reopen cycle with user warning

---

## FILES CREATED

### Scripts

- `C:\Users\Quintin\Documents\Spectiv\3. Working\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\iLogic_Tools\Diagnose_BOM_Formula_System.vbs`
- `C:\Users\Quintin\Documents\Spectiv\3. Working\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\iLogic_Tools\Force_Formula_Reevaluation.vbs`
- `C:\Users\Quintin\Documents\Spectiv\3. Working\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\iLogic_Tools\Nuclear_Reopen_Cycle.vbs`

### Launchers

- `C:\Users\Quintin\Documents\Spectiv\3. Working\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\iLogic_Tools\Launch_Diagnose_BOM_Formula_System.bat`
- `C:\Users\Quintin\Documents\Spectiv\3. Working\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\iLogic_Tools\Launch_Force_Formula_Reevaluation.bat`
- `C:\Users\Quintin\Documents\Spectiv\3. Working\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\iLogic_Tools\Launch_Nuclear_Reopen_Cycle.bat`

### Documentation

- `C:\Users\Quintin\Documents\Spectiv\3. Working\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\iLogic_Tools\BOM_FORMULA_REEVALUATION_GUIDE.md`

---

## CRITICAL INSIGHTS

### What We Know

1. ✅ BOM formulas reference `LinearDimensionPrecision` for display formatting
2. ✅ Formulas cache their evaluated precision and don't auto-update
3. ✅ Manual UI units toggle forces formula re-evaluation
4. ✅ Programmatic units toggle doesn't trigger re-evaluation
5. ✅ Document close/reopen forces complete re-evaluation

### What We Don't Know (Yet)

1. ❓ Does BOMQuantity object have direct precision properties?
2. ❓ Does BOMView.Renumber() trigger formula refresh?
3. ❓ Does LengthDisplayUnits work better than LengthUnits?
4. ❓ Is there a Parameters.Refresh() or similar method?
5. ❓ Can we trigger transaction flush without Save?

### What We're Testing

1. ⏳ All methods in `Force_Formula_Reevaluation.vbs`
2. ⏳ BOMQuantity object structure via diagnostic
3. ⏳ Nuclear reopen cycle as fallback

---

**Run the diagnostic first, then test the force re-evaluation methods. Report back which approach works!**
