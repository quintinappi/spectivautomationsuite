# BOM Precision Update - Guide & Troubleshooting

## The Problem

BOM precision values (decimal places) sometimes don't update even when you change settings via API. This is because:

1. **Document "Dirty" Flag**: Inventor only refreshes BOM when a document is marked as "dirty" (modified)
2. **API Toggling**: Simply changing `LinearDimensionPrecision` via API doesn't always trigger the dirty flag
3. **BOM Caching**: BOM views cache values and need explicit refresh
4. **Part-Level Settings**: BOM uses each part's individual precision settings, not assembly settings

---

## Available Solutions

### Method 1: API-Only (Recommended First Try)
**File**: `Launch_BOM_Precision_API.bat`

**Pros:**
- No UI automation (no SendKeys)
- Fast and reliable
- Parts open invisibly in background
- No risk of focus issues

**Cons:**
- May not work on all Inventor versions
- May not trigger the "dirty" flag in some cases

**When to use:**
- As your first attempt
- When you want a quick, automated solution

---

### Method 2: Robust (With Retry Logic)
**File**: `Launch_BOM_Precision_Robust.bat`

**Pros:**
- Auto-retry on failure (3 attempts per part)
- State validation at each step
- Recovery from interruptions
- Pre-flight checks
- Can detect and handle unexpected dialogs

**Cons:**
- Slower (includes safety delays)
- Uses UI automation as fallback
- Requires Inventor to have focus

**When to use:**
- When API-only method fails
- When you need reliability over speed
- When working with problematic files

---

### Method 3: Diagnostic (Investigation)
**File**: `Launch_BOM_Precision_Diagnostic.bat`

**Pros:**
- Analyzes root cause
- Shows all precision settings
- Checks BOM structure
- Identifies iLogic rules

**Cons:**
- Doesn't fix anything
- For analysis only

**When to use:**
- Before trying fixes
- When nothing else works
- To understand the problem

---

## How the "Dirty Flag" Trick Works

The key insight: Inventor needs to think the document changed.

### The Trick (Used in API-Only version):
```vbscript
' Create a dummy parameter
Set dummyParam = params.UserParameters.AddByValue("_BOM_REFRESH_", 0, "mm")

' Change it (marks document dirty)
dummyParam.Value = 1

' Delete it (cleanup)
params.UserParameters.RemoveByName("_BOM_REFRESH_")
```

This creates a temporary parameter, changes it (which marks the document as modified), then removes it. The document is now "dirty" and will refresh properly.

---

## Recommended Workflow

```
Step 1: Run API-Only version
        ↓ (if it works, you're done!)
        ↓ (if some parts fail)
Step 2: Run Robust version for failed parts
        ↓ (if still failing)
Step 3: Run Diagnostic to investigate
        ↓
Step 4: Manual intervention for problematic parts
```

---

## Manual Fix (When All Else Fails)

For stubborn parts, manual fix always works:

1. Open the part in Inventor
2. Go to **Manage** tab → **Document Settings**
3. Click **Units** tab
4. Change any setting (e.g., Linear Dim Display Precision: 0 → 1)
5. Click **OK**
6. Change it back (1 → 0)
7. Click **OK** again
8. **Save** the part (Ctrl+S)
9. **Close** the part
10. Return to assembly and **Update** (Ctrl+U)

The key is clicking **OK** after changing a setting - this triggers the dirty flag.

---

## Troubleshooting Common Issues

### Issue: "Script gets stuck on a part"
**Cause:** Unexpected dialog (save prompt, error message, etc.)
**Solution:** 
- Press Escape a few times
- Close any open dialogs
- The Robust version will auto-retry

### Issue: "BOM still shows decimals after running"
**Cause:** BOM view not refreshed
**Solution:**
- In Inventor: Right-click BOM → **Refresh**
- Or: Toggle BOM view off/on
- Or: Close and reopen assembly

### Issue: "Some parts work, others don't"
**Cause:** Parts may have different settings or be in different states
**Solution:**
- Run Diagnostic tool to identify differences
- Try manual fix on problematic parts

### Issue: "Script says success but BOM unchanged"
**Cause:** Assembly BOM view caching
**Solution:**
- Manually refresh BOM in Inventor
- Or close/reopen assembly
- The precision WAS updated, just not displayed yet

---

## Technical Details

### What Settings Affect BOM Precision?

1. **Part-Level** (most important):
   - `UnitsOfMeasure.LengthDisplayPrecision`
   - `Parameters.LinearDimensionPrecision`

2. **Assembly-Level** (less important):
   - Same settings but BOM uses part-level values

3. **BOM-Level**:
   - `BOM.StructuredViewEnabled` (toggling refreshes view)
   - `BOMView.Rebuild` (forces recalculation)

### Why UI Automation Sometimes Works Better

When you open Document Settings UI and click OK:
- Inventor runs internal validation
- Forces document dirty flag
- Triggers update events
- Saves settings in a specific order

The API doesn't always replicate this exact sequence.

---

## Files Reference

| File | Purpose |
|------|---------|
| `Force_BOM_Precision_API_Only.vbs` | API-only method (no UI) |
| `Force_BOM_Precision_Robust.vbs` | Robust method with retry logic |
| `Diagnose_BOM_Precision.vbs` | Diagnostic analysis tool |
| `Launch_BOM_Precision_API.bat` | Launcher for API method |
| `Launch_BOM_Precision_Robust.bat` | Launcher for Robust method |
| `Launch_BOM_Precision_Diagnostic.bat` | Launcher for Diagnostic |

---

## Summary

| Scenario | Recommended Method |
|----------|-------------------|
| First attempt | API-Only |
| API failed | Robust |
| Understanding problem | Diagnostic |
| Everything fails | Manual per-part |

The API-Only version includes the "dirty flag trick" which should work for most cases. The Robust version adds retry logic for when things go wrong.
