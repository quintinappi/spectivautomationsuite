# STEP 2: OPTION B - ASSEMBLY-BY-ASSEMBLY IDW UPDATER

## 🎯 What This Does

Uses **STEP 1's proven method** applied **assembly-by-assembly** to update IDWs.

## ✅ Why Option B Works

### **Architecture:**
```
1. Open Structure.iam (already open in Inventor)
2. Iterate through direct child sub-assemblies:
   - Column-1.iam, Column-2.iam, ...
   - Beam-1.iam, Beam-2.iam, ...
   - Channel-1.iam, ...
   - Staircase-1.iam, ...
3. For each sub-assembly:
   - Check if IDW exists in same folder
   - Use STEP 1's UpdateAllIDWsInDirectory method
   - Log success/failure for THIS assembly
4. Move to next assembly (isolated failures)
```

### **Key Advantages:**
- ✅ **Proven method:** Uses exact same logic as STEP 1 (which works)
- ✅ **Isolated failures:** Beam-17 fails ≠ everything fails
- ✅ **Clear logging:** Per-assembly status (success/failed/skipped)
- ✅ **Easy debugging:** Failure = check that specific assembly
- ✅ **Controlled recursion:** Max 1 level deep (safe)
- ✅ **No complex 4-tier resolution:** Direct mapping lookup

## 📊 How It Works

### **Processing Flow:**
```
For each sub-assembly in Structure.iam:

  1. Get assembly path (e.g., Column-1\Column-1.iam)
  2. Check for IDW in same folder (Column-1\MGY-100-SCR-01-50.idw)

  3. If IDW exists:
     - Open IDW
     - Read referenced files
     - Lookup each file in STEP 1 mapping
     - Update references using fd.ReplaceReference
     - Save IDW
     - Log: "✓ Column-1: Updated 1 IDW with 4 references"

  4. If IDW not found:
     - Log: "SKIP: Column-1 - No IDW found"

  5. If error occurs:
     - Log: "✗ Column-1: FAILED - [error details]"
     - Continue to next assembly (don't stop everything)
```

### **Example Log Output:**
```
---
ASSEMBLY: Column-1.iam
FOLDER: D:\Pentalin\...\Structure\Column-1
IDW FOUND: MGY-100-SCR-01-50.idw
IDW: Opening MGY-100-SCR-01-50.idw
IDW: Found 4 referenced files
IDW:   Updating Part1.ipt → NCRH01-000-PL1.ipt
IDW:   ✓ Updated
IDW:   Updating Part2 203 x 203 x 46 H.ipt → NCRH01-000-B1.ipt
IDW:   ✓ Updated
IDW:   Updating Part3.ipt → NCRH01-000-PL2.ipt
IDW:   ✓ Updated
IDW:   Updating Part4.ipt → NCRH01-000-PL3.ipt
IDW:   ✓ Updated
IDW: Saved with 4 reference updates
✓ SUCCESS: Column-1.iam
---
ASSEMBLY: Beam-1.iam
...
```

## 🚀 How to Use

### **Prerequisites:**
1. STEP 1 completed (mapping file exists)
2. Inventor is running
3. Structure.iam is open

### **Run:**
```
Launch_IDW_Reference_Updater_OptionB.bat
```

### **What Happens:**
1. Script connects to Inventor
2. Loads STEP 1 mapping file (329 entries)
3. Gets Structure.iam's sub-assemblies
4. Processes each one individually
5. Shows summary: "50 assemblies, 48 successful, 2 failed"

## 📋 Success Criteria

### **Full Success:**
- All sub-assemblies processed
- All IDWs updated
- Log shows "✓ SUCCESS" for each

### **Partial Success:**
- Some assemblies succeed, some fail
- Check log for specific failures
- Re-run after fixing issues

### **Failure Scenarios:**

| **Error** | **Meaning** | **Solution** |
|-----------|-------------|--------------|
| "No IDW found" | Assembly has no drawing | Normal - skip |
| "Could not open IDW" | File locked or corrupted | Close Inventor, retry |
| "No mapping found" | Part not in STEP 1 mapping | Check mapping file |
| "ReplaceReference failed" | Invalid path or file missing | Verify heritage file exists |

## 🔄 Fallback: Option A

If Option B has issues, we can fall back to **Option A** (no recursion):
- Process only direct children
- No recursion at all
- Even simpler, even safer

## 🛡️ Safety Features

1. **Backup created:** `IDW_Reference_Updater_BACKUP_2025-09-30.vbs`
2. **Isolated failures:** One assembly breaks ≠ all break
3. **Clear logging:** Per-assembly success/failure
4. **Controlled depth:** Max 1 level recursion
5. **Document closure:** Closes all docs between IDW updates (prevents locks)

## 📊 Comparison: Option B vs Old STEP 2

| **Feature** | **Old STEP 2** | **Option B** |
|-------------|----------------|--------------|
| **Method** | 4-tier resolution | STEP 1's proven method |
| **Processing** | All IDWs at once | Assembly-by-assembly |
| **Failures** | One fails = all obscured | Isolated per assembly |
| **Logging** | Generic | Per-assembly detailed |
| **Complexity** | High (4 tiers) | Low (direct lookup) |
| **Reliability** | Fragile | Proven |

## ✅ Ready to Test

**Current status:**
- ✅ Backup created
- ✅ Option B built
- ✅ Launcher created
- ⏳ Ready for testing

**Next step:** Run `Launch_IDW_Reference_Updater_OptionB.bat` with Structure.iam open.

### 2026-01-16: Search Scope Fix (Critical) ✅
- **Summary:** Fixed a bug where STEP 2 sometimes crawled unrelated project folders (e.g., `MGY-200-*`) and processed IDWs outside the selected mapping project.
- **Root Cause:** The script was *adding parent directories* of mapping entries to the search list and then performing a *recursive* search. This allowed the search to escalate to `...\3. Working\` and recursively visit other projects.
- **Fix implemented:** Removed parent-directory addition in `GetDirectoriesFromMapping()` and made `FindIDWsInDirectory()` non-recursive (searches only the exact directories from the mapping file).
- **Verification:** Ran the updated script against `TEMP - RENAME TEST` mapping. Script found only the expected IDWs (88 files) and applied updates successfully.
- **Known issue:** The final report counter showed `0` updates in one run despite actual updates occurring. Root cause: reporting variable not incremented/reset correctly in `UpdateSingleIDWWithDesignAssistantMethod()`; **action:** logging confirms updates — fix to reporting will be applied next.

**Recommended validation steps (post-fix):**
1. Run STEP 2 with known mapping file and confirm `Found: <N> IDW files` equals expected number from mapping directories. ✅
2. Spot-check 3–5 IDWs and verify references point to the new (renamed) files. ✅
3. Confirm no IDWs from other projects appear in the log. ✅

---

