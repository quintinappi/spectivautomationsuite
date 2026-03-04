# CRITICAL FIX - STEP 2 IDW UPDATER - October 1, 2025

## THE BUG

**Symptom:** STEP 2 mapped every IDW reference to the FIRST heritage file in each group:
- Column-7's Part1.ipt → NCRH01-000-PL1.ipt (Column-1's file!) ❌
- Column-6's Part1.ipt → NCRH01-000-PL1.ipt (Column-1's file!) ❌
- Beam-17's Part1.ipt → NCRH01-000-PL1.ipt (Column-1's file!) ❌

**Result:** Every column/beam got the same wrong references.

## ROOT CAUSE

**File:** `IDW_Reference_Updater_NEW.vbs`
**Lines 250-255 (OLD CODE):**

```vbscript
currentFileName = GetFileNameFromPath(fd.FullFileName)  ' Extracted ONLY filename
newPath = FindHeritagePathByOriginalFilename(currentFileName)  ' Searched by filename only
```

**Function `FindHeritagePathByOriginalFilename` (lines 284-312):**
- Searched mapping by **filename only** (Part1.ipt, Part2.ipt, etc.)
- Returned **FIRST match** found
- Ignored folder paths completely

**Why this failed:**
- Every folder has its own "Part1.ipt"
- Mapping file has full paths: `D:\...\Column-7\Part1.ipt` vs `D:\...\Column-1\Part1.ipt`
- Function threw away the path, searched by filename "Part1.ipt"
- Always found Column-1's entry first (line 4 of mapping file)
- Returned Column-1's NCRH01-000-PL1.ipt for ALL Part1.ipt references

## THE FIX

**Key insight:** The mapping dictionary is **already keyed by full path** (line 346), and `fd.FullFileName` **already provides the full path**!

**Solution:** Use full path directly as dictionary key instead of extracting filename and searching.

**NEW CODE (lines 250-272):**

```vbscript
' Get FULL PATH from IDW reference (not just filename!)
Dim currentFullPath
currentFullPath = fd.FullFileName

Dim currentFileName
currentFileName = GetFileNameFromPath(currentFullPath)

' Direct dictionary lookup using full path as key
Dim newPath
newPath = ""

If g_ComprehensiveMapping.Exists(currentFullPath) Then
    Dim mappingValue
    mappingValue = g_ComprehensiveMapping.Item(currentFullPath)

    ' Parse mapping: originalPath|newPath|originalFile|newFile|group|description
    Dim mappingParts
    mappingParts = Split(mappingValue, "|")

    If UBound(mappingParts) >= 1 Then
        newPath = mappingParts(1) ' New path is field #1
    End If
End If
```

**Also deleted:** Entire `FindHeritagePathByOriginalFilename` function (no longer needed)

## RESULT

**BEFORE (BROKEN):**
```
Column-7\Part1.ipt → looks up "Part1.ipt" → finds Column-1's PL1 → WRONG ❌
```

**AFTER (FIXED):**
```
Column-7\Part1.ipt → looks up full path → finds Column-7's PL18 → CORRECT ✓
```

## TESTING REQUIRED

⚠️ **IMPORTANT:** This bug corrupted IDW references. Before running the fixed STEP 2:

1. **Restore backup** from before STEP 2 ran (September 30)
2. **Verify STEP 1 mapping file** is still correct (329 entries, full paths)
3. **Run fixed STEP 2**
4. **Spot check 5 IDWs** across different folders to verify correct mapping:
   - Column-1: Should reference NCRH01-000-PL1, PL2, PL3, B1
   - Column-7: Should reference NCRH01-000-PL18, B5
   - Beam-17: Should reference its own unique heritage files

## LESSONS LEARNED

1. **Never discard context:** We had full paths, threw them away, then struggled to reconstruct context
2. **Use existing data structures:** Dictionary was already keyed correctly - just use it!
3. **Test with duplicate filenames:** Every folder having "Part1.ipt" should have revealed this bug immediately
4. **Validate results carefully:** Seeing "4 references updated" doesn't mean they were updated CORRECTLY

## FILES CHANGED

- `IDW_Reference_Updater_NEW.vbs` - Fixed (October 1, 2025)
  - Lines 244-291: Replaced filename-only lookup with full-path dictionary lookup
  - Lines 284-312: Deleted obsolete `FindHeritagePathByOriginalFilename` function

## VERIFICATION

After fix, log should show:
```
Column-7/Part1.ipt → NCRH01-000-PL18.ipt ✓ Updated
Column-6/Part1.ipt → NCRH01-000-PL16.ipt ✓ Updated
```

NOT:
```
Column-7/Part1.ipt → NCRH01-000-PL1.ipt ✓ Updated (WRONG!)
Column-6/Part1.ipt → NCRH01-000-PL1.ipt ✓ Updated (WRONG!)
```

---

**Fixed by:** Debug Specialist + Claude Code
**Date:** October 1, 2025
**Status:** ✅ READY FOR TESTING
