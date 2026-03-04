# STRUCTURE ASSEMBLY TEST PLAN - STEP 1 & STEP 2 WORKFLOW

**Date Created:** September 30, 2025
**Status:** Ready to test after conversation clear

---

## 🎯 OBJECTIVE

Test that STEP 1 + STEP 2 workflow correctly:
1. Renames parts in main assembly (Structure.iam) with all sub-assemblies
2. Updates ALL IDW files recursively in subdirectories (Column-1, Column-2, Beam-1, etc.)

---

## 📋 PRE-TEST SETUP (USER WILL DO)

- [ ] Clear conversation history
- [ ] Remove Windows Registry entries (use Registry Manager)
- [ ] Delete STEP_1_MAPPING.txt
- [ ] Revert Structure.iam and all sub-assemblies to original state
- [ ] Confirm all parts are named Part1.ipt, Part2.ipt, Part3.ipt, etc.

---

## ✅ VERIFIED WORKING (FROM AUDIT)

### STEP 1 (Assembly_Renamer.vbs)
- ✅ Recursive processing of all sub-assemblies
- ✅ Heritage file creation (creates new files, keeps originals)
- ✅ Mapping file generation with correct format
- ✅ Safety check: Skips already-renamed files (heritage detection)

### STEP 2 (IDW_Reference_Updater.vbs)
- ✅ Loads mapping file correctly (line 197-246)
- ✅ **RECURSIVE IDW search** (line 334-359) - searches ALL subdirectories
- ✅ Design Assistant method: `fd.ReplaceReference(newPath)` (line 547)
- ✅ 4-tier resolution strategy (lines 465-528)
- ✅ Falls back from intelligent search to full recursive scan

---

## 🧪 TEST PROCEDURE

### STEP 1: Part Renaming
```
1. Open Structure.iam in Inventor
2. Run: FINAL_PRODUCTION_SCRIPTS\Part_Renaming\Launch_Assembly_Renamer.bat
3. Wait for completion
```

**Expected Results:**
- Heritage files created: NCRH01-000-PL193.ipt, NCRH01-000-PL194.ipt, etc.
- Original files still exist: Part1.ipt, Part2.ipt, Part3.ipt, etc.
- Mapping file created: FINAL_PRODUCTION_SCRIPTS\STEP_1_MAPPING.txt
- Assembly references updated to use heritage files

**Verification Commands:**
```cmd
# Check heritage files exist
dir "D:\Pentalin\3. Working\19. NCRH01 - Secondary Crushing Station\05. Model\Structure\Column-1\NCRH01-000-PL*.ipt" /b

# Check original files still exist
dir "D:\Pentalin\3. Working\19. NCRH01 - Secondary Crushing Station\05. Model\Structure\Column-1\Part*.ipt" /b

# Check mapping file
type "C:\Users\Quintin\inventor renamer basic\FINAL_PRODUCTION_SCRIPTS\STEP_1_MAPPING.txt" | findstr "Column-1"
```

**Expected Mapping Format:**
```
D:\...\Column-1\Part1.ipt|D:\...\Column-1\NCRH01-000-PL193.ipt|Part1.ipt|NCRH01-000-PL193.ipt|PL|PL 20mm S355JR
```

---

### STEP 2: IDW Updates (THE CRITICAL TEST)
```
1. Keep Structure.iam open in Inventor (or open any assembly)
2. Run: FINAL_PRODUCTION_SCRIPTS\IDW_Updates\Launch_IDW_Reference_Updater.bat
3. Watch the log output carefully
```

**Expected Results:**
- Script finds IDWs in ALL subdirectories:
  - Structure\Column-1\*.idw
  - Structure\Column-2\*.idw
  - Structure\Beam-1\*.idw through Beam-49\*.idw
  - Structure\Channel-1\*.idw through Channel-4\*.idw
- Updates each IDW reference from Part1.ipt → NCRH01-000-PL193.ipt
- Reports total: X IDWs processed, Y references updated

**Log File Location:**
```
FINAL_PRODUCTION_SCRIPTS\Logs\Step2_IDW_Updates_DesignAssistant_YYYYMMDD_HHMMSS.log
```

---

## 🔍 CRITICAL CHECKS DURING STEP 2

### Watch for these log messages:

**✅ GOOD SIGNS:**
```
IDW: Found X IDW files using intelligent search
IDW: Searching directory: D:\...\Structure\Column-1
IDW: Found drawing file - Column-1.idw
IDW: Processing reference: Part1.ipt
IDW: Method 1: Found exact path match in mapping
IDW: ✓ SUCCESS - Reference updated using Design Assistant method
```

**❌ BAD SIGNS:**
```
IDW: No IDW files found using intelligent search
IDW: No IDW files found in directory structure either
IDW: ERROR - Could not open
IDW: ERROR - New file doesn't exist
```

---

## 🐛 KNOWN ISSUE (FIXED)

### Previous Problem:
- STEP 1 only updated IDWs in **main assembly directory**
- Subdirectory IDWs (Column-1, Beam-1, etc.) were NOT updated
- `UpdateAllIDWsInDirectory` function only searched parent folder (line 1041 of Assembly_Renamer.vbs)

### Solution:
- STEP 2 has **recursive search** built-in (line 334-359)
- Separate STEP 2 script is the intended workflow
- STEP 1 creates mapping, STEP 2 uses mapping to update ALL IDWs

---

## 📊 SUCCESS CRITERIA

### STEP 1 Success:
- [ ] Mapping file contains entries for all parts
- [ ] Heritage files exist (NCRH01-000-*.ipt)
- [ ] Original files still exist (Part*.ipt)
- [ ] Log shows no errors

### STEP 2 Success:
- [ ] Found IDWs in multiple subdirectories (not just main folder)
- [ ] Processed IDWs from: Column-1, Column-2, Beam-1, Beam-2, etc.
- [ ] Updated references: Part1.ipt → NCRH01-000-PL193.ipt
- [ ] Total updates > 0
- [ ] Total errors = 0

### Final Verification:
```
1. Open a random sub-assembly IDW in Inventor (e.g., Column-1.idw)
2. Check references: Should show NCRH01-000-PL193.ipt (not Part1.ipt)
3. Verify drawing views display correctly with heritage parts
```

---

## 📝 NOTES FOR CLAUDE AFTER CONVERSATION CLEAR

**Context:**
- User ran STEP 1 multiple times → corrupted mapping file (heritage→heritage instead of original→heritage)
- Fixed by adding heritage detection to STEP 1 (line 232)
- Deleted corrupted mapping
- User restored original files
- Mapping now shows correct format: `Part1.ipt → NCRH01-000-PL193.ipt`

**Key Finding:**
- STEP 1 creates heritage files but doesn't update subdirectory IDWs
- This is BY DESIGN - STEP 2 is meant to update ALL IDWs recursively
- Audit confirmed STEP 2 has recursive search (line 356: `Call FindAllIDWFiles(subFolder.Path, idwFiles)`)

**What to Check:**
1. After STEP 2 runs, check log for "Found X IDW files"
2. Verify X includes IDWs from subdirectories (not just main folder)
3. Look for multiple "IDW: Searching directory" entries in log
4. Confirm "Updated: Y references" where Y > 0

**If STEP 2 Fails:**
- Check log file in FINAL_PRODUCTION_SCRIPTS\Logs\
- Look for which search method found IDWs (intelligent vs traditional)
- Verify mapping file exists and has correct format
- Check if new heritage files actually exist on disk