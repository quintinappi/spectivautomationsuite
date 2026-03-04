# TESTING CHECKLIST - NEW PLANT PROJECT

**Purpose:** Step-by-step checklist for testing the renaming workflow on a new plant
**Created:** September 30, 2025
**Status:** Ready for testing

---

## 📋 PRE-TEST PREPARATION

### **□ Step 1: Identify Target Project**
- [ ] Project name: ___________________
- [ ] Plant/Project code: ___________________
- [ ] Prefix to use: ___________________
- [ ] Main assembly file: ___________________
- [ ] Expected part count: ___________________

### **□ Step 2: Create Backup**
- [ ] Backup location: ___________________
- [ ] Backup timestamp: ___________________
- [ ] Backup size verified: ___________ MB
- [ ] Backup accessible and openable: ✓

### **□ Step 3: Document Current State**
- [ ] Screenshot of project folder structure
- [ ] Note any existing renamed files (if any)
- [ ] Document current registry state (run Registry Scanner)
- [ ] List of sub-assemblies to verify (pick 5):
  1. ___________________
  2. ___________________
  3. ___________________
  4. ___________________
  5. ___________________

---

## 🚀 PHASE 1: STEP 1 - PART RENAMING

### **□ Pre-Execution Checklist**
- [ ] Inventor is running
- [ ] Main assembly is open (e.g., Structure.iam)
- [ ] No other documents open (close all extras)
- [ ] Old STEP_1_MAPPING.txt deleted (if exists)
- [ ] Registry cleared (if starting fresh): ✓ / N/A

### **□ Execute STEP 1**
- [ ] Run: `FINAL_PRODUCTION_SCRIPTS\Part_Renaming\Launch_Assembly_Renamer.bat`
- [ ] **Dialog 1 - Assembly Confirmation:**
  - [ ] Part count matches expected: _____ (expected: _____)
  - [ ] Folder location correct: ✓
  - [ ] Clicked "Yes" to proceed
- [ ] **Dialog 2 - Prefix Input:**
  - [ ] Entered prefix: ___________________
  - [ ] Verified spelling before pressing OK
- [ ] **Monitoring:**
  - [ ] Script is running (window shows progress)
  - [ ] Estimated time: 20-30 minutes for 500 parts
  - [ ] Start time: ___________________
  - [ ] End time: ___________________
  - [ ] Duration: ___________________

### **□ Post-Execution Validation**
- [ ] **Success dialog appeared:** ✓
- [ ] **Check mapping file:**
  - [ ] File exists: `STEP_1_MAPPING.txt`
  - [ ] Line count: _____ (expected: 300-500+)
  - [ ] Sample lines checked:
    ```
    Line 4: ___________________
    Line 50: ___________________
    Line 100: ___________________
    ```
  - [ ] All paths point to production folder (not test/): ✓
  - [ ] Correct prefix used throughout: ✓

- [ ] **Check heritage files created:**
  - [ ] Sub-assembly 1: Heritage files exist: ✓
  - [ ] Sub-assembly 2: Heritage files exist: ✓
  - [ ] Sub-assembly 3: Heritage files exist: ✓
  - [ ] Pattern verified: Original files + Heritage files coexist

- [ ] **Check log file:**
  - [ ] Log location: `LOGS\Assembly_Renamer_[timestamp].txt`
  - [ ] No critical errors: ✓
  - [ ] All parts processed: ✓

- [ ] **Spot check assembly references:**
  - [ ] Open one sub-assembly (e.g., Column-1.iam)
  - [ ] Check occurrence references point to heritage files
  - [ ] Example: Part1.ipt now shows as PLANTX-YYY-PL1.ipt: ✓

### **□ STEP 1 Status**
- [ ] **PASS** - All checks passed, proceed to STEP 2
- [ ] **FAIL** - Document issues below, restore backup, investigate

**Issues found (if any):**
```
___________________
___________________
```

---

## 🚀 PHASE 2: STEP 2 - IDW UPDATES

### **□ Pre-Execution Checklist**
- [ ] STEP 1 completed successfully
- [ ] Structure.iam is STILL OPEN in Inventor
- [ ] STEP_1_MAPPING.txt exists and validated
- [ ] Ready to proceed

### **□ Execute STEP 2**
- [ ] Run: `FINAL_PRODUCTION_SCRIPTS\IDW_Updates\Launch_IDW_Reference_Updater.bat`
- [ ] **Dialog - Confirmation:**
  - [ ] Clicked "Yes" to proceed
- [ ] **Monitoring:**
  - [ ] Script is running
  - [ ] Watch console for assembly names being processed
  - [ ] Should see multiple "✓ SUCCESS" messages
  - [ ] Start time: ___________________
  - [ ] End time: ___________________
  - [ ] Duration: ___________________

### **□ Post-Execution Validation**
- [ ] **Check summary dialog:**
  - [ ] Total assemblies processed: _____ (expected: 50+)
  - [ ] Successful: _____ (should be 90%+)
  - [ ] Failed/Skipped: _____ (should be low)

- [ ] **Check log file:**
  - [ ] Log location: `LOGS\IDW_Reference_Updater_OptionB_[timestamp].txt`
  - [ ] Search for "SUCCESS" count: _____
  - [ ] Search for "FAILED" count: _____
  - [ ] Review any failures: ___________________

- [ ] **Spot check IDWs (CRITICAL):**

  **IDW Test 1: Column/Beam IDW**
  - [ ] File: ___________________
  - [ ] Open in Inventor
  - [ ] Check references in design tree
  - [ ] Verify shows heritage names (PLANTX-YYY-*): ✓
  - [ ] No "Part1.ipt" references remaining: ✓

  **IDW Test 2: Channel IDW**
  - [ ] File: ___________________
  - [ ] Heritage references verified: ✓

  **IDW Test 3: Staircase IDW** (historically problematic)
  - [ ] File: ___________________
  - [ ] Heritage references verified: ✓

  **IDW Test 4: Random selection**
  - [ ] File: ___________________
  - [ ] Heritage references verified: ✓

  **IDW Test 5: Random selection**
  - [ ] File: ___________________
  - [ ] Heritage references verified: ✓

### **□ STEP 2 Status**
- [ ] **PASS** - All checks passed, proceed to STEP 3 (optional)
- [ ] **FAIL** - Document issues below, investigate

**Issues found (if any):**
```
___________________
___________________
```

---

## 🚀 PHASE 3: TITLE UPDATES (OPTIONAL)

### **□ Pre-Execution Checklist**
- [ ] STEP 2 completed successfully
- [ ] IDWs verified to reference heritage files
- [ ] Ready for title formatting

### **□ Execute STEP 3**
- [ ] Run: `FINAL_PRODUCTION_SCRIPTS\Title_Automation\Launch_Title_Updater.bat`
- [ ] **Dialog - Confirmation:**
  - [ ] Clicked "Yes" to proceed
- [ ] **Monitoring:**
  - [ ] Script processes IDWs
  - [ ] Duration: ___________________

### **□ Post-Execution Validation**
- [ ] **Spot check titles:**
  - [ ] IDW 1: View title updated: ✓
  - [ ] IDW 2: View title updated: ✓
  - [ ] IDW 3: View title updated: ✓
  - [ ] Format correct (bold, proper scale): ✓

### **□ STEP 3 Status**
- [ ] **PASS** - Titles updated correctly
- [ ] **FAIL** - Document issues, may need IDW migration first

---

## 📊 FINAL VERIFICATION

### **□ Overall Health Check**
- [ ] **Random sampling:** Open 10 random IDWs across project
  - [ ] All reference heritage files (not Part1.ipt): ✓
  - [ ] No broken references: ✓
  - [ ] Views display correctly: ✓

- [ ] **Assembly check:** Open 3 random sub-assemblies
  - [ ] All reference heritage files: ✓
  - [ ] No missing components: ✓

- [ ] **File count verification:**
  - [ ] Original files still exist (.ipt): ✓
  - [ ] Heritage files created (PLANTX-YYY-*.ipt): ✓
  - [ ] Both coexist in same folders: ✓

### **□ Documentation**
- [ ] Take "after" screenshots of folder structure
- [ ] Document final part count: _____
- [ ] Document any issues encountered: ___________________
- [ ] Note time taken: ___________________

### **□ Success Criteria**
- [ ] ✅ **FULL SUCCESS** - All steps passed, no issues
- [ ] ⚠️ **PARTIAL SUCCESS** - Minor issues, documented below
- [ ] ❌ **FAILED** - Major issues, restore backup required

---

## 🎯 LESSONS LEARNED

**What went well:**
```
___________________
___________________
```

**What could be improved:**
```
___________________
___________________
```

**Unexpected issues:**
```
___________________
___________________
```

**Time comparison:**
- Expected duration: ___________________
- Actual duration: ___________________

---

## 📁 TEST RESULTS ARCHIVE

**Save this file as:**
`TESTING_RESULTS_[PLANT_NAME]_[DATE].md`

**Attach supporting files:**
- [ ] STEP_1_MAPPING.txt
- [ ] Log files from all 3 steps
- [ ] Screenshots (before/after)
- [ ] List of any failures/issues

---

## 🚨 EMERGENCY ROLLBACK

**If test fails catastrophically:**

1. **STOP** - Don't run more scripts
2. **Close Inventor** - Shut down application
3. **Restore backup:**
   ```
   Delete: [Project Folder]
   Copy: [Backup Folder] → [Project Folder]
   ```
4. **Verify restoration:**
   - [ ] Open main assembly
   - [ ] Verify original file names intact
   - [ ] No heritage files present
5. **Document failure** in this checklist
6. **Review CRITICAL_LESSONS_LEARNED.md** for guidance

---

**Test Completed By:** ___________________
**Date:** ___________________
**Plant/Project:** ___________________
**Overall Result:** PASS / PARTIAL / FAIL

---

**Next Steps:**
- [ ] If PASS: Proceed with production use
- [ ] If PARTIAL: Review issues, decide if acceptable
- [ ] If FAIL: Investigate root cause, fix, retest
