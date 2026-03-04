# START HERE - Inventor Renamer Documentation Index

**Last Updated:** January 20, 2026
**Status:** ✅ Production Ready - Tested and Working

---

## 📚 DOCUMENTATION INDEX

### **🎯 FOR TOMORROW'S TESTING:**

1. **`QUICK_START_GUIDE.md`** ⚡ **START HERE FIRST**
   - 5-minute overview
   - Fast reference workflow
   - Estimated timeline: 1 hour
   - Red flags to watch for

2. **`TESTING_CHECKLIST_NEW_PLANT.md`** 📋 **USE DURING TEST**
   - Step-by-step checklist
   - Fill-in-the-blank format
   - Validation checkpoints
   - Results documentation template

3. **`CRITICAL_LESSONS_LEARNED.md`** 🚨 **READ BEFORE TESTING**
   - The two fatal mistakes we fixed
   - Why they broke everything
   - How to prevent them forever
   - Golden rules to never break

---

## 🚀 PRODUCTION SCRIPTS

### **Main Workflow:**

| **Step** | **Script** | **Purpose** | **Time** |
|----------|------------|-------------|----------|
| **1** | `Part_Renaming\Launch_Assembly_Renamer.bat` | Rename parts with heritage method | 20-30 min |
| **2** | `IDW_Updates\Launch_IDW_Reference_Updater.bat` | Update all IDW references | 10-15 min |
| **3** | `Title_Automation\Launch_Title_Updater.bat` | Format view titles (optional) | 5-10 min |

### **Utility Scripts:**

| **Tool** | **Purpose** | **When to Use** |
|----------|-------------|-----------------|
| `Part_Renaming\Assembly_Cloner.vbs` | Clone assembly with isolated parts + update IDW | Creating variants or fixing reference mismatches |
| `Part_Renaming\Part_Cloner.vbs` | Clone individual part + display iProperties | Creating part variants or backups |
| `Registry_Management\Registry_Manager.vbs` | Scan registry counters (read-only) | Before STEP 1 to check state |
| `Part_Renaming\Smart_Prefix_Scanner.vbs` | Rebuild registry from files | Before adding new assemblies |
| `Part_Renaming\Protect_Mapping_File.bat` | Hide mapping file | After STEP 1 succeeds |

---

## 🆕 Assembly Cloner (Updated January 2026)

### **What It Does:**
Copies an assembly with ALL sub-assemblies, parts, and IDW drawings to a new folder with complete isolation - no cross-references to original location.

### **Features:**
- ✅ Detects open assembly in Inventor automatically
- ✅ Dual folder selection: browse or type path
- ✅ Heritage naming for parts (WALKWAY-3-PL1.ipt, WALKWAY-3-CH1.ipt, etc.)
- ✅ Auto-renames assembly to match destination folder
- ✅ Updates all assembly references to new parts
- ✅ **Recursive sub-assembly support** - copies and updates ALL sub-assemblies
- ✅ **Recursive IDW processing** - updates ALL IDWs in ALL subfolders (Jan 20, 2026)
- ✅ Updates IDW references for both IAM and IPT files
- ✅ Auto-increments IDW filename if needed
- ✅ Generates STEP_1_MAPPING.txt for reference tracking

### **When to Use:**
- Copying an assembly folder to create a variation
- Avoiding "cross-part reference" issues between folders
- Creating isolated copies for different plant areas

### **How to Run:**
1. Open the source assembly in Inventor
2. Run `Part_Renaming\Launch_Assembly_Cloner.bat`
3. Select or type destination folder path
4. Script copies everything and updates all references

### **Future Plans:**
See `INVENTOR_PLUGIN_CONVERSION.md` for plans to convert all scripts to a native Inventor add-in with ribbon integration.

---

## ✅ WHAT'S FIXED (September 30, 2025)

### **STEP 2 IDW Updater - WORKING**

**Problem:** Processed only 1 assembly, skipped 220+ others

**Root Causes:**
1. ❌ **Hardcoded IDW name assumption** (Column-1.iam → Column-1.idw)
   - Reality: IDW names don't match assembly names
   - Fix: Dynamic folder scanning

2. ❌ **Closed Structure.iam during iteration**
   - Reality: Broke the occurrence loop
   - Fix: Only close IDW documents, not assemblies

**Result:** ✅ Now processes **50+ assemblies successfully** with **200+ IDW updates**

### **Mapping File Safety (December 9, 2025)**

**Update:** Mapping files now saved per-assembly in the assembly folder instead of shared scripts root.

**Benefits:**
- No overwriting mappings from different assemblies
- Isolated per-project mappings
- Safer for parallel processing

**Location:** `[ASSEMBLY_FOLDER]\STEP_1_MAPPING.txt`

---

## 🎯 SUCCESS CRITERIA

### **How to Know It's Working:**

**STEP 1 Success:**
- [ ] Mapping file has 300-500+ entries
- [ ] Heritage files created (PLANTX-YYY-PL1.ipt, etc.)
- [ ] Log shows all parts processed

**STEP 2 Success:**
- [ ] Console shows 50+ assemblies processed (NOT 1!)
- [ ] Final summary: "Successful: 48+ / Failed: 2"
- [ ] Random IDWs show heritage references

**Overall Success:**
- [ ] Open any IDW → see PLANTX-YYY-PL27.ipt (not Part1.ipt)
- [ ] No "Part1.ipt" references anywhere
- [ ] Drawings display correctly

---

## 🚨 GOLDEN RULES

**NEVER break these:**

1. **NEVER HARDCODE FILE NAMES**
   - Always scan folders dynamically
   - IDW names ≠ assembly names

2. **NEVER ASSUME NAMING CONVENTIONS**
   - Every project is different
   - Check what actually exists

3. **NEVER CLOSE DOCUMENTS DURING ITERATION**
   - Breaks parent assembly context
   - Only close specific document types

4. **ALWAYS USE PROVEN METHODS**
   - Don't overcomplicate
   - STEP 1's logic works - reuse it

---

## 📋 QUICK REFERENCE

### **File Locations:**

```
FINAL_PRODUCTION_SCRIPTS/
├── Part_Renaming/
│   ├── Launch_Assembly_Renamer.bat        [STEP 1]
│   └── Launch_Assembly_Cloner.bat         [Copy assembly with isolation]
├── IDW_Updates/
│   └── Launch_IDW_Reference_Updater.bat   [STEP 2]
├── Title_Automation/
│   └── Launch_Title_Updater.bat           [STEP 3]
├── STEP_1_MAPPING.txt                     [Generated by STEP 1]
├── LOGS/                                  [All execution logs]
├── README_START_HERE.md                   [This file]
├── QUICK_START_GUIDE.md                   [Fast reference]
├── TESTING_CHECKLIST_NEW_PLANT.md         [Test procedure]
├── CRITICAL_LESSONS_LEARNED.md            [What went wrong & why]
└── INVENTOR_PLUGIN_CONVERSION.md          [Future add-in conversion plan]
```

### **Typical Workflow:**

```
1. Backup project folder
2. Open Structure.iam in Inventor
3. Run STEP 1 → validate mapping file
4. Run STEP 2 → validate IDW updates
5. (Optional) Run STEP 3 → validate titles
6. Spot check 5 random IDWs → verify references
7. Done!
```

---

## 🔍 TROUBLESHOOTING QUICK REFERENCE

| **Symptom** | **Likely Cause** | **Solution** |
|-------------|------------------|--------------|
| "220 skipped" | Old bug (should NOT happen) | Verify using NEW script version |
| "Processed: 1" | Old bug (should NOT happen) | Verify using NEW script version |
| Mapping has 4 lines | Wrong folder | Delete, re-run STEP 1 on correct assembly |
| "test-222-" prefix | Wrong prefix entered | Delete mapping, re-run with correct prefix |
| "No Inventor instance" | Inventor not running | Start Inventor, open assembly |
| "Mapping not found" | STEP 1 not run | Run STEP 1 first |

---

## ⚡ TOMORROW'S PLAN

### **Preparation (15 minutes):**
1. Read `QUICK_START_GUIDE.md`
2. Read `CRITICAL_LESSONS_LEARNED.md`
3. Print or open `TESTING_CHECKLIST_NEW_PLANT.md`

### **Execution (1 hour):**
1. Create backup
2. Run STEP 1 → validate
3. Run STEP 2 → validate
4. Spot check results
5. Document in checklist

### **Verification (15 minutes):**
1. Check 5 random IDWs
2. Verify assembly count (50+, not 1!)
3. Confirm no "Part1.ipt" remaining
4. Complete test checklist

---

## 📞 SUPPORT FILES

### **For Detailed Information:**

- **Technical deep dive:** `CRITICAL_LESSONS_LEARNED.md`
- **Testing procedure:** `TESTING_CHECKLIST_NEW_PLANT.md`
- **Quick workflow:** `QUICK_START_GUIDE.md`
- **Project memory:** `C:\Users\Quintin\CLAUDE.md`

### **For Troubleshooting:**

- **Log files:** `LOGS\*.txt` (timestamped execution logs)
- **Mapping file:** `STEP_1_MAPPING.txt` (original → heritage mappings)
- **Registry state:** Run Registry_Manager.vbs to scan

---

## ✅ PRE-FLIGHT CHECKLIST

**Before tomorrow's test:**
- [ ] Read QUICK_START_GUIDE.md
- [ ] Read CRITICAL_LESSONS_LEARNED.md
- [ ] Have TESTING_CHECKLIST_NEW_PLANT.md ready to fill out
- [ ] Identify target project for testing
- [ ] Verify backup drive has space
- [ ] Know the prefix to use

**Day-of checklist:**
- [ ] Create backup FIRST
- [ ] Open Inventor
- [ ] Load main assembly
- [ ] Follow TESTING_CHECKLIST_NEW_PLANT.md exactly
- [ ] Document everything

---

## 🎯 EXPECTED RESULTS

**If everything works correctly:**

```
STEP 1 Complete:
✓ 329 mappings created
✓ Heritage files in all folders
✓ Log shows success

STEP 2 Complete:
✓ Processed: 52 assemblies
✓ Successful: 50
✓ Failed/Skipped: 2 (no IDW)

Final State:
✓ All IDWs show heritage references
✓ No "Part1.ipt" remaining
✓ Drawings functional
```

---

## 🚀 YOU'RE READY!

Everything is documented, tested, and working. Follow the guides tomorrow and it should go smoothly.

**Good luck with the new plant testing!**

---

**Quick Links:**
- Start: `QUICK_START_GUIDE.md`
- Test: `TESTING_CHECKLIST_NEW_PLANT.md`
- Learn: `CRITICAL_LESSONS_LEARNED.md`
