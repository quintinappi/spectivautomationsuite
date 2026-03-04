# QUICK START GUIDE - NEW PLANT TESTING

**🎯 USE THIS TOMORROW** - Fast reference for the complete workflow

---

## ⚡ 5-MINUTE OVERVIEW

### **What This Does:**
1. **STEP 1:** Renames 500+ parts with heritage method (Part1.ipt → PLANTX-YYY-PL1.ipt)
2. **STEP 2:** Updates all IDW drawings to reference new heritage files
3. **STEP 3:** (Optional) Formats view titles with proper parameters

### **Time Required:**
- Backup: 5 minutes
- STEP 1: 20-30 minutes
- STEP 2: 10-15 minutes
- STEP 3: 5-10 minutes
- **Total: ~1 hour**

---

## 🚨 BEFORE YOU START

### **1. BACKUP (CRITICAL!)**
```
Copy entire project folder:
FROM: D:\Pentalin\3. Working\[PROJECT]\05. Model
TO:   D:\Pentalin\3. Working\[PROJECT]\05. Model_BACKUP_2025-10-01
```

### **2. Open Inventor**
- Open main assembly (e.g., Structure.iam)
- Close any other documents
- Make sure assembly loads completely

### **3. Clear Registry (if starting fresh)**
```
Run: Registry_Management\Registry_Manager.vbs
Select: Scan (to view current state)
If needed: Manually clear using Windows Registry Editor
```

### **4. Delete Old Mapping (if re-running)**
```
Delete: [ASSEMBLY_FOLDER]\STEP_1_MAPPING.txt
(Only if this is a re-run after previous failure - mapping is now saved per-assembly)
```

---

## 🚀 THE WORKFLOW

### **STEP 1: Part Renaming** ⏱️ 20-30 minutes

**Run:**
```
FINAL_PRODUCTION_SCRIPTS\Part_Renaming\Launch_Assembly_Renamer.bat
```

**Dialog 1 - Confirm Assembly:**
- Check part count (should match expected)
- Check folder location
- Click **YES**

**Dialog 2 - Enter Prefix:**
- Type prefix carefully: `PLANTX-YYY-`
- Double-check spelling
- Click **OK**

**Wait for completion** (shows "SUCCESS" dialog)

**Validate:**
```
✓ [ASSEMBLY_FOLDER]\STEP_1_MAPPING.txt exists (300-500+ lines)
✓ Heritage files created in folders (PLANTX-YYY-*.ipt)
✓ Assembly references updated
✓ Check log: LOGS\Assembly_Renamer_[timestamp].txt
```

---

### **STEP 2: IDW Updates** ⏱️ 10-15 minutes

**Run:**
```
FINAL_PRODUCTION_SCRIPTS\IDW_Updates\Launch_IDW_Reference_Updater.bat
```

**Make sure:** Structure.iam is STILL OPEN

**Dialog - Confirm:**
- Click **YES**

**Watch console:** Should process 50+ assemblies (NOT just 1!)

**Validate:**
```
✓ Processed: 50+ assemblies
✓ Successful: 90%+
✓ Open 5 random IDWs → verify heritage references
✓ Check log: LOGS\IDW_Reference_Updater_OptionB_[timestamp].txt
```

---

### **STEP 3: Title Updates** ⏱️ 5-10 minutes (OPTIONAL)

**Run:**
```
FINAL_PRODUCTION_SCRIPTS\Title_Automation\Launch_Title_Updater.bat
```

**Validate:**
```
✓ View titles formatted correctly
✓ Scale parameters working
✓ Bold text applied
```

---

## ✅ SUCCESS CRITERIA

### **STEP 1 Success:**
- [ ] 300-500 entries in mapping file
- [ ] All paths point to correct production folder
- [ ] Heritage files exist alongside originals
- [ ] Log shows all parts processed

### **STEP 2 Success:**
- [ ] 50+ assemblies processed (NOT 1!)
- [ ] High success rate (90%+)
- [ ] Spot-checked IDWs show heritage references
- [ ] No "Part1.ipt" remaining

### **Overall Success:**
- [ ] Open any IDW → see PLANTX-YYY-PL47.ipt (not Part1.ipt)
- [ ] All drawings functional
- [ ] No broken references

---

## 🚨 RED FLAGS (STOP IF YOU SEE THESE!)

| **Warning Sign** | **Meaning** | **Action** |
|------------------|-------------|------------|
| **"220 skipped"** | STEP 2 broke (old bug) | Should NOT happen with fix |
| **"Processed: 1"** | STEP 2 stopped early | Should NOT happen with fix |
| **Mapping has 4 lines** | Wrong folder in mapping | Delete, re-run STEP 1 |
| **"test-222-" prefix** | Wrong prefix used | Delete mapping, re-run |
| **Structure.iam closes** | Script error | Should NOT happen with fix |

---

## 📊 QUICK VALIDATION COMMANDS

### **Check Mapping File:**
```powershell
# Count lines (update path to your assembly folder)
Get-Content "C:\Path\To\Assembly\STEP_1_MAPPING.txt" | Measure-Object -Line

# Check first few entries
Get-Content "C:\Path\To\Assembly\STEP_1_MAPPING.txt" | Select-Object -First 10

# Verify prefix used
Select-String "PLANTX-YYY" "C:\Path\To\Assembly\STEP_1_MAPPING.txt" | Measure-Object
```

### **Check Heritage Files:**
```powershell
# Count heritage files in one folder
Get-ChildItem "D:\...\Structure\Column-1" -Filter "PLANTX-YYY-*.ipt" | Measure-Object
```

### **Check Log for Errors:**
```powershell
# Search for errors
Select-String "ERROR" "LOGS\Assembly_Renamer_*.txt"
Select-String "FAILED" "LOGS\IDW_Reference_Updater_*.txt"
```

---

## 🔧 TROUBLESHOOTING

### **Problem: "No Inventor instance found"**
- **Solution:** Start Inventor first, open assembly

### **Problem: "Mapping file not found"**
- **Solution:** STEP 1 didn't complete, run STEP 1 first

### **Problem: "No IDW found for assembly"**
- **Solution:** Normal if assembly has no drawing

### **Problem: STEP 2 shows "220 skipped"**
- **Solution:** This is the OLD BUG - should NOT happen with fixed version
- Check you're using the NEW IDW_Reference_Updater.vbs
- OLD version is archived as IDW_Reference_Updater_OLD_BROKEN.vbs

---

## 📁 FILES TO CHECK

### **Before Starting:**
- [ ] `CRITICAL_LESSONS_LEARNED.md` - Read the two fatal mistakes
- [ ] `TESTING_CHECKLIST_NEW_PLANT.md` - Full detailed checklist

### **After Each Step:**
- [ ] `[ASSEMBLY_FOLDER]\STEP_1_MAPPING.txt` - Validate mapping
- [ ] `LOGS\*.txt` - Check for errors
- [ ] Random IDW files - Spot check references

---

## 🎯 THE GOLDEN RULES

1. **NEVER HARDCODE** - Always scan folders dynamically
2. **NEVER ASSUME** - IDW names ≠ assembly names
3. **NEVER CLOSE DURING ITERATION** - Breaks parent context
4. **ALWAYS USE PROVEN METHODS** - Don't overcomplicate

---

## ⏱️ ESTIMATED TIMELINE

```
09:00 - Read documentation (this file + CRITICAL_LESSONS_LEARNED.md)
09:15 - Create backup
09:20 - Start STEP 1
09:50 - STEP 1 completes → validate
10:00 - Start STEP 2
10:15 - STEP 2 completes → validate
10:30 - Spot check 5 IDWs
10:45 - (Optional) Run STEP 3
11:00 - Final verification → DONE
```

---

## 📞 EMERGENCY ROLLBACK

**If something goes catastrophically wrong:**

1. **STOP** - Don't run more scripts
2. **Close Inventor**
3. **Delete project folder**
4. **Copy backup back**
5. **Verify restoration**
6. **Review logs to understand what failed**

---

## ✅ FINAL CHECKLIST

Before declaring success:
- [ ] STEP 1 completed with 300-500 mappings
- [ ] STEP 2 processed 50+ assemblies successfully
- [ ] 5 random IDWs checked - all show heritage references
- [ ] No "Part1.ipt" references remaining
- [ ] Assemblies open without errors
- [ ] Backup safely stored

**If all checked:** ✅ **SUCCESS!**

---

## 🛠️ RESCUE TOOLS

If IDW updates fail due to reference mismatches:

1. **Sync IDW References:** Open IDW in Inventor → Manage → Update → Replace Model Reference → Point to current assembly parts
2. **Use Assembly Cloner (Option 9):** Creates isolated copy with updated references
3. **Emergency IDW Fixer (Option 7):** Fixes specific folders missed by Step 2

---

**Ready to proceed?** Follow `TESTING_CHECKLIST_NEW_PLANT.md` for detailed step-by-step process.

**Questions?** Review `CRITICAL_LESSONS_LEARNED.md` for deep technical explanation.

**Good luck tomorrow!** 🚀
