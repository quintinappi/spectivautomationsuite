# 🚀 QUICK START - Build Your First Inventor Add-in

## ⚡ 10 Steps to Working Add-in (30 minutes)

### Step 1: Create Project (2 min)
```
Visual Studio → Create New Project
Search: "Class Library (.NET Framework)"
Name: InventorAutomationSuiteAddIn
Framework: .NET Framework 4.8
Click: Create
```

### Step 2: Add Inventor Reference (1 min)
```
Right-click "References" → Add Reference
COM tab → Search: "Autodesk Inventor"
Check: "Autodesk Inventor Object Library"
Click: OK
```

### Step 3: Import My Files (5 min)
```
Right-click project → Add → Existing Item
Select these files from FINAL_PRODUCTION_SCRIPTS folder:
1. InventorAutomationSuiteAddIn.vb
2. MainForm.vb
3. AssemblyClonerForm.vb
4. AssemblyCloner.vb
5. RegistryManager.vb
Repeat for each file
```

### Step 4: Add System.Windows.Forms Reference (1 min)
```
Right-click "References" → Add Reference
Assemblies tab → Search: "Windows Forms"
Check: "System.Windows.Forms"
Check: "System.Drawing"
Click: OK
```

### Step 5: Configure Build Settings (2 min)
```
Right-click project → Properties
Build tab:
  ☑ Platform target: x64 (CRITICAL!)
  ☑ Register for COM interop: True
Build Events tab → Post-build event:
  copy "$(ProjectDir)InventorAutomationSuiteAddIn.xml" "$(TargetDir)"
```

### Step 6: Generate GUIDs (2 min)
```
Tools → Create GUID
Select: Registry Format
Click: New GUID
Click: Copy
Replace in files:
  1. Open InventorAutomationSuiteAddIn.vb
  2. Search: YOUR-GUID-HERE
  3. Replace all 3 occurrences with your GUID
  4. Open InventorAutomationSuiteAddIn.xml
  5. Search: YOUR-GUID-HERE
  6. Replace both occurrences with SAME GUID
Save both files
```

### Step 7: Copy XML Manifest (1 min)
```
Copy InventorAutomationSuiteAddIn.xml to project folder
Right-click project → Add → Existing Item
Select: InventorAutomationSuiteAddIn.xml
```

### Step 8: Configure Debug (2 min)
```
Project Properties → Debug tab
Start Action: ○ Start external program
External Program:
  C:\Program Files\Autodesk\Inventor 2025\Bin\Inventor.exe
(Adjust path for your Inventor version)
```

### Step 9: Build Project (1 min)
```
Build → Build Solution
Check Output window for errors
Should show: "Build succeeded"
```

### Step 10: Test in Inventor (13 min)
```
1. Press F5 (Inventor will launch)
2. Wait for Inventor to fully load
3. Tools → Add-ins Manager
4. Find: "Inventor Automation Suite"
5. Click: Load button
6. Go to: Assembly tab
7. Find: "Automation Suite" panel
8. Click: Button (opens main form)
9. Test: Assembly Cloner tool
10. Close Inventor to stop debugging
```

---

## ✅ Success Indicators

You'll know it worked when:
- ✅ Inventor launches automatically
- ✅ Add-in appears in Add-ins Manager
- ✅ Button appears in Assembly tab
- ✅ Main form opens when button clicked
- ✅ Assembly Cloner form opens
- ✅ Registry scanning works

---

## 🐛 Common Fixes

### "Cannot load Inventor API"
```
Solution: Add COM reference (Step 2)
```

### "Type Form is not defined"
```
Solution: Add System.Windows.Forms reference (Step 4)
```

### "Add-in doesn't appear"
```
Solution: Check XML manifest is in project folder (Step 7)
         Verify GUIDs match (Step 6)
         Verify "Register for COM Interop" is checked (Step 5)
```

### "Platform mismatch"
```
Solution: Set Platform Target to x64 (Step 5)
```

---

## 🎯 After It Works

Once add-in loads successfully:

1. **Test functionality:**
   - Open any assembly in Inventor
   - Click your ribbon button
   - Try Assembly Cloner tool
   - Click "Scan Registry"

2. **Debug any issues:**
   - Set breakpoints in code
   - Use F10 to step through
   - Check variable values
   - Read log file: Documents\Assembly_Cloner_Log.txt

3. **Move to installer:**
   - Follow INSTALLER_CREATION_GUIDE.md
   - Create Visual Studio Setup Project
   - Build MSI installer
   - Test install/uninstall

---

## 📞 Need Help?

**Check these files:**
- [BUILD_INSTRUCTIONS.md](BUILD_INSTRUCTIONS.md) - Detailed build guide
- [VISUAL_STUDIO_SETUP_GUIDE.md](VISUAL_STUDIO_SETUP_GUIDE.md) - Project structure
- [INSTALLER_CREATION_GUIDE.md](INSTALLER_CREATION_GUIDE.md) - Installer creation

**Common issues:**
- Wrong Inventor version path → Update in Step 8
- GUID mismatch → Replace ALL occurrences in both files
- Platform target wrong → Must be x64, not Any CPU
- XML manifest missing → Must be in same folder as DLL

---

## ⏱️ Time Estimate

- Steps 1-8: 15 minutes
- Step 9: 1 minute
- Step 10: 13 minutes (mostly waiting for Inventor)

**Total: ~30 minutes for first-time build**

---

## 🚀 You're Ready!

Everything is prepared. All files are created. Documentation is complete.

**Just open Visual Studio and follow the 10 steps above!**

Good luck! 🎉
