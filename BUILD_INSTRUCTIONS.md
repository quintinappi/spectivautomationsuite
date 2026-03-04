# Quick Build Guide - Inventor Automation Suite Add-in

## 🔧 Current Redeploy Process (InventorAddIn - March 2026)

Use this process when updating the VB.NET add-in in `InventorAddIn/AssemblyClonerAddIn`.

1. **Close Inventor completely**
   - Ensure no `Inventor.exe` remains in Task Manager.

2. **Build Release x64 from workspace root**
   - `dotnet build "InventorAddIn\AssemblyClonerAddIn\AssemblyClonerAddIn.vbproj" -c Release -p:Platform=x64`

3. **Deploy to ProgramData (elevated PowerShell)**
   - `Set-Location "c:\Users\Quintin\Documents\Spectiv\3. Working\INVENTOR_AUTOMATION_SUITE_2026\InventorAddIn"`
   - `Start-Process PowerShell -Verb RunAs -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ".\Deploy_ProgramData_Elevated.ps1"'`

4. **Sync per-user DLL (Roaming)**
   - Primary: `Copy-Item -Path ".\AssemblyClonerAddIn\bin\x64\Release\AssemblyClonerAddIn.dll" -Destination (Join-Path $env:APPDATA 'Autodesk\Inventor 2026\Addins\AssemblyClonerAddIn.dll') -Force`
   - Also sync legacy folder if present: `Copy-Item -Path ".\AssemblyClonerAddIn\bin\x64\Release\AssemblyClonerAddIn.dll" -Destination (Join-Path $env:APPDATA 'Autodesk\Inventor Addins\AssemblyClonerAddIn.dll') -Force`

5. **Verify hashes (required)**
   - Compare SHA256 for:
     - source: `InventorAddIn\AssemblyClonerAddIn\bin\x64\Release\AssemblyClonerAddIn.dll`
     - ProgramData: `C:\ProgramData\Autodesk\Inventor 2026\Addins\AssemblyClonerAddIn.dll`
   - Roaming (primary): `%APPDATA%\Autodesk\Inventor 2026\Addins\AssemblyClonerAddIn.dll`
   - Roaming (legacy): `%APPDATA%\Autodesk\Inventor Addins\AssemblyClonerAddIn.dll`
   - All must match before testing.

6. **Reopen Inventor and test**
   - If behavior is wrong, report exact command, sheet, expected vs actual output.

## 🚨 READ THIS FIRST: Critical Setup Issues

**Issue 1: C# vs VB.NET Project**
- Your code files are all `.vb` (VB.NET)
- If you created a C# project, **DELETE IT AND START OVER**
- Choose "Class Library (.NET Framework)" with the **purple Visual Basic icon**
- Verify: Solution Explorer shows `My Project` folder, not `Properties` folder

**Issue 2: System.Windows.Forms Reference Error**
- If you see error about `.tlb` file, you clicked wrong thing
- **FIX:** In Reference Manager, click **"Assemblies"** tab (NOT COM or Type Libraries)
- Search for `System.Windows.Forms` in Assemblies tab
- Add the `.dll` version, NOT the `.tlb` version

**Issue 3: Inventor 2026 COM Reference Missing (NEW!)**
- The "Autodesk Inventor Object Library" does NOT appear in COM tab for Inventor 2026
- **FIX:** Use **Browse** tab → Navigate to `C:\Program Files\Autodesk\Inventor 2026\Bin\Public Assemblies\`
- **Select:** `Autodesk.Inventor.Interop.dll`
- **Set property:** Embed Interop Types = **False**
- **DO NOT** use COM tab for Inventor 2026!

**ALTERNATIVE: Use Your Existing Working Project**
- If you have `AssemblyClonerAddIn` that already works, just modify that
- It already has the correct references configured

---

## 🎯 Files Created So Far

You now have all the core VB.NET files for the add-in:

### ✅ Core Files (Ready to Import into Visual Studio)

1. **[InventorAutomationSuiteAddIn.vb](InventorAutomationSuiteAddIn.vb)** - Main entry point
2. **[InventorAutomationSuiteAddIn.xml](InventorAutomationSuiteAddIn.xml)** - Add-in manifest
3. **[MainForm.vb](MainForm.vb)** - Unified launcher UI
4. **[AssemblyClonerForm.vb](AssemblyClonerForm.vb)** - Assembly Cloner tool UI
5. **[AssemblyCloner.vb](AssemblyCloner.vb)** - Assembly Cloner logic
6. **[RegistryManager.vb](RegistryManager.vb)** - Registry management

---

## 🚀 Step-by-Step: Build in Visual Studio

### Step 1: Create New Project

**CRITICAL: Choose VB.NET, not C#!**

1. **Open Visual Studio 2026**
2. **Create new project** (Ctrl+Shift+N)
3. **Search for:** "Class Library"
4. **CRITICAL:** Select "Class Library (.NET Framework)" with **Visual Basic** label
   - Look for the **purple VB icon**
   - ⚠️ **NOT** the blue C# icon!
   - ⚠️ Must specify ".NET Framework" (not just "Class Library")
5. **Project name:** `InventorAutomationSuiteAddIn`
6. **Framework:** .NET Framework 4.8
7. **Click Create**

**Verify it's correct:**
- Solution Explorer shows `My Project` folder (VB.NET)
- Default file is `Class1.vb` (not .cs)

### Step 2: Add Inventor Reference

**⚠️ For Inventor 2026: Use Browse method, NOT COM tab!**

1. **Right-click "References"** in Solution Explorer
2. **Select:** "Add Reference..."
3. **Click "Browse" tab** (bottom left)
4. **Navigate to:** `C:\Program Files\Autodesk\Inventor 2026\Bin\Public Assemblies\`
5. **Select:** `Autodesk.Inventor.Interop.dll`
6. **Click Add**
7. **Right-click the reference** → Properties → Set **"Embed Interop Types"** to **False**

**Optional - Also add:**
- `stdole.dll` from `C:\Program Files\Autodesk\Inventor 2026\Bin\`

### Step 3: Import the VB Files I Created

For each file I created:

1. **Right-click project** → Add → Existing Item
2. **Browse to:** `FINAL_PRODUCTION_SCRIPTS 1 Oct 2025`
3. **Select the file** and click Add
4. **Repeat for all 6 files**

**Files to add:**
- InventorAutomationSuiteAddIn.vb
- MainForm.vb
- AssemblyClonerForm.vb
- AssemblyCloner.vb
- RegistryManager.vb

**Note:** Don't add the XML file yet - we'll handle that in build events

### Step 4: Configure Project Settings

**Right-click project → Properties:**

**Application tab:**
- Root namespace: `InventorAutomationSuiteAddIn`
- Target framework: .NET Framework 4.8

**Compile tab:**
- Option Explicit: On
- Option Strict: On
- Option Infer: On

**Build tab:**
- Platform target: **x64** (CRITICAL!)
- Register for COM interop: **True**

**Build Events tab:**
- Post-build event:
  ```batch
  copy "$(ProjectDir)InventorAutomationSuiteAddIn.xml" "$(TargetDir)"
  ```

### Step 5: Fix Namespace Issues

**IMPORTANT:** All my files use namespace `InventorAutomationSuiteAddIn`, so:

1. **Open each .vb file**
2. **Verify namespace matches** your project's root namespace
3. **If different:**
   - Either change your project's Root Namespace property
   - Or do a Find/Replace in the VB files

### Step 6: Generate GUIDs

**Critical:** Replace the placeholder GUIDs:

1. **Open Visual Studio** → Tools → Create GUID
2. **Select Registry Format**
3. **Click New GUID**
4. **Copy the GUID**
5. **Replace in files:**
   - InventorAutomationSuiteAddIn.vb (line 16, 20, 25)
   - InventorAutomationSuiteAddIn.xml (line 19, 23)

**Use same GUID in both files!**

### Step 7: Add References for Windows Forms

**IMPORTANT: Don't skip this or your forms won't work!**

1. **Right-click "References"** → "Add Reference..."
2. **Click "Assemblies" tab** (left sidebar - NOT COM!)
3. **Search box:** Type "Windows.Forms"
4. **Check:** "System.Windows.Forms" (shows as .dll file)
5. **Click OK**
6. **Repeat** for "System.Drawing" if not already present

⚠️ **Reference Summary for Inventor 2026:**
- **Browse tab:** `Autodesk.Inventor.Interop.dll` (from Inventor Bin\Public Assemblies)
- **Assemblies tab:** `System.Windows.Forms`, `System.Drawing`
- **DO NOT** use COM tab (type library is missing in 2026)

### Step 8: Build the Project

1. **Set configuration to Debug**
2. **Build → Build Solution**
3. **Check for errors**

**Common errors:**
- "Type Inventor.Application not defined" → Add Inventor COM reference
- "Type Form not defined" → Add System.Windows.Forms reference
- "Platform mismatch" → Set Platform Target to x64

### Step 9: Test in Inventor

1. **Set Debug settings:**
   - Project Properties → Debug
   - Start Action: Start external program
   - External Program: `C:\Program Files\Autodesk\Inventor 2025\Bin\Inventor.exe`

2. **Press F5** to start debugging
3. **Inventor will launch**
4. **Go to:** Tools → Add-ins Manager
5. **Find:** "Inventor Automation Suite"
6. **Click Load**

### Step 10: Test the Ribbon Button

1. **Go to Assembly tab** in Inventor
2. **Look for:** "Automation Suite" panel
3. **Click button** to open main form
4. **Test Assembly Cloner tool**

---

## 🐛 Troubleshooting

### "Cannot load Inventor API"
**Solution:** Add reference to `Autodesk.Inventor.Interop.dll` via Browse tab
- Path: `C:\Program Files\Autodesk\Inventor 2026\Bin\Public Assemblies\`
- Set Embed Interop Types = False

### "Autodesk Inventor Object Library missing in COM tab"
**Solution:** This is normal for Inventor 2026!
- DO NOT use COM tab
- Use Browse tab instead (see Issue 3 at top of this guide)

### "Add-in doesn't appear in Inventor"
**Solution:**
- Check XML manifest is in output folder (bin\Debug)
- Verify GUIDs match in .vb and .xml files
- Check Register for COM Interop is enabled

### "Platform mismatch"
**Solution:** Set Platform Target to x64 in Build settings

### "Type Form is not defined"
**Solution:** Add reference to System.Windows.Forms via Assemblies tab

### "Namespace error"
**Solution:** Change project's Root Namespace to match code files

### "System.Windows.Forms .tlb error"
**Solution:** You clicked Type Libraries tab - use Assemblies tab instead

---

## 📦 Creating the Installer (After Build Works)

Once your add-in builds and runs:

1. **Add Setup Project** to solution
2. **Follow:** [INSTALLER_CREATION_GUIDE.md](INSTALLER_CREATION_GUIDE.md)
3. **Build MSI installer**
4. **Test install/uninstall**
5. **Prepare for Autodesk App Store**

---

## ✅ Success Checklist

Before moving to installer:

- [ ] Project builds without errors
- [ ] Inventor launches when debugging
- [ ] Add-in appears in Add-ins Manager
- [ ] Can load/unload add-in successfully
- [ ] Ribbon button appears in Assembly tab
- [ ] Main form opens when button clicked
- [ ] Assembly Cloner form opens
- [ ] Registry scanning works
- [ ] No runtime errors

---

## 🎯 Next Steps After Build

1. ✅ **Create Setup Project** in Visual Studio
2. ✅ **Build MSI installer**
3. ✅ **Test installation**
4. ✅ **Test uninstallation**
5. ✅ **Prepare store assets:**
   - 120x120 icon
   - Screenshots
   - App description
6. ✅ **Submit to Autodesk App Store**

---

**Current Status:** All core files created! Ready to import into Visual Studio and build.
