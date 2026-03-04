# Visual Studio Setup Guide - Inventor Automation Suite Add-in

## 🚨 Common Mistakes & Inventor 2026 Issues

**Issue 1: System.Windows.Forms .tlb Error**
```
A reference to 'System.Windows.Forms' could not be added.
The ActiveX type library '...\System.Windows.Forms.tlb' was exported
from a .NET assembly and cannot be added as a reference.
```
**You made this mistake:**
- You were in the **"Type Libraries"** tab or clicked a `.tlb` file
- **SOLUTION:** Go to **"Assemblies"** tab instead (left side of Reference Manager)
- Look for `System.Windows.Forms.dll` (NOT `System.Windows.Forms.tlb`)

**Issue 2: C# vs VB.NET Project**
- Created a **C# Class Library** instead of **Visual Basic Class Library**
- **SOLUTION:** Start over, choose "Class Library (.NET Framework)" with the **purple VB icon**

**Issue 3: Inventor 2026 COM Reference Missing (NEW!)**
- "Autodesk Inventor Object Library" does NOT appear in COM tab
- This is a **known issue** with Inventor 2026 - the folder structure changed
- **SOLUTION:** Use **Browse** tab → `C:\Program Files\Autodesk\Inventor 2026\Bin\Public Assemblies\`
- Select `Autodesk.Inventor.Interop.dll`
- Set **Embed Interop Types = False**

**RECOMMENDED: Use Your Existing Working Project**
- If you have `AssemblyClonerAddIn` that already works, modify that instead
- It already has correct references for Inventor 2026

---

## 🎯 Project Structure for VB.NET Inventor Add-in

### Step 1: Create New Project in Visual Studio

1. **Open Visual Studio 2026**
2. **Create New Project** (Ctrl+Shift+N)
3. **CRITICAL:** Search for and select **"Class Library (.NET Framework)"** with **Visual Basic** icon
   - ⚠️ **NOT** C# Class Library!
   - ⚠️ **NOT** just "Class Library" - must specify ".NET Framework"
   - Look for the **purple VB icon** (not the blue C# icon)
4. **Name:** `InventorAutomationSuiteAddIn`
5. **Framework:** .NET Framework 4.8
6. **Location:** Your preferred folder
7. **Click Create**

**How to verify you picked the right one:**
- Solution Explorer should show: `My Project` folder (VB.NET) instead of `Properties` folder (C#)
- Default file should be `Class1.vb` (not `Class1.cs`)

### Step 2: Add Required References

**⚠️ CRITICAL: For Inventor 2026, use direct DLL references, NOT COM references!**

**A. Add Inventor API Reference (REQUIRED) - Use Browse Method:**
1. **Right-click "References"** in Solution Explorer
2. **Select:** "Add Reference..." (or press Alt+R+A)
3. **Click "Browse" tab** (bottom left, NOT COM!)
4. **Navigate to:** `C:\Program Files\Autodesk\Inventor 2026\Bin\Public Assemblies\`
5. **Select:** `Autodesk.Inventor.Interop.dll`
6. **Click Add**
7. **After adding:** Right-click the reference → Properties → Set **"Embed Interop Types"** to **False**

**B. Additional Inventor References (Optional but Recommended):**
1. **Still in Reference Manager** → Browse tab
2. **Navigate to:** `C:\Program Files\Autodesk\Inventor 2026\Bin\`
3. **Select:** `stdole.dll`
4. **Click Add**
5. **Set Embed Interop Types** to **True** for stdole

**C. Add Windows Forms References (REQUIRED):**
1. **Right-click "References"** → "Add Reference..."
2. **Click "Assemblies" tab** on the left side
3. **Search box:** Type "System.Windows.Forms"
4. **Check:** "System.Windows.Forms" (Version 4.x.x)
5. **Search again:** Type "System.Drawing"
6. **Check:** "System.Drawing"
7. **Click OK**

⚠️ **CRITICAL DISTINCTION:**
- **Browse tab:** For Autodesk.Inventor.Interop.dll (direct file reference)
- **Assemblies tab:** For System.Windows.Forms, System.Drawing, etc.
- **DO NOT** use COM tab for Inventor 2026! (The COM type library is missing in 2026)

**Alternative: Use Your Existing Working Project**
If you have `AssemblyClonerAddIn` that already works, just modify that project instead of creating a new one.

### Step 3: Project File Structure

Create this folder structure in your project:

```
InventorAutomationSuiteAddIn/
├── My Project/
│   ├── AssemblyInfo.vb
│   └── Application.myapp
├── Forms/
│   ├── MainForm.vb
│   ├── AssemblyClonerForm.vb
│   ├── PartRenamingForm.vb
│   ├── IDWUpdatesForm.vb
│   ├── TitleAutomationForm.vb
│   ├── RegistryManagementForm.vb
│   ├── PrefixScannerForm.vb
│   ├── EmergencyFixerForm.vb
│   └── MappingProtectionForm.vb
├── Core/
│   ├── InventorConnection.vb
│   ├── RegistryManager.vb
│   ├── Logger.vb
│   └── CommonFunctions.vb
├── Tools/
│   ├── AssemblyCloner.vb
│   ├── PartRenaming.vb
│   ├── IDWUpdates.vb
│   ├── TitleAutomation.vb
│   └── ...
├── Resources/
│   └── (Images, icons, etc.)
├── InventorAutomationSuiteAddIn.vb (Entry point)
└── InventorAutomationSuiteAddIn.xml (Add-in manifest)
```

### Step 4: Create Add-in Server Class

**File: InventorAutomationSuiteAddIn.vb**

This is the main entry point that Inventor loads.

### Step 5: Create Add-in Manifest

**File: InventorAutomationSuiteAddIn.xml**

This XML file tells Inventor how to load and display your add-in.

### Step 6: Build Configuration

**Configuration Manager:**
- **Debug:** Test with Inventor
- **Release:** Production build
- **Platform:** x64 (Inventor is 64-bit only)
- **Output Path:** bin\Release\

### Step 7: Post-Build Event

Add this to Project Properties → Build Events → Post-Build Event:

```batch
copy "$(ProjectDir)InventorAutomationSuiteAddIn.xml" "$(TargetDir)"
```

This ensures the XML manifest is copied to the output folder.

---

## 🚀 Quick Start Instructions

### Option A: Using Inventor Add-in Template (EASIEST)

1. In Visual Studio, search for "Inventor Add-in" template
2. Select "Inventor 2025 Add-in" template
3. Name your project: `InventorAutomationSuiteAddIn`
4. The template will create all necessary files automatically

### Option B: Manual Class Library Setup

1. Create "Class Library (.NET Framework)" project
2. Add Inventor COM references
3. Create add-in server class manually
4. Create XML manifest manually
5. Add Windows Forms

---

## 📋 Next Steps After Setup

1. **Create Main Entry Point** - InventorAutomationSuiteAddIn.vb
2. **Create XML Manifest** - InventorAutomationSuiteAddIn.xml
3. **Create Main Form** - MainForm.vb with all tool buttons
4. **Implement Core Classes** - Registry manager, logger, etc.
5. **Implement Tool Forms** - Individual forms for each tool
6. **Build and Test** - Run in Debug mode with Inventor
7. **Create Installer** - MSI setup project

---

## 🔧 Development Settings

### Project Properties Configuration:

**Application:**
- Target Framework: .NET Framework 4.8
- Root Namespace: InventorAutomationSuiteAddIn
- Startup Object: (None)

**Compile:**
- Option Explicit: On
- Option Strict: On
- Option Infer: On

**Build:**
- Platform Target: x64 (CRITICAL!)
- Register for COM Interop: True

**Debug:**
- Start Action: Start external program
- External Program: `C:\Program Files\Autodesk\Inventor 2025\Bin\Inventor.exe`

---

## 📦 Building the Project

### Debug Build (Testing):
1. Set configuration to "Debug"
2. Press F5 or click "Start"
3. Visual Studio will launch Inventor automatically
4. Your add-in will load automatically
5. Set breakpoints and debug as needed

### Release Build (Production):
1. Set configuration to "Release"
2. Build → Build Solution
3. Output in: `bin\Release\`
4. Files:
   - `InventorAutomationSuiteAddIn.dll`
   - `InventorAutomationSuiteAddIn.xml`

---

## 🎯 Testing Your Add-in

1. **Build in Debug mode**
2. **Inventor launches automatically**
3. **Check Add-ins Manager** in Inventor:
   - Tools → Add-ins Manager
   - Find "Inventor Automation Suite"
   - Load/Unload buttons should work
4. **Test ribbon button** appears in specified tab
5. **Click button** to show main form

---

## 🚨 Common Issues & Solutions

### Issue: "Cannot load Inventor API"
**Solution:** Add reference to `Autodesk.Inventor.Interop.dll` via **Browse** tab (not COM!)

### Issue: "Autodesk Inventor Object Library missing in COM tab"
**Solution:** This is normal for Inventor 2026! Use Browse tab instead:
- Go to: `C:\Program Files\Autodesk\Inventor 2026\Bin\Public Assemblies\`
- Select: `Autodesk.Inventor.Interop.dll`
- Set: Embed Interop Types = False

### Issue: "Add-in doesn't appear in Inventor"
**Solution:** Check XML manifest file is in same folder as DLL

### Issue: "Platform mismatch"
**Solution:** Set Platform Target to x64 in project properties

### Issue: "COM Interop error"
**Solution:** Enable "Register for COM Interop" in Build settings

### Issue: "System.Windows.Forms .tlb error"
**Solution:** You clicked Type Libraries tab - use Assemblies tab instead

---

## 📞 Next File to Create

After setting up the project structure, I'll provide:
1. Main add-in server class (InventorAutomationSuiteAddIn.vb)
2. XML manifest file
3. Main form with all tool buttons
4. Individual tool forms
5. Core functionality classes

**STATUS: Ready to create Visual Basic .NET code files!**
