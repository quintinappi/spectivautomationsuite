# MSI Installer Creation Guide - Inventor Automation Suite

## 🎯 Overview

The batch file deployers won't work for the Autodesk App Store. We need a proper **MSI Installer** that registers the add-in correctly with Inventor.

## 📦 Two Options for Creating the Installer

### Option A: Visual Studio Setup Project (EASIEST)

Built into Visual Studio, easier to use.

### Option B: WiX Toolset (MORE FLEXIBLE)

More powerful, but steeper learning curve.

---

## 🚀 Option A: Visual Studio Setup Project (Recommended)

### Step 1: Add Setup Project to Solution

1. **Right-click your Solution** in Solution Explorer
2. **Add → New Project**
3. Search for "Setup Project"
4. If not visible:
   - Click "Install more tools and features"
   - Select "Visual Studio Installer Projects Extension"
   - Install and restart Visual Studio
5. Name: `InventorAutomationSuiteSetup`
6. Click OK

### Step 2: Configure Setup Project

1. **In Setup Project, right-click → Add → Project Output**
2. **Select:** "Primary output" from "InventorAutomationSuiteAddIn"
3. **Click OK**

This automatically includes:
- Your compiled DLL
- All dependencies
- XML manifest file

### Step 3: Add XML Manifest to Installer

1. **Right-click Setup Project → Add → File**
2. **Browse to:** `InventorAutomationSuiteAddIn.xml`
3. **Click OK**

### Step 4: Configure Installation Settings

**Right-click Setup Project → Properties:**

- **ProductName:** Inventor Automation Suite
- **Manufacturer:** Spectiv Solutions
- **Title:** Inventor Automation Suite Setup
- **Version:** 1.0.0

### Step 5: Configure File System

1. **Double-click "File System"** in Setup Project
2. **Application Folder** should contain:
   - Primary output from InventorAutomationSuiteAddIn
   - InventorAutomationSuiteAddIn.xml

3. **Add Special Folders:**
   - Right-click → Add Special Folder → User's Application Data Folder
   - This creates a folder for log files

### Step 6: Create Registry Entries

1. **Right-click Setup Project → View → Registry**
2. **Navigate to:** `HKEY_CURRENT_USER\Software`
3. **Right-click → New Key** → Create keys:
   ```
   Software/
     └── SpectivSolutions/
         └── InventorAutomationSuite/
             └── InstallDate (String Value)
             └── Version (String Value)
   ```

4. **Add Values:**
   - Right-click → New String Value
   - Name: `InstallDate`, Value: `[Date]`
   - Name: `Version`, Value: `[ProductVersion]`

### Step 7: Build the Installer

1. **Right-click Setup Project → Build**
2. **Output:** `Setup.exe` and `InventorAutomationSuiteSetup.msi`
3. **Location:** Project folder under `bin\Debug\` or `bin\Release\`

### Step 8: Test the Installer

1. **Double-click `Setup.exe`** (or right-click → Install)
2. **Follow installation wizard**
3. **Launch Inventor**
4. **Check:** Tools → Add-ins Manager
5. **Your add-in should appear** in the list

---

## 🔧 Option B: WiX Toolset (Advanced)

### Step 1: Install WiX Toolset

1. **Download:** https://wixtoolset.org/releases/
2. **Install WiX Toolset Visual Studio Extension**
3. **Restart Visual Studio**

### Step 2: Create WiX Project

1. **Add → New Project**
2. **Select:** "Setup Project for WiX" or "WiX Project"
3. **Name:** `InventorAutomationSuiteWiXSetup`

### Step 3: Create Product.wxs

**File: Product.wxs**

```xml
<?xml version="1.0" encoding="UTF-8"?>
<Wix xmlns="http://schemas.microsoft.com/wix/2006/wi">
  <Product Id="*"
           Name="Inventor Automation Suite"
           Language="1033"
           Version="1.0.0.0"
           Manufacturer="Spectiv Solutions"
           UpgradeCode="PUT-GUID-HERE">

    <Package InstallerVersion="200" Compressed="yes" InstallScope="perUser" />

    <MajorUpgrade DowngradeErrorMessage="A newer version is already installed." />
    <MediaTemplate EmbedCab="yes" />

    <Feature Id="ProductFeature" Title="Inventor Automation Suite" Level="1">
      <ComponentGroupRef Id="ProductComponents" />
      <ComponentGroupRef Id="XMLManifest" />
    </Feature>

    <!-- Directory Structure -->
    <Directory Id="TARGETDIR" Name="SourceDir">
      <Directory Id="LocalAppDataFolder">
        <Directory Id="INSTALLFOLDER" Name="InventorAutomationSuite">
          <!-- Components go here -->
        </Directory>
      </Directory>
    </Directory>

    <!-- Components -->
    <ComponentGroup Id="ProductComponents" Directory="INSTALLFOLDER">
      <Component Id="MainDLL" Guid="PUT-GUID-HERE">
        <File Id="AddInDLL" Source="$(var.InventorAutomationSuiteAddIn.TargetPath)" />
      </Component>
    </ComponentGroup>

    <!-- XML Manifest -->
    <ComponentGroup Id="XMLManifest" Directory="INSTALLFOLDER">
      <Component Id="ManifestFile" Guid="PUT-GUID-HERE">
        <File Id="AddInManifest" Source="InventorAutomationSuiteAddIn.xml" />
      </Component>
    </ComponentGroup>

    <!-- Registry Entries -->
    <DirectoryRef Id="INSTALLFOLDER">
      <Component Id="RegistryEntries" Guid="PUT-GUID-HERE">
        <RegistryKey Root="HKCU" Key="Software\SpectivSolutions\InventorAutomationSuite">
          <RegistryValue Type="string" Name="InstallDate" Value="[Date]" />
          <RegistryValue Type="string" Name="Version" Value="[ProductVersion]" />
        </RegistryKey>
      </Component>
    </DirectoryRef>

  </Product>
</Wix>
```

### Step 4: Build WiX Installer

1. **Right-click WiX Project → Build**
2. **Output:** `InventorAutomationSuiteWiXSetup.msi`

---

## ✅ Testing Your Installer

### Test Checklist:

1. **Install on clean machine:**
   - Double-click MSI/Setup.exe
   - Follow wizard
   - Installation completes successfully

2. **Verify files installed:**
   - Navigate to installation folder
   - Check DLL and XML are present

3. **Test in Inventor:**
   - Launch Inventor
   - Tools → Add-ins Manager
   - Find "Inventor Automation Suite"
   - Click Load

4. **Test functionality:**
   - Click ribbon button
   - Main form appears
   - Assembly Cloner works

5. **Uninstall test:**
   - Control Panel → Programs and Features
   - Uninstall Inventor Automation Suite
   - Files removed
   - Registry cleaned up
   - Add-in removed from Inventor

---

## 🎯 Installer Best Practices

### DO:
✅ Use **per-user installation** (HKCU, not HKLM)
✅ Include **XML manifest** with DLL
✅ Create **uninstall shortcut**
✅ Set **digital signature** (if you have certificate)
✅ Test on **clean VM** before distribution
✅ Include **license agreement** dialog

### DON'T:
❌ Use **per-machine installation** (requires admin)
❌ Forget to include **XML manifest**
❌ Hardcode paths (use installer variables)
❌ Skip **uninstall testing**

---

## 📦 Autodesk App Store Submission

Once installer is tested:

1. **Create publisher account** on Autodesk App Store
2. **Build release version** of your add-in
3. **Build release version** of installer
4. **Test install/uninstall** thoroughly
5. **Gather required assets:**
   - 120x120 icon (PNG or ICO)
   - Up to 10 screenshots (2000x2000 max)
   - App description (4000 chars max)
   - Installation guide

6. **Submit to store:**
   - Upload MSI installer
   - Upload icon and screenshots
   - Fill in all required fields
   - Select compatibility (Inventor 2023, 2024, 2025)
   - Submit for review

---

## 🚨 Common Issues & Solutions

### Issue: "Add-in doesn't load"
**Solution:** Check XML manifest is in same folder as DLL

### Issue: "Registry permission error"
**Solution:** Use per-user installation (HKCU), not per-machine (HKLM)

### Issue: "Inventor can't find add-in"
**Solution:** Ensure Register for COM Interop is enabled in project settings

### Issue: "Installer won't build"
**Solution:** Install Visual Studio Installer Projects Extension

---

## 📞 Next Steps

1. ✅ **COMPLETED:** Create VB.NET add-in project
2. ✅ **COMPLETED:** Implement core functionality
3. ⏳ **NEXT:** Create Visual Studio Setup Project
4. ⏳ Build and test installer
5. ⏳ Prepare store assets (icon, screenshots)
6. ⏳ Submit to Autodesk App Store

---

**Current Status:** Ready to build installer in Visual Studio!
