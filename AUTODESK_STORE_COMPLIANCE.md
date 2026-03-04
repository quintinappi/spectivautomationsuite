# Autodesk App Store Compliance - Solution Architecture

## 🚨 Critical Issue Identified

**HTA applications are NOT allowed** in the Autodesk App Store. The store requires **native Inventor Add-ins** (.dll files) built with C# or VB.NET that integrate directly with Inventor's API.

## ✅ Autodesk App Store Requirements

### Technical Requirements:
1. **Native Add-in (.dll)** - Must be a COM component
2. **Visual Studio Project** - C# or VB.NET
3. **Inventor API Integration** - Use Autodesk.Inventor.Interop library
4. **Installer** - Proper MSI or EXE installer
5. **Digital Signature** - Code signing certificate (recommended)
6. **Icon (120x120)** - Professional app icon
7. **Screenshots** - Up to 10 screenshots (2000x2000 max)
8. **Documentation** - Installation guide and help files

### Submission Requirements:
1. **Publisher Account** - Autodesk Single Sign-On
2. **App Description** - Up to 4000 characters
3. **Compatibility** - Specify supported Inventor versions
4. **Categories** - Select up to 4 relevant categories
5. **Support Information** - Contact details for support
6. **Price Type** - Free, Trial, Paid, or Subscription

## 🎯 New Solution Architecture

### Phase 1: C# Inventor Add-in (COM DLL)

**Technology Stack:**
- **Language:** C# (.NET Framework 4.8)
- **IDE:** Visual Studio 2019 or later
- **API:** Autodesk.Inventor.Interop (Inventor API)
- **UI:** Windows Forms or WPF
- **Installer:** WiX or Visual Studio Setup Project

**Project Structure:**
```
InventorAutomationSuite.sln
├── InventorAutomationSuiteAddIn/
│   ├── Properties/
│   │   └── AssemblyInfo.cs
│   ├── Resources/
│   │   ├── Icon.ico (120x120)
│   │   └── Screenshots/
│   ├── Forms/
│   │   ├── MainForm.cs (Unified Launcher)
│   │   ├── AssemblyClonerForm.cs
│   │   ├── PartRenamingForm.cs
│   │   └── ... (other tool forms)
│   ├── Core/
│   │   ├── InventorConnection.cs
│   │   ├── RegistryManager.cs
│   │   ├── Logger.cs
│   │   └── CommonFunctions.cs
│   ├── Tools/
│   │   ├── AssemblyCloner.cs
│   │   ├── PartRenaming.cs
│   │   ├── IDWUpdates.cs
│   │   └── ... (tool implementations)
│   ├── InventorAutomationSuiteAddIn.cs (Entry point)
│   └── app.config
├── InventorAutomationSuiteSetup/
│   └── (Installer project)
└── Documentation/
    ├── User Guide.pdf
    ├── Installation Guide.pdf
    └── Release Notes.txt
```

### Phase 2: Add-in Integration

**Add-in Manifest:**
```xml
<?xml version="1.0" encoding="utf-8"?>
<Addin xmlns="http://schemas.autodesk.com/inventor/addin">
  <ClassId>{GUID}</ClassId>
  <ClientId>{GUID}</ClientId>
  <DisplayName>Inventor Automation Suite</DisplayName>
  <Description>Professional automation tools for Autodesk Inventor</Description>
  <Assembly>InventorAutomationSuiteAddIn.dll</Assembly>
  <FullClassName>InventorAutomationSuiteAddIn.StandardAddInServer</FullClassName>
  <Versions>
    <Version>
      <Software>Inventor</Software>
      <Min>24.0</Min> <!-- Inventor 2025 -->
      <Max>25.0</Max>
    </Version>
  </Versions>
  <LoadOnStart>1</LoadOnStart>
</Addin>
```

## 🔧 Implementation Plan

### Step 1: Create C# Project Structure

**File:** `InventorAutomationSuiteAddIn.cs`

```csharp
using System;
using System.AddIn;
using Inventor;
using Autodesk.Inventor.Interop;

namespace InventorAutomationSuiteAddIn
{
    /// <summary>
    /// Inventor Automation Suite - Add-in Server
    /// </summary>
    public class StandardAddInServer : ApplicationAddInServer
    {
        private InventorApplication m_inventorApplication;
        private UserControl m_userControl;

        public void Activate(InventorApplication addInApplication, int firstTime)
        {
            m_inventorApplication = addInApplication;

            // Create main button in Inventor UI
            CreateButton();

            if (firstTime == 1)
            {
                // First-time initialization
                InitializeFirstTime();
            }
        }

        public void Deactivate()
        {
            // Cleanup
        }

        public void ExecuteCommand(int commandIndex)
        {
            // Show main form
            MainForm mainForm = new MainForm(m_inventorApplication);
            mainForm.Show();
        }

        private void CreateButton()
        {
            // Add button to Inventor ribbon
            // Implementation details...
        }
    }
}
```

### Step 2: Main Form (Unified Launcher)

**File:** `MainForm.cs`

```csharp
using System;
using System.Windows.Forms;
using Inventor;

namespace InventorAutomationSuiteAddIn
{
    public partial class MainForm : Form
    {
        private InventorApplication m_invApp;

        public MainForm(InventorApplication invApp)
        {
            m_invApp = invApp;
            InitializeComponent();
        }

        private void btnAssemblyCloner_Click(object sender, EventArgs e)
        {
            AssemblyClonerForm form = new AssemblyClonerForm(m_invApp);
            form.ShowDialog();
        }

        private void btnPartRenaming_Click(object sender, EventArgs e)
        {
            PartRenamingForm form = new PartRenamingForm(m_invApp);
            form.ShowDialog();
        }

        // ... other tool buttons
    }
}
```

### Step 3: Tool Forms

**File:** `AssemblyClonerForm.cs`

```csharp
using System;
using System.Windows.Forms;
using Inventor;

namespace InventorAutomationSuiteAddIn
{
    public partial class AssemblyClonerForm : Form
    {
        private InventorApplication m_invApp;
        private RegistryManager m_regManager;

        public AssemblyClonerForm(InventorApplication invApp)
        {
            m_invApp = invApp;
            m_regManager = new RegistryManager();
            InitializeComponent();
        }

        private void btnScanRegistry_Click(object sender, EventArgs e)
        {
            string prefix = txtPrefix.Text;
            var counters = m_regManager.ScanCounters(prefix);
            DisplayCounters(counters);
        }

        private void btnClone_Click(object sender, EventArgs e)
        {
            string prefix = txtPrefix.Text;
            int cloneCount = (int)numCloneCount.Value;

            AssemblyCloner cloner = new AssemblyCloner(m_invApp);
            cloner.Clone(prefix, cloneCount);
        }
    }
}
```

### Step 4: Core Functionality

**File:** `AssemblyCloner.cs`

```csharp
using System;
using Inventor;
using Microsoft.Win32;

namespace InventorAutomationSuiteAddIn
{
    public class AssemblyCloner
    {
        private InventorApplication m_invApp;

        public AssemblyCloner(InventorApplication invApp)
        {
            m_invApp = invApp;
        }

        public void Clone(string prefix, int count)
        {
            AssemblyDocument asmDoc = m_invApp.ActiveDocument as AssemblyDocument;

            for (int i = 1; i <= count; i++)
            {
                string clonePath = GenerateClonePath(asmDoc.FullFileName, i);
                asmDoc.SaveAs(clonePath, false);

                // Process parts in clone
                ProcessCloneAssembly(clonePath, prefix);
            }
        }

        private void ProcessCloneAssembly(string clonePath, string prefix)
        {
            // Open clone and update parts with new numbers
            // Implementation...
        }

        private string GenerateClonePath(string originalPath, int cloneNumber)
        {
            // Generate clone file path
            // Implementation...
            return "";
        }
    }
}
```

### Step 5: Registry Manager

**File:** `RegistryManager.cs`

```csharp
using System;
using Microsoft.Win32;
using System.Collections.Generic;

namespace InventorAutomationSuiteAddIn
{
    public class RegistryManager
    {
        private const string RegistryPath = @"HKEY_CURRENT_USER\Software\InventorRenamer\";

        public Dictionary<string, int> ScanCounters(string prefix)
        {
            Dictionary<string, int> counters = new Dictionary<string, int>();

            string[] groups = { "PL", "B", "CH", "A", "FL" };

            foreach (string group in groups)
            {
                string key = RegistryPath + prefix + group;
                int value = (int)Registry.GetValue(key, 0);
                counters[group] = value;
            }

            return counters;
        }

        public void UpdateCounter(string prefix, string group, int value)
        {
            string key = RegistryPath + prefix + group;
            Registry.SetValue(key, value, RegistryValueKind.DWord);
        }
    }
}
```

## 📦 Deployment Package

### Installer Requirements:
1. **MSI Installer** - Windows Installer package
2. **Registry Keys** - Add-in registration
3. **File Locations** - Program Files folder
4. **Uninstaller** - Clean removal
5. **Digital Signature** - Code signing certificate

### WiX Installer Example:
```xml
<?xml version="1.0" encoding="UTF-8"?>
<Wix xmlns="http://schemas.microsoft.com/wix/2006/wi">
  <Product Id="*"
           Name="Inventor Automation Suite"
           Language="1033"
           Version="1.0.0.0"
           Manufacturer="Spectiv Solutions"
           UpgradeCode="GUID">

    <Package InstallerVersion="200" Compressed="yes" />

    <Directory Id="TARGETDIR" Name="SourceDir">
      <Directory Id="ProgramFilesFolder">
        <Directory Id="INSTALLFOLDER" Name="InventorAutomationSuite" />
      </Directory>
    </Directory>

    <ComponentGroup Id="ProductComponents" Directory="INSTALLFOLDER">
      <Component Id="MainDLL">
        <File Id="AddInDLL" Source="InventorAutomationSuiteAddIn.dll" />
      </Component>
    </ComponentGroup>

  </Product>
</Wix>
```

## 📊 Store Submission Checklist

- [ ] Create publisher account on Autodesk App Store
- [ ] Build C# add-in DLL
- [ ] Create MSI installer
- [ ] Design 120x120 icon
- [ ] Create up to 10 screenshots (2000x2000 max)
- [ ] Write app description (up to 4000 characters)
- [ ] Specify compatibility (Inventor 2025+)
- [ ] Select categories (Automation, Productivity, etc.)
- [ ] Provide support information
- [ ] Set price type (Free/Trial/Paid/Subscription)
- [ ] Test installation and functionality
- [ ] Submit for review

## 🎯 Next Steps

1. ✅ **COMPLETED:** Research Autodesk Store requirements
2. ⏳ **NEXT:** Create Visual Studio C# project
3. ⏳ Implement add-in architecture
4. ⏳ Migrate VBScript logic to C#
5. ⏳ Create Windows Forms UI
6. ⏳ Build MSI installer
7. ⏳ Create icons and screenshots
8. ⏳ Test and submit to store

## 🚀 Benefits of Proper Add-in Architecture

1. **Store Compliance** - Meets all Autodesk requirements
2. **Better Integration** - Native Inventor UI integration
3. **Professional Appearance** - Modern Windows Forms/WPF UI
4. **Better Performance** - Compiled C# vs interpreted VBScript
5. **Easier Maintenance** - Visual Studio IDE and debugging
6. **Commercial Viability** - Can be sold on Autodesk App Store
7. **Update Management** - Easy updates through store
8. **Customer Reach** - Global visibility on Autodesk platform

---

**Conclusion:** The HTA approach must be abandoned in favor of a proper C# Inventor Add-in for Autodesk Store compliance.
