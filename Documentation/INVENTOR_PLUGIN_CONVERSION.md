# Converting VBScript Tools to Inventor Add-in

## Current Approach vs Inventor Add-in

| Aspect | Current (VBScript) | Inventor Add-in |
|--------|-------------------|-----------------|
| **Language** | VBScript (.vbs) | VB.NET or C# (.dll) |
| **Execution** | External via cscript | Runs inside Inventor |
| **UI** | MsgBox dialogs | Ribbon buttons, panels, custom dialogs |
| **Performance** | Good, but external COM | Faster, native .NET |
| **Debugging** | Limited | Full Visual Studio debugging |
| **Distribution** | Copy files | Installer or .addin file |

---

## Option 1: iLogic (Easiest - Quick Conversion)

iLogic is built into Inventor and runs VB.NET-like code directly.

### How to Use:
1. In Inventor: **Manage** tab → **iLogic** → **Add Rule**
2. Paste adapted VB.NET code (minor syntax changes from VBScript)
3. Can trigger from ribbon button or external event

### Pros:
- Minimal code changes required
- No Visual Studio needed
- Quick to implement

### Cons:
- Limited UI customization
- Rules stored per-document or externally
- Less professional distribution

---

## Option 2: Full VB.NET Add-in (Recommended for Production)

### Requirements:
- Visual Studio 2019/2022 (Community Edition is free)
- Inventor SDK (installed with Inventor)
- .NET Framework 4.8

### Project Structure:
```
SpectivInventorTools/
├── SpectivTools.vbproj
├── StandardAddInServer.vb          ' Entry point - implements AddInServer
├── Commands/
│   ├── AssemblyClonerCommand.vb    ' Assembly Cloner logic
│   ├── AssemblyRenamerCommand.vb   ' Part renaming logic
│   ├── IDWUpdaterCommand.vb        ' IDW reference updates
│   ├── DuplicateFinderCommand.vb   ' Duplicate file finder
│   └── TitleUpdaterCommand.vb      ' Title automation
├── UI/
│   ├── RibbonSetup.vb              ' Creates ribbon tab and buttons
│   └── Forms/
│       ├── ClonerOptionsForm.vb    ' Settings dialog for cloner
│       └── RenamerOptionsForm.vb   ' Settings dialog for renamer
├── Utilities/
│   ├── FileOperations.vb           ' Shared file copy/rename logic
│   ├── HeritageNaming.vb           ' Heritage naming classification
│   └── InventorHelpers.vb          ' Common Inventor API helpers
├── Autodesk.SpectivTools.Inventor.addin   ' Registration file
└── Resources/
    └── Icons/
        ├── Cloner16.png
        ├── Cloner32.png
        └── ...
```

### Key Files Explained:

#### StandardAddInServer.vb (Entry Point)
```vb.net
Imports Inventor

Public Class StandardAddInServer
    Implements Inventor.ApplicationAddInServer

    Private m_inventorApp As Inventor.Application
    Private m_ribbonPanel As RibbonPanel

    Public Sub Activate(addInSiteObject As ApplicationAddInSite, firstTime As Boolean) _
        Implements ApplicationAddInServer.Activate
        
        m_inventorApp = addInSiteObject.Application
        
        ' Create ribbon UI
        CreateRibbonUI()
    End Sub

    Public Sub Deactivate() Implements ApplicationAddInServer.Deactivate
        ' Cleanup
    End Sub

    Private Sub CreateRibbonUI()
        ' Add "Spectiv Tools" tab to ribbon
        ' Add buttons for each tool
    End Sub
End Class
```

#### .addin Registration File
```xml
<Addin Type="Standard">
    <ClassId>{GENERATE-A-GUID-HERE}</ClassId>
    <ClientId>{GENERATE-A-GUID-HERE}</ClientId>
    <DisplayName>Spectiv Tools</DisplayName>
    <Description>Assembly cloning, part renaming, and IDW tools</Description>
    <Assembly>SpectivTools.dll</Assembly>
    <LoadOnStartUp>1</LoadOnStartUp>
    <UserUnloadable>1</UserUnloadable>
    <Hidden>0</Hidden>
    <SupportedSoftwareVersionGreaterThan>24..</SupportedSoftwareVersionGreaterThan>
</Addin>
```

---

## Conversion Steps

### Step 1: Set Up Visual Studio Project
1. Install Visual Studio 2022 Community (free)
2. Create new "Class Library (.NET Framework)" project
3. Target .NET Framework 4.8
4. Add references:
   - `Autodesk.Inventor.Interop` (from Inventor install folder)
   - `System.Windows.Forms` (for dialogs)

### Step 2: Convert VBScript to VB.NET

Key syntax differences:

| VBScript | VB.NET |
|----------|--------|
| `Set obj = CreateObject("Inventor.Application")` | `Dim obj As Inventor.Application` (passed in) |
| `Dim x` | `Dim x As String` (typed) |
| `MsgBox "text"` | `MessageBox.Show("text")` |
| `On Error Resume Next` | `Try...Catch...End Try` |
| `CreateObject("Scripting.FileSystemObject")` | `System.IO.File`, `System.IO.Directory` |
| `WScript.Echo` | `Debug.WriteLine` or log file |

### Step 3: Create Ribbon UI
- Design icons (16x16 and 32x32 PNG)
- Create button definitions
- Wire up click handlers to command classes

### Step 4: Test and Debug
- Use Visual Studio debugger attached to Inventor
- Set breakpoints in your code
- Test each command

### Step 5: Deploy
- Copy `.dll` to: `%APPDATA%\Autodesk\ApplicationPlugins\SpectivTools\`
- Copy `.addin` file to: `%APPDATA%\Autodesk\Inventor 20XX\Addins\`
- Or create installer with Inno Setup / WiX

---

## Tools to Convert

| Current Script | Add-in Command | Priority |
|---------------|----------------|----------|
| `Assembly_Cloner.vbs` | AssemblyClonerCommand | High |
| `Assembly_Renamer.vbs` | AssemblyRenamerCommand | High |
| `IDW_Reference_Updater.vbs` | IDWUpdaterCommand | High |
| `Smart_Prefix_Scanner.vbs` | PrefixScannerCommand | Medium |
| `Title_Updater.vbs` | TitleUpdaterCommand | Medium |
| `Duplicate_File_Finder.vbs` | DuplicateFinderCommand | Low |

---

## Resources

- **Inventor API Documentation**: Help → Programming Help in Inventor
- **SDK Samples**: `C:\Users\Public\Documents\Autodesk\Inventor 20XX\SDK\`
- **Autodesk Forums**: https://forums.autodesk.com/t5/inventor-programming/bd-p/705
- **Brian Ekins Blog**: Excellent Inventor API tutorials

---

## Timeline Estimate

| Phase | Time |
|-------|------|
| Visual Studio setup | 1-2 hours |
| Convert first tool (Assembly Cloner) | 4-6 hours |
| Create ribbon UI | 2-3 hours |
| Convert remaining tools | 2-3 hours each |
| Testing and polish | 4-6 hours |
| **Total** | **~20-30 hours** |

---

## Quick Win Alternative: External Rules

If you want a middle ground, you can keep the VBScript files but create a simple add-in that just adds ribbon buttons that launch your existing .vbs files via `Shell.Run`. This gives you:
- Professional ribbon UI
- Existing code unchanged
- Quick to implement (~4-6 hours)

---

## Next Steps

1. Decide: Full conversion or ribbon launcher?
2. Install Visual Studio 2022 Community
3. Start with Assembly Cloner as pilot conversion
4. Expand to other tools once pattern is established
