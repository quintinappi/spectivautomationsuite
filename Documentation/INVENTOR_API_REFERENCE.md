# INVENTOR API REFERENCE - Complete Working Solutions

**Author:** Quintin de Bruin © 2025
**Last Updated:** December 11, 2025

This document contains the EXACT working methodology for all Inventor automation tasks. **NEVER deviate from these methods.**

---

## Table of Contents

1. [File Operations (Copy, Rename, Replace)](#1-file-operations)
2. [Assembly Reference Updates](#2-assembly-reference-updates)
3. [IDW/Drawing Reference Updates](#3-idwdrawing-reference-updates)
4. [iLogic Rules (Read and Modify)](#4-ilogic-rules)
5. [iProperties Access](#5-iproperties-access)
6. [Mass Properties](#6-mass-properties)
7. [Complete Assembly Cloning Workflow](#7-complete-assembly-cloning-workflow)
8. [Part Cloning Workflow](#8-part-cloning-workflow)
9. [Critical Rules - NEVER BREAK](#9-critical-rules)

---

## 1. File Operations

### Copy Files (Parts, Assemblies, IDWs)

**ALWAYS use simple file copy - NOT SaveAs for cloning:**

```vbscript
' VBScript Method (PROVEN WORKING)
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CopyFile originalPath, newPath, True
```

```vb.net
' VB.NET Method (PROVEN WORKING)
System.IO.File.Copy(sourcePath, destPath, True)
```

**WHY:** Simple file copy preserves all file data. `SaveAs` creates heritage relationships which can cause issues.

### Rename Files

**Use file copy with new name - the file system handles renaming:**

```vbscript
' VBScript - Copy with new name
fso.CopyFile originalPath, destFolder & "\" & newFileName, True
```

```vb.net
' VB.NET - Copy with new name
Dim newPath As String = System.IO.Path.Combine(destFolder, newFileName)
System.IO.File.Copy(sourcePath, newPath, True)
```

---

## 2. Assembly Reference Updates

### Method: fd.ReplaceReference + Occurrence Renaming (PROVEN WORKING)

**CRITICAL: fd.ReplaceReference updates FILE PATH but NOT occurrence NAME!**

**Complete workflow:**

```vbscript
' VBScript Method (from Assembly_Cloner.vbs)
Dim invApp
Set invApp = GetObject(, "Inventor.Application")

' Open assembly
Dim asmDoc
Set asmDoc = invApp.Documents.Open(asmPath, False)  ' False = invisible

' Get file descriptors
Dim fileDescriptors
Set fileDescriptors = asmDoc.File.ReferencedFileDescriptors

' STEP 1: Update file references
Dim i
For i = 1 To fileDescriptors.Count
    Dim fd
    Set fd = fileDescriptors.Item(i)

    Dim refPath
    refPath = fd.FullFileName

    ' Check if we have mapping for this file
    If g_CopiedFiles.Exists(refPath) Then
        Dim newRefPath
        newRefPath = g_CopiedFiles.Item(refPath)

        ' CRITICAL: Use fd.ReplaceReference (updates FILE path only!)
        fd.ReplaceReference newRefPath
    End If
Next

' STEP 2: CRITICAL FIX - Rename occurrences to match new filenames
' fd.ReplaceReference does NOT rename occurrences automatically!
Dim occurrences
Set occurrences = asmDoc.ComponentDefinition.Occurrences

For i = 1 To occurrences.Count
    Dim occ
    Set occ = occurrences.Item(i)

    Dim refPath
    refPath = occ.ReferencedFileDescriptor.FullFileName

    Dim newFileName
    newFileName = Mid(refPath, InStrRev(refPath, "\") + 1)
    newFileName = Left(newFileName, Len(newFileName) - 4)  ' Remove .ipt

    Dim currentOccName
    currentOccName = occ.Name

    ' Extract base name (without :1, :2, etc)
    Dim colonPos
    colonPos = InStrRev(currentOccName, ":")

    Dim baseName, instanceNum
    If colonPos > 0 Then
        baseName = Left(currentOccName, colonPos - 1)
        instanceNum = Mid(currentOccName, colonPos)
    Else
        baseName = currentOccName
        instanceNum = ""
    End If

    ' Rename if needed
    If baseName <> newFileName Then
        occ.Name = newFileName & instanceNum
    End If
Next

' STEP 3: Update assembly to resolve changes
asmDoc.Update

' Save and close
asmDoc.Save
asmDoc.Close
```

```vb.net
' VB.NET Method (from AssemblyCloner.vb)
m_InventorApp.SilentOperation = True

Dim asmDoc As AssemblyDocument = CType(m_InventorApp.Documents.Open(asmPath, False), AssemblyDocument)

' STEP 1: Update file references (late binding for COM compatibility)
Dim fileDescriptors As Object = asmDoc.File.ReferencedFileDescriptors
Dim fdCount As Integer = fileDescriptors.Count

For i As Integer = 1 To fdCount
    Dim fd As Object = fileDescriptors.Item(i)
    Dim refPath As String = fd.FullFileName

    If m_CopiedFiles.ContainsKey(refPath) Then
        Dim newRefPath As String = m_CopiedFiles(refPath)
        fd.ReplaceReference(newRefPath)  ' Updates FILE path only!
    End If
Next

' STEP 2: CRITICAL - Rename occurrences to match new filenames
Dim occurrences As ComponentOccurrences = asmDoc.ComponentDefinition.Occurrences

For i As Integer = 1 To occurrences.Count
    Dim occ As ComponentOccurrence = occurrences.Item(i)
    Dim refPath As String = occ.ReferencedFileDescriptor.FullFileName
    Dim newFileName As String = System.IO.Path.GetFileNameWithoutExtension(refPath)

    ' Get current occurrence name and instance
    Dim currentOccName As String = occ.Name
    Dim colonPos As Integer = currentOccName.LastIndexOf(":"c)
    Dim baseName As String = currentOccName
    Dim instanceNum As String = ""

    If colonPos > 0 Then
        baseName = currentOccName.Substring(0, colonPos)
        instanceNum = currentOccName.Substring(colonPos)
    End If

    ' Rename if needed
    If baseName <> newFileName Then
        occ.Name = newFileName & instanceNum
    End If
Next

' STEP 3: Update assembly to resolve changes
asmDoc.Update()

' STEP 4: Patch iLogic (if needed) - do this in SAME session!
' Don't close and reopen - that causes issues

asmDoc.Save()
asmDoc.Close(False)
m_InventorApp.SilentOperation = False
```

### Why Occurrence Renaming is Critical

When you use `fd.ReplaceReference(newPath)`:
- ✅ File reference is updated: `C:\old\Part7.ipt` → `C:\new\staircasetest2-PL6.ipt`
- ❌ Occurrence name is NOT updated: still `Part7 Stair End Plate S-2:1`

This causes iLogic to fail because:
- iLogic code references: `Parameter("Part7 Stair End Plate S-2:1", ...)`
- After text replacement: `Parameter("staircasetest2-PL6:1", ...)`
- But occurrence is still named: `Part7 Stair End Plate S-2:1`
- Result: **Error - component not found!**

**Solution:** Manually rename occurrences to match their new filenames BEFORE patching iLogic.

---

## 3. IDW/Drawing Reference Updates

### Method: Open ORIGINAL, Update References, SaveAs to NEW Location

**CRITICAL: This is the ONLY way to avoid "Non-Unique Project File Names" dialogs:**

```vbscript
' VBScript Method (from Assembly_Cloner.vbs ProcessIDWFilesWithReferenceUpdate)

' Close all documents first
invApp.Documents.CloseAll

' Suppress dialogs
invApp.SilentOperation = True

' Open the ORIGINAL IDW (it has valid references to original parts)
Dim idwDoc
Set idwDoc = invApp.Documents.Open(originalIdwPath, False)

' Update references to point to NEW paths
Dim fileDescriptors
Set fileDescriptors = idwDoc.File.ReferencedFileDescriptors

Dim i
For i = 1 To fileDescriptors.Count
    Dim fd
    Set fd = fileDescriptors.Item(i)

    Dim refPath
    refPath = fd.FullFileName

    If g_CopiedFiles.Exists(refPath) Then
        Dim newRefPath
        newRefPath = g_CopiedFiles.Item(refPath)

        ' CRITICAL: fd.ReplaceReference
        fd.ReplaceReference newRefPath
    End If
Next

' Save to NEW location with SaveAs
idwDoc.SaveAs destIdwPath, False

idwDoc.Close

invApp.SilentOperation = False
```

```vb.net
' VB.NET Method (from AssemblyCloner.vb ProcessIDWFilesWithReferenceUpdate)

m_InventorApp.Documents.CloseAll()
m_InventorApp.SilentOperation = True

' Open ORIGINAL IDW
Dim idwDoc As DrawingDocument = CType(m_InventorApp.Documents.Open(idwPath, False), DrawingDocument)

' Update references
Dim fileDescriptors As ReferencedFileDescriptors = idwDoc.File.ReferencedFileDescriptors

For i As Integer = 1 To fileDescriptors.Count
    Dim fd As ReferencedFileDescriptor = fileDescriptors.Item(i)
    Dim refPath As String = fd.FullFileName

    If m_CopiedFiles.ContainsKey(refPath) Then
        Dim newRefPath As String = m_CopiedFiles(refPath)
        fd.ReplaceReference(newRefPath)  ' CRITICAL METHOD
    End If
Next

' SaveAs to NEW location
idwDoc.SaveAs(newIdwPath, False)
idwDoc.Close()

m_InventorApp.SilentOperation = False
```

---

## 4. iLogic Rules

### 4.1 Access iLogic Add-In

```vbscript
' VBScript
Dim iLogicAddIn
Set iLogicAddIn = invApp.ApplicationAddIns.ItemById("{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}")

If iLogicAddIn Is Nothing Then
    ' iLogic not available
    Exit Sub
End If

Dim iLogicAuto
Set iLogicAuto = iLogicAddIn.Automation
```

```vb.net
' VB.NET
Private m_iLogicAddIn As ApplicationAddIn
Private m_iLogicAuto As Object
Private Const ILOGIC_ADDIN_GUID As String = "{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}"

Public Function Initialize() As Boolean
    Try
        m_iLogicAddIn = m_InventorApp.ApplicationAddIns.ItemById(ILOGIC_ADDIN_GUID)
        If m_iLogicAddIn IsNot Nothing AndAlso m_iLogicAddIn.Activated Then
            m_iLogicAuto = m_iLogicAddIn.Automation
            Return True
        End If
    Catch
    End Try
    Return False
End Function
```

### 4.2 Read iLogic Rules

```vbscript
' VBScript - Get all rules from document
Dim rules
Set rules = iLogicAuto.Rules(doc)

Dim rule
For Each rule In rules
    Dim ruleName
    ruleName = rule.Name

    Dim isActive
    isActive = rule.IsActive

    Dim sourceCode
    sourceCode = rule.Text  ' FULL SOURCE CODE

    WScript.Echo "Rule: " & ruleName
    WScript.Echo "Active: " & isActive
    WScript.Echo "Code: " & sourceCode
Next
```

```vb.net
' VB.NET - Get all rules
Public Function GetRuleNames(doc As Document) As List(Of String)
    Dim ruleNames As New List(Of String)

    If m_iLogicAuto Is Nothing Then Return ruleNames

    Try
        Dim rules As Object = m_iLogicAuto.Rules(doc)
        For Each rule As Object In rules
            ruleNames.Add(rule.Name)
        Next
    Catch
    End Try

    Return ruleNames
End Function

Public Function GetRuleText(doc As Document, ruleName As String) As String
    Try
        Dim rules As Object = m_iLogicAuto.Rules(doc)
        For Each rule As Object In rules
            If rule.Name = ruleName Then
                Return rule.Text  ' FULL SOURCE CODE
            End If
        Next
    Catch
    End Try
    Return ""
End Function
```

### 4.3 Modify iLogic Rules (CRITICAL FOR RENAMING)

**The `rule.Text` property is READ/WRITE - you CAN modify rules:**

```vbscript
' VBScript - Modify rule text
Dim rules
Set rules = iLogicAuto.Rules(doc)

Dim rule
For Each rule In rules
    Dim originalText
    originalText = rule.Text

    Dim modifiedText
    modifiedText = originalText

    ' Replace old part names with new part names
    ' Example: "Part1 TFC" -> "DMS-STAIR-PL1"
    modifiedText = Replace(modifiedText, "Part1 TFC", "DMS-STAIR-PL1")
    modifiedText = Replace(modifiedText, "Part2 TFC", "DMS-STAIR-CH1")

    ' Only update if text changed
    If modifiedText <> originalText Then
        rule.Text = modifiedText  ' WRITE BACK MODIFIED TEXT
    End If
Next
```

```vb.net
' VB.NET - Modify rule text with mappings
Public Function PatchRules(doc As Document, replacements As Dictionary(Of String, String)) As Integer
    Dim patchedCount As Integer = 0

    If m_iLogicAuto Is Nothing Then Return 0

    Try
        Dim rules As Object = m_iLogicAuto.Rules(doc)

        For Each rule As Object In rules
            Dim originalText As String = rule.Text
            Dim modifiedText As String = originalText

            ' Apply all replacements
            For Each kvp As KeyValuePair(Of String, String) In replacements
                If modifiedText.Contains(kvp.Key) Then
                    modifiedText = modifiedText.Replace(kvp.Key, kvp.Value)
                End If
            Next

            ' Only update if text changed
            If modifiedText <> originalText Then
                rule.Text = modifiedText  ' WRITE BACK
                patchedCount += 1
            End If
        Next
    Catch
    End Try

    Return patchedCount
End Function
```

### 4.4 Run iLogic Rules

```vbscript
' VBScript
iLogicAuto.RunRule doc, "RuleName"
```

```vb.net
' VB.NET
Public Function RunRule(doc As Document, ruleName As String) As Boolean
    Try
        m_iLogicAuto.RunRule(doc, ruleName)
        Return True
    Catch
        Return False
    End Try
End Function
```

---

## 5. iProperties Access

### 5.1 Read iProperties

```vbscript
' VBScript - Read Design Tracking Properties
Dim propertySet
Set propertySet = doc.PropertySets.Item("Design Tracking Properties")

Dim descriptionProp
Set descriptionProp = propertySet.Item("Description")
Dim description
description = Trim(descriptionProp.Value)

Dim partNumberProp
Set partNumberProp = propertySet.Item("Part Number")
Dim partNumber
partNumber = partNumberProp.Value

Dim stockNumberProp
Set stockNumberProp = propertySet.Item("Stock Number")
Dim stockNumber
stockNumber = stockNumberProp.Value
```

```vb.net
' VB.NET - Read Design Tracking Properties
Public Function ReadDesignTrackingProperty(doc As Document, propertyName As String) As String
    Try
        Dim propSet As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
        Dim prop As [Property] = propSet.Item(propertyName)
        Return prop.Value.ToString().Trim()
    Catch
        Return ""
    End Try
End Function

' Usage:
Dim description As String = ReadDesignTrackingProperty(doc, "Description")
Dim partNumber As String = ReadDesignTrackingProperty(doc, "Part Number")
Dim stockNumber As String = ReadDesignTrackingProperty(doc, "Stock Number")
```

### 5.2 Property Set Names

- `"Inventor Summary Information"` - Title, Subject, Author, Keywords
- `"Design Tracking Properties"` - Part Number, Description, Stock Number, Mass, etc.
- `"Inventor Document Summary Information"` - Category, Company
- `"Inventor User Defined Properties"` - Custom user properties

---

## 6. Mass Properties

```vbscript
' VBScript
Dim massProps
Set massProps = doc.ComponentDefinition.MassProperties

Dim mass
mass = massProps.Mass  ' In internal units (kg)

Dim volume
volume = massProps.Volume  ' In cm³

Dim area
area = massProps.Area  ' In cm²
```

```vb.net
' VB.NET
Public Function GetMassKg(doc As Document) As Double
    Try
        If doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            Dim partDoc As PartDocument = CType(doc, PartDocument)
            Return partDoc.ComponentDefinition.MassProperties.Mass
        ElseIf doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)
            Return asmDoc.ComponentDefinition.MassProperties.Mass
        End If
    Catch
    End Try
    Return 0
End Function
```

---

## 7. Complete Assembly Cloning Workflow

**This is the EXACT sequence from AssemblyCloner.vb (Inventor Add-In):**

```
STEP 1: Detect open assembly
        - invApp.ActiveDocument
        - Verify it's .iam file

STEP 2: Get destination folder
        - Use folder browser dialog
        - Create if doesn't exist

STEP 3: Ask about heritage renaming (optional)
        - Get plant section prefix
        - e.g., "DMS-STAIR-", "staircasetest2"

STEP 4: Collect all referenced parts
        - Recursive traversal through sub-assemblies
        - Skip suppressed occurrences
        - Skip bolted connections
        - Store: fullPath -> description
        - CRITICAL: Capture occurrence names (without :1 suffix)

STEP 5: CLOSE source document before copying
        - sourceDoc.Close()
        - This prevents file locks

STEP 6: Copy assembly file
        - System.IO.File.Copy (NOT SaveAs!)
        - Store mapping: originalPath -> newPath

STEP 7: Copy all parts
        - System.IO.File.Copy (NOT SaveAs!)
        - If renaming: generate new names from description
        - Store file mappings: originalPath -> newPath
        - Store occurrence mappings: oldOccName -> newOccName (for iLogic)

STEP 8: Update assembly references AND patch iLogic (SAME SESSION!)
        - Open copied assembly invisibly

        STEP 8A: Update file references
        - Use fd.ReplaceReference for each file descriptor
        - Use late binding (Object) to avoid COM cast errors

        STEP 8B: Rename occurrences to match new filenames
        - CRITICAL: fd.ReplaceReference does NOT rename occurrences!
        - Manually rename: "Part7 Stair End Plate S-2:1" -> "staircasetest2-PL6:1"
        - Preserve instance numbers (:1, :2, etc.)

        STEP 8C: Update assembly
        - asmDoc.Update() to resolve all changes

        STEP 8D: Patch iLogic rules (in SAME session!)
        - For each rule: replace old occurrence names with new names
        - DON'T close and reopen - that causes reference resolution issues!

        STEP 8E: Save and close
        - Save once after ALL changes complete

STEP 9: Process IDW files
        - For each .idw in source folder:
          a) CloseAll documents first
          b) Open ORIGINAL IDW from source location (not copy!)
          c) Update references with fd.ReplaceReference
          d) Patch iLogic if needed
          e) SaveAs to destination location
          f) Close IDW

STEP 10: Save iLogic patch log
         - Document all occurrence mappings
         - Document all patched rules

STEP 11: Success message
```

### Critical Points in the Workflow

1. **COM Interface Casting**: Use late binding (`As Object`) for `ReferencedFileDescriptors` to avoid E_NOINTERFACE errors
2. **Occurrence Renaming**: Must manually rename occurrences after `fd.ReplaceReference` - it only updates file paths!
3. **Single Session**: Update references, rename occurrences, and patch iLogic in ONE session - don't close/reopen
4. **IDW Processing**: Open ORIGINAL IDW (with valid references), update, then SaveAs to new location

---

## 8. Part Cloning Workflow

**Simple part copy from Part_Cloner.vbs (Option 10):**

```
STEP 1: Detect open part
        - invApp.ActiveDocument
        - Verify it's .ipt file

STEP 2: Display part iProperties
        - Show Part Number, Description, Stock Number

STEP 3: Get destination folder
        - Use folder browser

STEP 4: Get new part name (optional)
        - InputBox with original name as default

STEP 5: Copy file
        - fso.CopyFile sourcePath, destPath, True

STEP 6: Success message
```

---

## 9. Critical Rules - NEVER BREAK

### File Operations
1. **ALWAYS use File.Copy / fso.CopyFile** for copying parts
2. **NEVER use SaveAs** for simple file duplication (use for heritage only when needed)
3. **CLOSE source documents** before copying to prevent locks

### Reference Updates
4. **ALWAYS use fd.ReplaceReference** for updating file references
5. **CRITICAL: fd.ReplaceReference updates FILE PATH only - NOT occurrence names!**
6. **ALWAYS rename occurrences** after fd.ReplaceReference to match new filenames
7. **Use late binding (Object)** for ReferencedFileDescriptors to avoid COM cast errors
8. **NEVER use filename-only matching** - always use FULL PATHS
9. **ALWAYS use SilentOperation = True** to suppress dialogs

### Occurrence Management
10. **fd.ReplaceReference does NOT rename occurrences** - must do manually
11. **Preserve instance numbers** when renaming (e.g., `:1`, `:2`)
12. **Update assembly after renaming** - call asmDoc.Update()

### IDW Processing
13. **Open ORIGINAL IDW**, update references, then **SaveAs to new location**
14. **NEVER copy IDW first** then try to update - causes dialog issues
15. **CloseAll documents** before opening IDW for processing
16. **Use late binding (Object)** for IDW FileDescriptors too

### iLogic
17. **iLogic uses occurrence names** (e.g., "Part1 TFC:1") not filenames
18. **Strip ":1" suffix** from occurrence names when building mappings
19. **rule.Text is READ/WRITE** - you CAN modify rules
20. **Patch iLogic in SAME session** as reference updates - don't close/reopen!

### Workflow
21. **Single session for assembly work**: Update refs → Rename occs → Update asm → Patch iLogic → Save
22. **NEVER close and reopen** assembly between steps - causes reference resolution issues
23. **Process IDWs separately** after assembly is complete

### Mappings
24. **NEVER hardcode file mappings** - always build dynamically
25. **Use full paths as dictionary keys** - filenames can collide
26. **Store both file mappings AND occurrence mappings** for iLogic

### Opening Documents
27. **Documents.Open(path, False)** - False = open invisibly
28. **ALWAYS close documents after processing** - don't leave them open
29. **Keep source assembly open for user** at the END (after all processing)

---

## API Object Reference

### Key Objects

| Object | Access | Purpose |
|--------|--------|---------|
| `Application` | `GetObject(, "Inventor.Application")` | Main Inventor app |
| `Document` | `invApp.ActiveDocument` | Current document |
| `AssemblyDocument` | Cast from Document | Assembly-specific |
| `DrawingDocument` | Cast from Document | IDW-specific |
| `ReferencedFileDescriptors` | `doc.File.ReferencedFileDescriptors` | All file references |
| `ReferencedFileDescriptor` | `.Item(i)` | Single file reference |
| `PropertySets` | `doc.PropertySets` | All property sets |
| `PropertySet` | `.Item("Design Tracking Properties")` | Single property set |
| `iLogicAuto` | Via AddIn GUID | iLogic automation |

### Key Methods

| Method | Object | Purpose |
|--------|--------|---------|
| `ReplaceReference(newPath)` | ReferencedFileDescriptor | Update file reference |
| `SaveAs(path, False)` | Document | Save to new location |
| `Save()` | Document | Save in place |
| `Close()` | Document | Close document |
| `Open(path, False)` | Documents | Open invisibly |
| `CloseAll()` | Documents | Close all documents |

### Key Properties

| Property | Object | Read/Write | Purpose |
|----------|--------|------------|---------|
| `FullFileName` | ReferencedFileDescriptor | Read | Full path to file |
| `Text` | iLogic Rule | **Read/Write** | Rule source code |
| `Name` | iLogic Rule | Read | Rule name |
| `IsActive` | iLogic Rule | Read | Rule active state |
| `SilentOperation` | Application | Read/Write | Suppress dialogs |

---

## File Locations

### VBScript Solutions (Option 9, Option 10)
```
FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\
├── Part_Renaming\
│   ├── Assembly_Cloner.vbs      <- Option 9 source
│   ├── Part_Cloner.vbs          <- Option 10 source
│   └── Assembly_Renamer.vbs     <- Option 1 source
```

### VB.NET Add-In
```
FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\
├── InventorAddIn\
│   └── AssemblyClonerAddIn\
│       ├── AssemblyCloner.vb        <- Main cloner logic
│       ├── iLogicPatcher.vb         <- iLogic patching
│       ├── DocumentInfoScanner.vb   <- iProperties/iLogic reader
│       └── StandardAddInServer.vb   <- Add-In entry point
```

---

**END OF REFERENCE DOCUMENT**

*This document represents the accumulated knowledge from months of development and debugging. Follow these methods EXACTLY.*
