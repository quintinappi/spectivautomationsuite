# Q Smart Tools - Plugin Logic Documentation

## Assembly Cloner - Silent Reference Replacement Flow

**Date:** December 10, 2025
**Author:** Quintin de Bruin

---

## Key Discovery: Why Option 1 Works Silently

The critical difference between working (silent) and broken (dialog prompts) flows:

### BROKEN Flow (shows "where is file?" dialogs):
1. Close the assembly
2. Copy files to destination
3. Open copied assembly with ApprenticeServer
4. Try to update references -> **FAILS** because ApprenticeServer can't resolve original paths

### WORKING Flow (silent, no dialogs):
1. **Keep assembly OPEN in Inventor**
2. Copy parts to destination
3. Update references using `occ.Replace()` on the **OPEN** document
4. SaveAs assembly to new location
5. Process IDWs from SOURCE, update refs, SaveAs to DEST

---

## Assembly Cloner Flow (Clone + Rename)

```
STEP 1: Get user input
  - New assembly name
  - Destination folder

STEP 2: Collect part info (assembly STAYS OPEN)
  - Traverse all occurrences recursively
  - Build dictionary: originalPath -> occurrenceBaseName
  - Skip suppressed occurrences
  - Skip bolted connections

STEP 3: Copy parts to destination (assembly STILL OPEN)
  - For each part in partInfo:
    - Copy file to destFolder with new name (prefix_originalName.ipt)
    - Add to m_copiedFiles: originalPath -> newPath
    - Add to m_partNameMapping: occBaseName -> newOccBaseName

STEP 4: Update assembly references (assembly STILL OPEN)
  - Recursively traverse all occurrences
  - For each .ipt occurrence:
    - If m_copiedFiles contains originalPath:
      - occ.Replace(newPath, True)  <-- KEY: Works silently on OPEN doc
  - For each .iam sub-assembly:
    - Recurse into it
    - Save sub-assembly after updating

STEP 5: SaveAs assembly to new location
  - asmDoc.SaveAs(newAsmPath, False)

STEP 6: Process IDW files
  - SilentOperation = True
  - For each IDW in source folder:
    - Open ORIGINAL IDW from SOURCE (has valid refs)
    - Update references using fd.ReplaceReference(newPath)
    - SaveAs to destination with new name
    - Close IDW
  - SilentOperation = False

STEP 7: Patch iLogic rules
  - Find all iLogic rules in assembly and parts
  - Replace old occurrence names with new names
  - Pattern: "OldName:1" -> "NewName:1"

STEP 8: Final save
```

---

## Key Technical Points

### 1. occ.Replace() on OPEN Document
```vb
' This works SILENTLY because Inventor already has the document loaded
occ.Replace(newPath, True)
```

### 2. IDW Processing - Open from SOURCE
```vb
' Open from SOURCE (original location) - refs are valid
Dim idwDoc = m_inventorApp.Documents.Open(sourceIdwPath, False)

' Update refs to point to NEW location
fd.ReplaceReference(newPath)

' SaveAs to DESTINATION
idwDoc.SaveAs(destIdwPath, False)
```

### 3. SilentOperation Flag
```vb
m_inventorApp.SilentOperation = True  ' Suppress dialogs
' ... do IDW processing ...
m_inventorApp.SilentOperation = False ' Restore
```

### 4. Recursive Assembly Traversal
```vb
For Each occ As ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
    If occ.Suppressed Then Continue For

    Dim doc = occ.Definition.Document
    Dim fullPath = doc.FullFileName

    If fullPath.EndsWith(".ipt") Then
        ' Process part
        occ.Replace(newPath, True)
    ElseIf fullPath.EndsWith(".iam") Then
        ' Recurse into sub-assembly
        UpdateReferencesRecursively(doc)
        doc.Save()
    End If
Next
```

---

## Part Renamer Flow (Rename in Place - No Copy)

This is similar to Assembly Cloner but:
- Does NOT copy files
- Renames parts using SaveAs with heritage method
- Updates assembly references to new names
- Updates IDW references

```
STEP 1: Get user input
  - Project prefix (e.g., "PLANT1-000-")
  - Confirm component groups

STEP 2: Collect and group parts (assembly STAYS OPEN)
  - Read Description iProperty
  - Classify into groups (PL, B, CH, A, etc.)

STEP 3: Create heritage copies (SaveAs in SAME directory)
  - partDoc.SaveAs(newPath, True)  <-- True = create copy
  - This creates renamed files alongside originals

STEP 4: Update assembly references (assembly STILL OPEN)
  - occ.Replace(newPath, True)

STEP 5: Save assembly

STEP 6: Update IDW references
  - Open IDW
  - fd.ReplaceReference(newPath)
  - Save IDW

STEP 7: (Optional) Delete original files
```

---

## Classification Logic (from Client Requirements)

| Group | Description Pattern | Code |
|-------|-------------------|------|
| B | UB, UC, or IPE prefix | Beams/Columns (I and H sections) |
| PL | PL + S355JR | Platework |
| LPL | PL + other material | Liners |
| A | L + dimensions | Angles |
| CH | PFC or TFC | Channels |
| P | CHS | Circular hollow |
| SQ | SHS | Square hollow |
| FL | FL (not FLOOR) | Flatbar |

---

## Common Pitfalls to Avoid

1. **NEVER close the assembly before updating references**
   - ApprenticeServer can't resolve paths properly

2. **NEVER copy IDW then try to open it**
   - Causes "non-unique project file names" dialog
   - Instead: Open from SOURCE, update refs, SaveAs to DEST

3. **ALWAYS use full paths for reference mapping**
   - Filename-only matching fails when same name exists in different folders

4. **ALWAYS save sub-assemblies after updating their references**
   - Changes won't persist otherwise

5. **Use SilentOperation for IDW processing**
   - Prevents unwanted dialogs
