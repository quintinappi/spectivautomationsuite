# CLAUDE.md - AI Agent Context for Spectiv Production Scripts

**Last Updated:** January 20, 2026

## Overview

This workspace contains VBScript automation tools for Autodesk Inventor, primarily for:
- Part/assembly renaming with heritage naming conventions
- IDW drawing reference updates  
- Assembly cloning with full isolation

## Key Scripts

### Part_Renaming/Assembly_Cloner.vbs
**Purpose:** Clone an assembly with ALL sub-assemblies, parts, and IDW files to a new isolated location.

**Key Features:**
- Recursive sub-assembly collection and copying
- Heritage renaming (CLONE-001-PL1.ipt, CLONE-001-A1.ipt, etc.)
- In-memory reference updates using Inventor API
- **Recursive IDW processing** - updates ALL IDWs in ALL subfolders (Jan 20, 2026 fix)
- Generates STEP_1_MAPPING.txt for tracking

**Critical Functions:**
- `CollectPartsRecursively()` - Builds dictionary of all parts/sub-assemblies
- `CopyAllFiles()` - Copies with optional renaming
- `UpdateAllReferencesInMemory()` - Updates assembly references while docs are open
- `UpdateIDWReferencesWithInventor()` - Updates IDW drawing references
- `ScanIDWFilesForUpdate()` - Recursively collects ALL IDW files from folder tree

**Known Fix (Jan 20, 2026):**
The IDW update phase originally only scanned the root folder. Added `ScanIDWFilesForUpdate()` to recursively collect ALL IDW files before processing, ensuring sub-assembly drawings have their IPT references updated correctly.

### Part_Renaming/Assembly_Renamer.vbs
**Purpose:** Rename parts in an open assembly with heritage naming.

### IDW_Updates/Recursive_IDW_Updater.vbs
**Purpose:** Update IDW drawing references after renaming.

## Inventor API Patterns

### Opening Documents Silently
```vbscript
invApp.SilentOperation = True
invApp.FileOptions.SetFileResolveOption kSkipUnresolvedFiles  ' = 54275
```

### Updating Assembly References (In-Memory)
```vbscript
For Each ref In asmDoc.File.ReferencedFileDescriptors
    ref.ReplaceReference newPath
Next
```

### Updating IDW References
```vbscript
Set idwDoc = invApp.Documents.Open(idwPath, False)
For Each fd In idwDoc.File.ReferencedFileDescriptors
    fd.ReplaceReference newPath
Next
idwDoc.Save
idwDoc.Close
```

## Golden Rules

1. **Never hardcode file names** - Always scan folders dynamically
2. **Never assume naming conventions** - IDW names ≠ assembly names
3. **Never close documents during iteration** - Breaks parent assembly context
4. **Use recursive scanning** - Sub-assemblies have their own IDWs in subfolders

## Testing

Test assembly: Head Chute with 8 sub-assemblies (Top, Middle, Bottom, Launder, Lid-1, Lid-2, Support Beam-1, Design Accelerator)
- 100 files total (parts + assemblies + IDWs)
- 8 IDW files across root and subfolders
