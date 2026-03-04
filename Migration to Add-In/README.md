# Assembly Cloner Migration to VB.NET Add-In

## Project Overview

**Objective:** Port proven VBScript `Assembly_Cloner.vbs` to VB.NET Inventor Add-In for Autodesk App Store distribution.

**Status:** 🔄 In Progress - Step-by-step migration

---

## Source: VBScript (Proven Working)

**Location:** `../Part_Renaming/Assembly_Cloner.vbs`

**Functionality:**
1. Detect open assembly in Inventor
2. Ask for destination folder
3. Copy assembly + ALL sub-assemblies + parts to new folder
4. Update references to use local copies
5. Copy associated IDW drawings
6. Optionally rename parts with heritage naming
7. Generate STEP_1_MAPPING.txt for reference tracking

**File Stats:**
- 39 functions/subs
- ~2,200 lines of code
- Proven production-ready

---

## Target: VB.NET Add-In (To Be Built)

**Location:** `./src/`

**Requirements:**
- ✅ Autodesk App Store compliant
- ✅ Native Inventor ribbon integration
- ✅ Identical functionality to VBScript
- ✅ Professional Windows Forms UI
- ✅ Comprehensive error handling

---

## Migration Strategy: Step-by-Step

### Phase 1: Core Infrastructure (Current)

| Step | Task | Status |
|------|------|--------|
| 1.1 | Create project structure | ✅ Done |
| 1.2 | Create VB.NET class skeleton | ⏳ In Progress |
| 1.3 | Setup logging system | ⏳ Pending |
| 1.4 | Setup Inventor connection | ⏳ Pending |

### Phase 2: Core Functions (Next)

| Step | VBScript Function | VB.NET Method | Status |
|------|-------------------|---------------|--------|
| 2.1 | `DetectOpenAssembly()` | `ValidateActiveDocument()` | ⏳ Pending |
| 2.2 | `GetDestinationFolder()` | `GetDestinationFolder()` | ⏳ Pending |
| 2.3 | `GetPlantSectionNaming()` | `GetPrefixFromUser()` | ⏳ Pending |
| 2.4 | `CollectAllReferencedParts()` | `CollectReferencedParts()` | ⏳ Pending |

### Phase 3: File Operations

| Step | VBScript Function | VB.NET Method | Status |
|------|-------------------|---------------|--------|
| 3.1 | `CopyAllFiles()` | `CopyAllFiles()` | ⏳ Pending |
| 3.2 | `GroupPartsForRenaming()` | `GroupPartsForRenaming()` | ⏳ Pending |
| 3.3 | `ClassifyByDescription()` | `ClassifyByDescription()` | ⏳ Pending |

### Phase 4: Reference Updates

| Step | VBScript Function | VB.NET Method | Status |
|------|-------------------|---------------|--------|
| 4.1 | `UpdateInMemoryAssemblyReferences()` | `UpdateAssemblyReferences()` | ⏳ Pending |
| 4.2 | `UpdateReferencesRecursively()` | `UpdateReferencesRecursively()` | ⏳ Pending |
| 4.3 | `UpdateIDWReferences()` | `UpdateIDWReferences()` | ⏳ Pending |

### Phase 5: Registry & Mapping

| Step | VBScript Function | VB.NET Method | Status |
|------|-------------------|---------------|--------|
| 5.1 | `ScanRegistryForCounters()` | `RegistryManager.Scan()` | ⏳ Pending |
| 5.2 | `SaveCounterToRegistry()` | `RegistryManager.Save()` | ⏳ Pending |
| 5.3 | `WriteMappingFile()` | `WriteMappingFile()` | ⏳ Pending |

---

## Testing Strategy

### Functional Equivalence Test

For each migrated function:

```
┌─────────────────────────────────────────────────────────────┐
│  TEST: FunctionName()                                       │
├─────────────────────────────────────────────────────────────┤
│  1. Run VBScript version on test assembly                   │
│  2. Record outputs (log files, copied files)                │
│  3. Run VB.NET version on SAME assembly                     │
│  4. Compare outputs                                         │
│  5. ✓ PASS if identical, ✗ FAIL and debug                   │
└─────────────────────────────────────────────────────────────┘
```

### Test Assembly

Use a simple test assembly:
- `Structure.iam` (main assembly)
- 3 sub-assemblies (`Column-1.iam`, `Beam-1.iam`, `Plate-1.iam`)
- 10 parts total
- 5 IDW drawings

---

## Project Structure

```
Migration to Add-In/
├── README.md                        (This file)
├── MAPPING.md                       (Function mapping tracker)
├── PROGRESS.md                      (Step-by-step progress log)
├── src/                             (VB.NET source code)
│   ├── AssemblyCloner.vb           (Main class)
│   ├── AssemblyClonerForm.vb       (UI form)
│   ├── RegistryManager.vb          (Registry operations)
│   ├── Logger.vb                   (Logging system)
│   └── FileHelper.vb               (File utilities)
├── docs/                            (Documentation)
│   ├── FUNCTION_REFERENCE.md       (Complete function reference)
│   ├── VBSCRIPT_TO_VBNET.md        (Syntax conversion guide)
│   └── TESTING_GUIDE.md            (Testing procedures)
└── tests/                           (Test files)
    ├── Test_Assembly/
    │   ├── Structure.iam
    │   ├── Column-1.iam
    │   └── (test parts)
    └── Expected_Results/
        └── (baseline outputs)
```

---

## Key Design Decisions

### 1. Class Structure

```
AssemblyCloner (Main Class)
├── Properties
│   ├── InventorApplication
│   ├── Logger
│   ├── RegistryManager
│   └── CopiedFiles (Dictionary)
├── Methods
│   ├── Clone() (Main entry point)
│   ├── ValidateActiveDocument()
│   ├── GetDestinationFolder()
│   ├── CollectReferencedParts()
│   ├── CopyAllFiles()
│   ├── UpdateAssemblyReferences()
│   └── UpdateIDWReferences()
```

### 2. Error Handling

```vb
' VBScript
On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Then
    ' Handle error
End If

' VB.NET
Try
    invApp = GetObject(, "Inventor.Application")
Catch ex As Exception
    Logger.LogError("Failed to connect: " & ex.Message)
    Throw
End Try
```

### 3. Collections

```vb
' VBScript
Dim dict
Set dict = CreateObject("Scripting.Dictionary")
dict.Add "key", "value"

' VB.NET
Dim dict As New Dictionary(Of String, String)()
dict.Add("key", "value")
```

---

## Progress Tracking

**Current Step:** Phase 1 - Core Infrastructure

**Last Updated:** 2025-01-21

**Next Action:** Create AssemblyCloner.vb skeleton class

---

## References

- **Source:** `../Part_Renaming/Assembly_Cloner.vbs`
- **Inventor API:** Autodesk Inventor 2026 API Help
- **VB.NET Guide:** `docs/VBSCRIPT_TO_VBNET.md`
- **Testing:** `docs/TESTING_GUIDE.md`
