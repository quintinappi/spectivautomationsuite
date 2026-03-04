# Derived Parts Handling for Assembly Cloner

## Overview

Inventor's **Derived Parts** are parts that get their geometry from a "base component" file. When cloning assemblies that contain derived parts, the cloned parts still reference the **original** base files - creating external dependencies.

This documentation covers the detection and fixing of these external derived part references.

---

## The Problem

When you clone an assembly containing derived parts:

```
Original Assembly/
├── Assembly.iam
├── Plate1.ipt  ──derives from──> BaseSheet.ipt (external)
├── Plate2.ipt  ──derives from──> BaseSheet.ipt (external)
└── Plate3.ipt  ──derives from──> BaseSheet.ipt (external)

Cloned Assembly/
├── CLONE-001-Assembly.iam
├── CLONE-001-Plate1.ipt  ──still points to──> Original/BaseSheet.ipt ❌
├── CLONE-001-Plate2.ipt  ──still points to──> Original/BaseSheet.ipt ❌
└── CLONE-001-Plate3.ipt  ──still points to──> Original/BaseSheet.ipt ❌
```

The cloned assembly is **NOT self-contained** - it depends on external files.

---

## The Solution

Two scripts handle this:

### 1. Check_Derived_Parts.vbs (Detection)

**Purpose**: Scan the active assembly and report all derived parts with their base file locations.

**Usage**:
1. Open your assembly in Inventor
2. Run the script
3. Review the report showing:
   - Which parts are derived
   - Where their base files are located
   - Whether they're in SAME FOLDER or DIFFERENT folder tree

**Output**: `DerivedParts_Report.txt`

### 2. Fix_Derived_Parts.vbs (Repair)

**Purpose**: Make a cloned assembly fully self-contained by:
1. Detecting derived parts with EXTERNAL base files
2. Copying those base files into the assembly folder
3. Renaming with the clone prefix
4. Updating all derived parts to use the local copies

**Usage**:
1. Open your CLONED assembly in Inventor
2. Run the script
3. Enter the prefix (auto-detected from existing files)
4. Script copies base files, updates references, saves

**Output**: `Fix_Derived_Log.txt`

**Important**: Run multiple times if you have chained derivations (Part A derives from Part B, which derives from Part C).

---

## Inventor API Used

| Object | Property/Method | Purpose |
|--------|-----------------|---------|
| `PartComponentDefinition` | `.ReferenceComponents` | Access reference components collection |
| `ReferenceComponents` | `.DerivedPartComponents` | Get collection of derived part components |
| `DerivedPartComponent` | `.LinkedToFile` | Check if still linked (not broken) |
| `DerivedPartComponent` | `.ReferencedDocumentDescriptor` | Get the document descriptor |
| `DerivedPartComponent` | `.Replace(path, options)` | **Key method** - updates the base file reference |
| `DocumentDescriptor` | `.FullDocumentName` | Get full path to the base file |

---

## Chained Derivations

Derived parts can form chains:

```
BaseSheet.ipt (master geometry)
    └── DevelopmentSheet.ipt (derives from BaseSheet)
            ├── Plate1.ipt (derives from DevelopmentSheet)
            ├── Plate2.ipt (derives from DevelopmentSheet)
            └── Plate3.ipt (derives from DevelopmentSheet)
```

**Fix strategy**: Run the fixer multiple times until no external references remain.

Each run handles one level of the chain:
- Run 1: Fixes Plate1-3 → copies & links to local DevelopmentSheet
- Run 2: Fixes DevelopmentSheet → copies & links to local BaseSheet
- Run 3: No external references found - done!

---

## Integration Plan with Assembly_Cloner

### Recommended Approach: Post-Process Loop

After Assembly_Cloner completes its work:

```
1. Assembly_Cloner copies all files, updates references, saves
2. Opens the cloned assembly in Inventor
3. LOOP:
   a. Scan for external derived part base files
   b. If none found → EXIT LOOP (done!)
   c. Copy external base files with prefix
   d. Update derived references to local copies
   e. Save modified files
   f. REPEAT (for chained derivations)
4. Final validation
```

### Why Post-Process?

1. **Separation of concerns**: Core cloning logic stays clean
2. **Handles chains automatically**: Loop until clean
3. **Already proven**: Fix_Derived_Parts.vbs works
4. **Safe**: Original Assembly_Cloner functionality unchanged

### Implementation Options

**Option A**: Call Fix_Derived_Parts.vbs from Assembly_Cloner at the end
- Simple integration
- Two separate scripts

**Option B**: Merge the fix logic directly into Assembly_Cloner
- Single script
- More complex but self-contained

**Option C**: Create a wrapper script that runs both
- Cleanest separation
- Easy to maintain independently

---

## File Locations

```
Part_Renaming/
├── Assembly_Cloner.vbs          # Main cloning script
├── Check_Derived_Parts.vbs      # Detection/reporting
├── Fix_Derived_Parts.vbs        # Fixes external references
├── DerivedParts_Report.txt      # Check output
├── Fix_Derived_Log.txt          # Fix output
└── DERIVED_PARTS_HANDLING.md    # This documentation
```

---

## Testing Checklist

- [ ] Run Check on assembly with no derived parts (should report 0)
- [ ] Run Check on assembly with derived parts (should list all)
- [ ] Run Fix on cloned assembly (should copy & update)
- [ ] Run Check again (should show SAME FOLDER)
- [ ] Verify assembly opens without errors
- [ ] Verify derived parts update when base file changes

---

## Known Limitations

1. **Embedded derived parts**: If `LinkedToFile = False`, the link is broken and cannot be updated
2. **Very deep chains**: May need multiple fix passes
3. **Derived assemblies**: Currently handles `DerivedPartComponents` only, not `DerivedAssemblyComponents`

---

## Changelog

- **2026-01-15**: Initial creation of Check_Derived_Parts.vbs and Fix_Derived_Parts.vbs
- **2026-01-15**: Confirmed working on live assembly with chained derivations
