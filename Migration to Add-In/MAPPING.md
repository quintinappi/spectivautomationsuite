# Function Mapping Tracker

## VBScript → VB.NET Function Mapping

**Purpose:** Track migration status of each function from `Assembly_Cloner.vbs` to VB.NET.

**Legend:**
- ⏸️ Not Started
- 🔄 In Progress
- ✅ Complete
- ⚠️ Blocked
- ❌ Failed

---

## Main Flow Functions (12)

| # | VBScript Function | VB.NET Method | Class | Status | Notes |
|---|-------------------|---------------|-------|--------|-------|
| 1 | `ASSEMBLY_CLONER_MAIN()` | `Clone()` | AssemblyCloner | ⏸️ | Entry point |
| 2 | `DetectOpenAssembly()` | `ValidateActiveDocument()` | AssemblyCloner | ⏸️ | Validate active .iam |
| 3 | `GetDestinationFolder()` | `GetDestinationFolder()` | AssemblyCloner | ⏸️ | Folder browser |
| 4 | `GetPlantSectionNaming()` | `GetPrefixFromUser()` | AssemblyCloner | ⏸️ | Get prefix |
| 5 | `CollectAllReferencedParts()` | `CollectReferencedParts()` | AssemblyCloner | ⏸️ | Scan hierarchy |
| 6 | `CollectIDWFiles()` | `CollectIDWFiles()` | AssemblyCloner | ⏸️ | Find .idw files |
| 7 | `GroupPartsForRenaming()` | `GroupPartsForRenaming()` | AssemblyCloner | ⏸️ | Classify parts |
| 8 | `CopyAllFiles()` | `CopyAllFiles()` | AssemblyCloner | ⏸️ | Copy files |
| 9 | `UpdateInMemoryAssemblyReferences()` | `UpdateAssemblyReferences()` | AssemblyCloner | ⏸️ | Update refs |
| 10 | `UpdateIDWReferences()` | `UpdateIDWReferences()` | AssemblyCloner | ⏸️ | Update IDWs |
| 11 | `WriteMappingFile()` | `WriteMappingFile()` | AssemblyCloner | ⏸️ | Save mapping |
| 12 | `ValidateCloneAndLog()` | `ValidateClone()` | AssemblyCloner | ⏸️ | Verify output |

---

## Helper Functions (27)

### Recursive Collection

| # | VBScript Function | VB.NET Method | Class | Status | Notes |
|---|-------------------|---------------|-------|--------|-------|
| 13 | `CollectPartsRecursively()` | `CollectPartsRecursively()` | AssemblyCloner | ⏸️ | Recursive scan |
| 14 | `CollectIDWFilesRecursive()` | `CollectIDWFilesRecursive()` | AssemblyCloner | ⏸️ | Recursive IDW scan |

### User Interaction

| # | VBScript Function | VB.NET Method | Class | Status | Notes |
|---|-------------------|---------------|-------|--------|-------|
| 15 | `GetUserNamingSchemes()` | `GetUserNamingSchemes()` | AssemblyClonerForm | ⏸️ | UI dialog |

### Classification

| # | VBScript Function | VB.NET Method | Class | Status | Notes |
|---|-------------------|---------------|-------|--------|-------|
| 16 | `ClassifyByDescription()` | `ClassifyByDescription()` | PartClassifier | ⏸️ | PL/B/CH/A/FL |
| 17 | `GetDescriptionFromIProperty()` | `GetDescriptionFromIProperty()` | PartClassifier | ⏸️ | Read iProps |

### File Operations

| # | VBScript Function | VB.NET Method | Class | Status | Notes |
|---|-------------------|---------------|-------|--------|-------|
| 18 | `ProcessIDWFilesWithReferenceUpdate()` | `ProcessIDWFiles()` | AssemblyCloner | ⏸️ | IDW handling |
| 19 | `IncrementFileName()` | `IncrementFileName()` | FileHelper | ⏸️ | Handle dupes |
| 20 | `GetFileNameFromPath()` | `GetFileNameFromPath()` | FileHelper | ⏸️ | Parse path |
| 21 | `CreateFolderRecursive()` | `CreateFolderRecursive()` | FileHelper | ⏸️ | Create folders |
| 22 | `BuildFileInventory()` | `BuildFileInventory()` | FileHelper | ⏸️ | Count files |

### Reference Updates (Alternate Methods)

| # | VBScript Function | VB.NET Method | Class | Status | Notes |
|---|-------------------|---------------|-------|--------|-------|
| 23 | `UpdateAssemblyReferencesWithApprentice()` | N/A | - | ⏸️ | NOT PORTING - using Inventor API |
| 24 | `UpdateAssemblyFileReferencesWithApprentice()` | N/A | - | ⏸️ | NOT PORTING - using Inventor API |
| 25 | `UpdateAssemblyFileReferences()` | N/A | - | ⏸️ | NOT PORTING - legacy method |
| 26 | `UpdateAssemblyReferencesWithInventor()` | N/A | - | ⏸️ | NOT PORTING - using InMemory method |
| 27 | `UpdateAssemblyReferences()` | N/A | - | ⏸️ | NOT PORTING - legacy method |
| 28 | `UpdateReferencesRecursively()` | `UpdateReferencesRecursively()` | AssemblyCloner | ⏸️ | Recursive ref update |

### Registry Operations

| # | VBScript Function | VB.NET Method | Class | Status | Notes |
|---|-------------------|---------------|-------|--------|-------|
| 29 | `ScanRegistryForCounters()` | `ScanCounters()` | RegistryManager | ⏸️ | Read registry |
| 30 | `SaveCounterToRegistry()` | `SaveCounter()` | RegistryManager | ⏸️ | Write registry |
| 31 | `CheckIfPrefixExistsInRegistry()` | `PrefixExists()` | RegistryManager | ⏸️ | Check prefix |

### Logging

| # | VBScript Function | VB.NET Method | Class | Status | Notes |
|---|-------------------|---------------|-------|--------|-------|
| 32 | `StartLogging()` | `Initialize()` | Logger | ⏸️ | Setup logging |
| 33 | `StopLogging()` | `Close()` | Logger | ⏸️ | Close log |
| 34 | `LogMessage()` | `Log()` | Logger | ⏸️ | Write log |
| 35 | `StartDestinationLogging()` | N/A | - | ⏸️ | NOT PORTING - use single log |

### Other

| # | VBScript Function | VB.NET Method | Class | Status | Notes |
|---|-------------------|---------------|-------|--------|-------|
| 36 | `ScanIDWFilesForUpdate()` | `ScanIDWFiles()` | AssemblyCloner | ⏸️ | Scan folder |
| 37 | `UpdateIPropertiesForCopiedDocuments()` | `UpdateIProperties()` | AssemblyCloner | ⏸️ | Update metadata |
| 38 | `UpdateIDWReferencesWithInventor()` | N/A | - | ⏸️ | MERGED into UpdateIDWReferences |

---

## Summary Statistics

```
Total Functions to Port: 38
├── Main Flow: 12
├── Recursive Collection: 2
├── User Interaction: 1
├── Classification: 2
├── File Operations: 5
├── Reference Updates: 6 (3 NOT PORTING)
├── Registry Operations: 3
├── Logging: 4 (1 NOT PORTING)
└── Other: 3 (1 MERGED)

Active Functions to Implement: 33
Skipped (Not Porting): 4
Merged: 1
```

---

## Progress by Phase

### Phase 1: Core Infrastructure
- Target: 5 functions
- Complete: 0
- In Progress: 0
- Pending: 5

### Phase 2: Core Functions
- Target: 4 functions
- Complete: 0
- In Progress: 0
- Pending: 4

### Phase 3: File Operations
- Target: 3 functions
- Complete: 0
- In Progress: 0
- Pending: 3

### Phase 4: Reference Updates
- Target: 3 functions
- Complete: 0
- In Progress: 0
- Pending: 3

### Phase 5: Registry & Mapping
- Target: 3 functions
- Complete: 0
- In Progress: 0
- Pending: 3

---

**Overall Progress: 0/38 functions (0%)**
