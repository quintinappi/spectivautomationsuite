# Inventor Add-In Deployment Locations

## Summary

This document summarizes the correct deployment locations for Inventor 2026 add-ins based on local system scan and Autodesk documentation.

## Local Inventor Installation

**Inventor 2026 Installation Path:**
```
C:\Program Files\Autodesk\Inventor 2026\Bin\Inventor.exe
```

**Key Folders:**
- `C:\Program Files\Autodesk\Inventor 2026\` - Main installation
- `C:\Program Files\Autodesk\Inventor 2026\Bin\Public Assemblies\` - API DLLs
- `C:\Program Files\Autodesk\Inventor 2026\Bin\` - Core binaries

## Add-In Deployment Locations

### 1. All-Users Location (RECOMMENDED - Currently Used)
```
C:\ProgramData\Autodesk\Inventor 2026\Addins\
```
- **Pros:** Available to all users, survives user profile deletion
- **Cons:** Requires Administrator privileges to deploy
- **Use case:** Production deployment, enterprise environments

### 2. Per-User Location
```
%APPDATA%\Autodesk\Inventor Addins\
```
- **Pros:** No admin privileges required
- **Cons:** Only available to current user
- **Use case:** Development, testing

### 3. Per-User Version-Specific Location
```
%APPDATA%\Autodesk\Inventor 2026\Addins\
```
- **Pros:** Version-specific, no admin required
- **Cons:** Only for current user
- **Use case:** Version-specific testing

## Current Project Deployment

**Existing Add-In:**
- Name: `AssemblyClonerAddIn`
- Location: `C:\ProgramData\Autodesk\Inventor 2026\Addins\`
- Files deployed:
  - `AssemblyClonerAddIn.dll`
  - `AssemblyClonerAddIn.addin`

**Deploy Script:** `InventorAddIn/DEPLOY_NOW.bat`
- Copies files to `C:\ProgramData\Autodesk\Inventor 2026\Addins\`
- Requires Administrator privileges

## API References

**Required DLLs (from Inventor installation):**
```
C:\Program Files\Autodesk\Inventor 2026\Bin\Public Assemblies\Autodesk.Inventor.Interop.dll
C:\Program Files\Autodesk\Inventor 2026\Bin\Autodesk.iLogic.Interfaces.dll
C:\Program Files\Autodesk\Inventor 2026\Bin\stdole.dll
```

## References

- [Autodesk Forum: Addin loading locations](https://forums.autodesk.com/t5/inventor-programming-forum/addin-loading-locations/td-p/5984471)
- [Autodesk Forum: Where does Inventor load its addins](https://forums.autodesk.com/t5/inventor-forum/where-does-inventor-load-it-s-addins-to-memory-registry-appdata/td-p/9709821)
