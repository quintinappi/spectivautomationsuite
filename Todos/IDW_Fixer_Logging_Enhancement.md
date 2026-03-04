# TODO: Enhanced IDW Fixer Logging for IAM vs IPT Debugging

## Status: PENDING

## Objective
Add enhanced logging to `IDW_Reference_Updater.vbs` to clearly distinguish between IAM (assembly) and IPT (part) reference updates, making it immediately obvious why IPTs might be failing when IAMs are succeeding.

## Key Principle
**DO NOT BREAK EXISTING LOGIC** - Only enhance logging. The script is partially working (IAMs update successfully), so we preserve all existing functionality.

---

## Changes Required

### File: `IDW_Reference_Updater.vbs`

#### 1. Add Global Counter Variables (after line 16)
```vbscript
Dim g_IAM_Updates, g_IAM_Errors, g_IAM_Skipped
Dim g_IPT_Updates, g_IPT_Errors, g_IPT_Skipped
```

#### 2. Add File Type Detection Function (new function after line 531)
```vbscript
Function GetFileTypeExtension(filePath)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileTypeExtension = UCase(fso.GetExtensionName(filePath))
End Function

Function GetFileTypeName(ext)
    Select Case ext
        Case "IAM"
            GetFileTypeName = "[ASSEMBLY]"
        Case "IPT"
            GetFileTypeName = "[PART]"
        Case Else
            GetFileTypeName = "[" & ext & "]"
    End Select
End Function
```

#### 3. Enhance Reference Processing Loop (lines 388-506)

**Current logging (line 398-399):**
```vbscript
LogMessage "IDW:   Processing reference: " & currentFileName
LogMessage "IDW:   Full path from IDW: " & currentFullPath
```

**Enhanced logging:**
```vbscript
Dim fileTypeExt
fileTypeExt = GetFileTypeExtension(currentFullPath)
Dim fileTypeName
fileTypeName = GetFileTypeName(fileTypeExt)

LogMessage "IDW:   Processing reference " & fileTypeName & ": " & currentFileName
LogMessage "IDW:   Full path from IDW: " & currentFullPath
```

#### 4. Update Success/Failure Logging (lines 471-505)

**When mapping found and updated successfully (after line 491):**
```vbscript
If Err.Number = 0 Then
    LogMessage "IDW:     ✓ SUCCESS - Reference updated using Design Assistant method"
    updateCount = updateCount + 1

    ' Track by file type
    If fileTypeExt = "IAM" Then
        g_IAM_Updates = g_IAM_Updates + 1
    ElseIf fileTypeExt = "IPT" Then
        g_IPT_Updates = g_IPT_Updates + 1
    End If
```

**When update fails (after line 494):**
```vbscript
Else
    LogMessage "IDW:     ✗ ERROR - ReplaceReference failed: " & Err.Description
    errorCount = errorCount + 1

    ' Track by file type
    If fileTypeExt = "IAM" Then
        g_IAM_Errors = g_IAM_Errors + 1
    ElseIf fileTypeExt = "IPT" Then
        g_IPT_Errors = g_IPT_Errors + 1
    End If
    Err.Clear
End If
```

**When file doesn't exist (after line 499):**
```vbscript
Else
    LogMessage "IDW:     ✗ ERROR - New file doesn't exist: " & newPath
    errorCount = errorCount + 1

    ' Track by file type
    If fileTypeExt = "IAM" Then
        g_IAM_Errors = g_IAM_Errors + 1
    ElseIf fileTypeExt = "IPT" Then
        g_IPT_Errors = g_IPT_Errors + 1
    End If
End If
```

**When no mapping found (after line 504):**
```vbscript
Else
    LogMessage "IDW:     (No mapping found - keeping current reference)"

    ' Track skipped by file type
    If fileTypeExt = "IAM" Then
        g_IAM_Skipped = g_IAM_Skipped + 1
    ElseIf fileTypeExt = "IPT" Then
        g_IPT_Skipped = g_IPT_Skipped + 1
    End If
End If
```

#### 5. Initialize Counters (in `UpdateAllIDWFilesWithDesignAssistantMethod`, line 302)
```vbscript
Sub UpdateAllIDWFilesWithDesignAssistantMethod(invApp, idwFiles, ByRef totalUpdates, ByRef totalErrors)
    LogMessage "IDW: Processing " & idwFiles.Count & " IDW files with Design Assistant method"

    ' Initialize file type counters
    g_IAM_Updates = 0
    g_IAM_Errors = 0
    g_IAM_Skipped = 0
    g_IPT_Updates = 0
    g_IPT_Errors = 0
    g_IPT_Skipped = 0

    ' ... rest of existing code
```

#### 6. Add Summary by File Type (in `UpdateAllIDWFilesWithDesignAssistantMethod`, after line 329)
```vbscript
LogMessage ""
LogMessage "=== DESIGN ASSISTANT METHOD SUMMARY ==="
LogMessage "Total IDW files processed: " & idwFiles.Count
LogMessage "Total reference updates: " & totalUpdates
LogMessage "Total errors: " & totalErrors
LogMessage ""
LogMessage "=== SUMMARY BY FILE TYPE ==="
LogMessage "[ASSEMBLY] Updates: " & g_IAM_Updates & " | Errors: " & g_IAM_Errors & " | Skipped: " & g_IAM_Skipped
LogMessage "[PART]     Updates: " & g_IPT_Updates & " | Errors: " & g_IPT_Errors & " | Skipped: " & g_IPT_Skipped

' Highlight potential problems
If g_IAM_Updates > 0 And g_IPT_Updates = 0 And g_IPT_Skipped = 0 Then
    LogMessage ""
    LogMessage "⚠ WARNING: ASSEMBLIES updated successfully but NO PARTS updated!"
    LogMessage "⚠ This suggests PART references are not in the mapping file."
    LogMessage "⚠ Check that your STEP 1 mapping includes IPT files, not just IAM files."
ElseIf g_IAM_Updates > 0 And g_IPT_Errors > 0 Then
    LogMessage ""
    LogMessage "⚠ WARNING: ASSEMBLIES updated but PARTS had errors!"
    LogMessage "⚠ Review the logs above to see which method failed for IPT files."
End If
```

---

## Example Output After Enhancement

```
IDW: Processing [1/10] MGY-200-DRD-01-01-01.idw
IDW:   Processing reference [ASSEMBLY]: Panel-1.iam
IDW:   Full path from IDW: C:\...\Panel1\Panel-1.iam
IDW:   Method 1: Found exact path match in mapping
IDW:     ✓ SUCCESS - Reference updated using Design Assistant method

IDW:   Processing reference [PART]: N1SCR04-730-B1.ipt
IDW:   Full path from IDW: C:\...\Panel1\N1SCR04-730-B1.ipt
IDW:   Method 1: Found exact path match in mapping
IDW:     (No mapping found - keeping current reference)

=== SUMMARY BY FILE TYPE ===
[ASSEMBLY] Updates: 15 | Errors: 0 | Skipped: 0
[PART]     Updates: 0 | Errors: 0 | Skipped: 50

⚠ WARNING: ASSEMBLIES updated successfully but NO PARTS updated!
⚠ This suggests PART references are not in the mapping file.
```

---

## Files to Modify

- `C:\Users\Quintin\Documents\Spectiv\3. Working\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\IDW_Updates\IDW_Reference_Updater.vbs`

## Lines to Modify

- After line 16: Add global counter variables
- After line 531: Add helper functions
- Lines 388-506: Enhance logging with file type
- Line 302: Initialize counters
- After line 329: Add summary section

---

## Verification Steps

1. Run the enhanced script on a test IDW
2. Check that the log clearly shows `[ASSEMBLY]` vs `[PART]` labels
3. Verify the summary shows counts by file type
4. Confirm warning messages appear when appropriate
5. **Most important**: Verify IAM updates still work (regression test)
