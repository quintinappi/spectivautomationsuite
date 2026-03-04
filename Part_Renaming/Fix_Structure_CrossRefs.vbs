' Fix_Structure_CrossRefs.vbs
' Fixes cross-references between SSCR04 and SSCR06 Structure & Walkway folders
' Remaps wrong folder references to local equivalents

Option Explicit

Dim oInventor, oFSO
Dim sFolder1, sFolder2
Dim sLogPath, oLogFile
Dim nTotalFixed, nTotalIDWs

' Folder paths
sFolder1 = "C:\Users\Quintin\Documents\Spectiv\3. Working\22. DMS plant\21. SSCR04 - Primary Floats D&R Screen Station\000 Structure & Walkway"
sFolder2 = "C:\Users\Quintin\Documents\Spectiv\3. Working\22. DMS plant\23. SSCR06 - Secondary Floats D&R Screen Station\000 Structure & Walkway"

' Path components for replacement
Const SSCR04_PATH = "21. SSCR04 - Primary Floats D&R Screen Station"
Const SSCR06_PATH = "23. SSCR06 - Secondary Floats D&R Screen Station"

' Part name prefixes
Const SSCR04_PREFIX = "N1SCR04-730-"
Const SSCR06_PREFIX = "N1SCR06-000-"

Set oFSO = CreateObject("Scripting.FileSystemObject")

' Log file
sLogPath = oFSO.GetParentFolderName(WScript.ScriptFullName) & "\Structure_CrossRef_Fix_Log.txt"
Set oLogFile = oFSO.CreateTextFile(sLogPath, True)

oLogFile.WriteLine "Structure & Walkway Cross-Reference Fix Log"
oLogFile.WriteLine "Started: " & Now()
oLogFile.WriteLine String(70, "=")
oLogFile.WriteLine ""

' Connect to Inventor
WScript.Echo "Connecting to Inventor..."
On Error Resume Next
Set oInventor = GetObject(, "Inventor.Application")
If Err.Number <> 0 Then
    Err.Clear
    Set oInventor = CreateObject("Inventor.Application")
    oInventor.Visible = True
End If
On Error GoTo 0

If oInventor Is Nothing Then
    WScript.Echo "ERROR: Could not connect to Inventor!"
    oLogFile.WriteLine "ERROR: Could not connect to Inventor!"
    oLogFile.Close
    WScript.Quit 1
End If

oInventor.SilentOperation = True
WScript.Echo "Connected to Inventor (SilentOperation enabled)!"
oLogFile.WriteLine "Connected to Inventor"
oLogFile.WriteLine ""

nTotalFixed = 0
nTotalIDWs = 0

' Fix SSCR04 IDWs - remap SSCR06 refs to SSCR04
WScript.Echo ""
WScript.Echo "Fixing SSCR04 Structure IDWs..."
oLogFile.WriteLine "FIXING SSCR04 IDWs (SSCR06 refs -> SSCR04):"
oLogFile.WriteLine String(70, "-")
FixIDWsInFolder sFolder1, "SSCR06", SSCR06_PATH, SSCR04_PATH, SSCR06_PREFIX, SSCR04_PREFIX, oLogFile

' Fix SSCR06 IDWs - remap SSCR04 refs to SSCR06
WScript.Echo ""
WScript.Echo "Fixing SSCR06 Structure IDWs..."
oLogFile.WriteLine ""
oLogFile.WriteLine "FIXING SSCR06 IDWs (SSCR04 refs -> SSCR06):"
oLogFile.WriteLine String(70, "-")
FixIDWsInFolder sFolder2, "SSCR04", SSCR04_PATH, SSCR06_PATH, SSCR04_PREFIX, SSCR06_PREFIX, oLogFile

' Summary
oLogFile.WriteLine ""
oLogFile.WriteLine String(70, "=")
oLogFile.WriteLine "FIX COMPLETE"
oLogFile.WriteLine "Total IDWs modified: " & nTotalIDWs
oLogFile.WriteLine "Total references fixed: " & nTotalFixed
oLogFile.WriteLine "Finished: " & Now()
oLogFile.Close

WScript.Echo ""
WScript.Echo "=========================================="
WScript.Echo "FIX COMPLETE!"
WScript.Echo "IDWs modified: " & nTotalIDWs
WScript.Echo "References fixed: " & nTotalFixed
WScript.Echo "Log saved to: " & sLogPath
WScript.Echo "=========================================="

oInventor.SilentOperation = False
WScript.Echo "Done!"


Sub FixIDWsInFolder(sFolderPath, sWrongFolderName, sWrongPath, sCorrectPath, sWrongPrefix, sCorrectPrefix, oLog)
    Dim oFolder, oSubFolder, oFile
    Dim arrIDWs, i, nCount
    
    If Not oFSO.FolderExists(sFolderPath) Then
        WScript.Echo "  ERROR: Folder not found: " & sFolderPath
        oLog.WriteLine "  ERROR: Folder not found: " & sFolderPath
        Exit Sub
    End If
    
    ' Collect all IDW files recursively
    arrIDWs = Array()
    CollectIDWFiles sFolderPath, arrIDWs
    
    nCount = UBound(arrIDWs) + 1
    WScript.Echo "  Found " & nCount & " IDW files"
    
    For i = 0 To UBound(arrIDWs)
        WScript.Echo "  [" & (i + 1) & "/" & nCount & "] " & oFSO.GetFileName(arrIDWs(i))
        FixSingleIDW arrIDWs(i), sWrongFolderName, sWrongPath, sCorrectPath, sWrongPrefix, sCorrectPrefix, oLog
    Next
End Sub


Sub CollectIDWFiles(sFolderPath, ByRef arrFiles)
    Dim oFolder, oSubFolder, oFile
    Dim nSize
    
    Set oFolder = oFSO.GetFolder(sFolderPath)
    
    ' Add IDW files from this folder
    For Each oFile In oFolder.Files
        If LCase(oFSO.GetExtensionName(oFile.Name)) = "idw" Then
            nSize = UBound(arrFiles) + 1
            ReDim Preserve arrFiles(nSize)
            arrFiles(nSize) = oFile.Path
        End If
    Next
    
    ' Recurse into subfolders
    For Each oSubFolder In oFolder.SubFolders
        CollectIDWFiles oSubFolder.Path, arrFiles
    Next
End Sub


Sub FixSingleIDW(sIDWPath, sWrongFolderName, sWrongPath, sCorrectPath, sWrongPrefix, sCorrectPrefix, oLog)
    Dim oDoc, oRefDesc
    Dim sRefPath, sNewPath, sFileName, sNewFileName
    Dim nFixed, bModified
    
    On Error Resume Next
    
    ' Open the IDW
    Set oDoc = oInventor.Documents.Open(sIDWPath, False)
    If Err.Number <> 0 Then
        oLog.WriteLine "  ERROR opening: " & sIDWPath & " - " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    
    nFixed = 0
    bModified = False
    
    ' Check each referenced file
    For Each oRefDesc In oDoc.File.ReferencedFileDescriptors
        sRefPath = oRefDesc.FullFileName
        
        ' Check if this reference points to the wrong folder
        If InStr(1, sRefPath, sWrongFolderName, vbTextCompare) > 0 Then
            ' Build the new path
            sNewPath = Replace(sRefPath, sWrongPath, sCorrectPath, 1, -1, vbTextCompare)
            
            ' Also fix the part name prefix (N1SCR06-000- -> N1SCR04-730- or vice versa)
            sFileName = oFSO.GetFileName(sNewPath)
            If InStr(1, sFileName, sWrongPrefix, vbTextCompare) > 0 Then
                sNewFileName = Replace(sFileName, sWrongPrefix, sCorrectPrefix, 1, -1, vbTextCompare)
                sNewPath = oFSO.GetParentFolderName(sNewPath) & "\" & sNewFileName
            End If
            
            ' Check if the target file exists
            If oFSO.FileExists(sNewPath) Then
                On Error Resume Next
                oRefDesc.ReplaceReference sNewPath
                If Err.Number = 0 Then
                    nFixed = nFixed + 1
                    bModified = True
                    oLog.WriteLine "    FIXED: " & oFSO.GetFileName(sRefPath) & " -> " & oFSO.GetFileName(sNewPath)
                Else
                    oLog.WriteLine "    ERROR remapping: " & oFSO.GetFileName(sRefPath) & " - " & Err.Description
                    Err.Clear
                End If
                On Error GoTo 0
            Else
                oLog.WriteLine "    SKIP (target missing): " & sNewPath
            End If
        End If
    Next
    
    ' Save if modified
    If bModified Then
        On Error Resume Next
        oDoc.Save
        If Err.Number <> 0 Then
            oLog.WriteLine "  ERROR saving: " & sIDWPath & " - " & Err.Description
            Err.Clear
        Else
            oLog.WriteLine "  SAVED: " & oFSO.GetFileName(sIDWPath) & " (" & nFixed & " refs fixed)"
            nTotalIDWs = nTotalIDWs + 1
            nTotalFixed = nTotalFixed + nFixed
        End If
        On Error GoTo 0
    End If
    
    ' Close without saving again (already saved)
    oDoc.Close True
    Set oDoc = Nothing
End Sub
