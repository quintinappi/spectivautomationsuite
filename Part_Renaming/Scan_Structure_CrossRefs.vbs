' Scan_Structure_CrossRefs.vbs
' Scans ONLY the "000 Structure & Walkway" folders for cross-references
' Usage: cscript Scan_Structure_CrossRefs.vbs

Option Explicit

Dim oShell, oFSO, oInventor
Dim sFolder1, sFolder2
Dim sLogFile, oLogFile

' Set paths - ONLY Structure folders
sFolder1 = "C:\Users\Quintin\Documents\Spectiv\3. Working\22. DMS plant\21. SSCR04 - Primary Floats D&R Screen Station\000 Structure & Walkway"
sFolder2 = "C:\Users\Quintin\Documents\Spectiv\3. Working\22. DMS plant\23. SSCR06 - Secondary Floats D&R Screen Station\000 Structure & Walkway"

Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")

sLogFile = oFSO.GetParentFolderName(WScript.ScriptFullName) & "\Structure_CrossRef_Report.txt"

' Connect to Inventor
WScript.Echo "Connecting to Inventor..."
On Error Resume Next
Set oInventor = GetObject(, "Inventor.Application")
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Inventor must be running! Please start Inventor first."
    WScript.Quit 1
End If
On Error GoTo 0

' Enable silent operation to suppress dialogs
oInventor.SilentOperation = True

WScript.Echo "Connected to Inventor (SilentOperation enabled)!"

' Create log file
Set oLogFile = oFSO.CreateTextFile(sLogFile, True)
oLogFile.WriteLine "Structure & Walkway Cross-Reference Report"
oLogFile.WriteLine "Generated: " & Now()
oLogFile.WriteLine "=" & String(70, "=")
oLogFile.WriteLine ""

' Scan Folder 1 (SSCR04) IDWs - looking for refs to SSCR06
oLogFile.WriteLine "SSCR04 Structure IDWs with references pointing to SSCR06:"
oLogFile.WriteLine "-" & String(70, "-")
WScript.Echo vbCrLf & "Scanning SSCR04 Structure IDWs..."
ScanIDWsInFolder sFolder1, "SSCR06", oLogFile

' Scan Folder 2 (SSCR06) IDWs - looking for refs to SSCR04
oLogFile.WriteLine ""
oLogFile.WriteLine "SSCR06 Structure IDWs with references pointing to SSCR04:"
oLogFile.WriteLine "-" & String(70, "-")
WScript.Echo vbCrLf & "Scanning SSCR06 Structure IDWs..."
ScanIDWsInFolder sFolder2, "SSCR04", oLogFile

oLogFile.WriteLine ""
oLogFile.WriteLine "=" & String(70, "=")
oLogFile.WriteLine "SCAN COMPLETE"
oLogFile.Close

' Disable silent operation
oInventor.SilentOperation = False

WScript.Echo vbCrLf & "Report saved to: " & sLogFile
WScript.Echo "Done!"

' ============================================
' FUNCTIONS
' ============================================

Sub ScanIDWsInFolder(sFolder, sTargetFolderName, oLog)
    Dim dicIDWs, sKey
    Dim oDoc, sRefPath
    Dim nIDWCount, nCrossRefCount, nIDWsWithIssues
    
    Set dicIDWs = CreateObject("Scripting.Dictionary")
    CollectIDWFiles sFolder, dicIDWs
    
    WScript.Echo "  Found " & dicIDWs.Count & " IDW files"
    
    nIDWCount = 0
    nCrossRefCount = 0
    nIDWsWithIssues = 0
    
    For Each sKey In dicIDWs.Keys
        nIDWCount = nIDWCount + 1
        Dim bHasCrossRef, bFirstRef
        bHasCrossRef = False
        bFirstRef = True
        
        WScript.Echo "  [" & nIDWCount & "/" & dicIDWs.Count & "] " & oFSO.GetFileName(sKey)
        
        On Error Resume Next
        Set oDoc = oInventor.Documents.Open(sKey, False) ' Open invisible
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            oLog.WriteLine "  SKIPPED (open error): " & oFSO.GetFileName(sKey)
        Else
            On Error GoTo 0
            
            ' Check all referenced file descriptors
            Dim oRefFiles, oRefFile
            Set oRefFiles = oDoc.File.ReferencedFileDescriptors
            
            For Each oRefFile In oRefFiles
                sRefPath = ""
                On Error Resume Next
                sRefPath = oRefFile.FullFileName
                On Error GoTo 0
                
                If Len(sRefPath) > 0 Then
                    ' Check if this reference points to the OTHER folder
                    If InStr(1, sRefPath, sTargetFolderName, vbTextCompare) > 0 Then
                        If bFirstRef Then
                            oLog.WriteLine ""
                            oLog.WriteLine "  IDW: " & oFSO.GetFileName(sKey)
                            oLog.WriteLine "       Path: " & sKey
                            nIDWsWithIssues = nIDWsWithIssues + 1
                            bFirstRef = False
                        End If
                        bHasCrossRef = True
                        nCrossRefCount = nCrossRefCount + 1
                        oLog.WriteLine "    -> WRONG REF: " & sRefPath
                    End If
                End If
            Next
            
            oDoc.Close True ' Close without saving
            Set oDoc = Nothing
        End If
    Next
    
    oLog.WriteLine ""
    oLog.WriteLine "  SUMMARY: " & nIDWsWithIssues & " IDWs with wrong references, " & nCrossRefCount & " total bad refs"
End Sub

Sub CollectIDWFiles(sFolder, dicIDWs)
    Dim oFolder, oFile, oSubFolder
    
    If Not oFSO.FolderExists(sFolder) Then Exit Sub
    Set oFolder = oFSO.GetFolder(sFolder)
    
    ' Skip OldVersions folders
    If LCase(oFolder.Name) = "oldversions" Then Exit Sub
    If LCase(oFolder.Name) = "old" Then Exit Sub
    
    For Each oFile In oFolder.Files
        If LCase(oFSO.GetExtensionName(oFile.Name)) = "idw" Then
            dicIDWs(LCase(oFile.Path)) = oFile.Path
        End If
    Next
    
    For Each oSubFolder In oFolder.SubFolders
        CollectIDWFiles oSubFolder.Path, dicIDWs
    Next
End Sub
