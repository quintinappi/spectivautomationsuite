' Scan_IDW_CrossRefs.vbs
' Scans IDW files in two folders and identifies cross-references
' Usage: cscript Scan_IDW_CrossRefs.vbs

Option Explicit

Dim oShell, oFSO, oInventor
Dim sFolder1, sFolder2
Dim dicFolder1IPTs, dicFolder2IPTs
Dim dicFolder1IAMs, dicFolder2IAMs
Dim oCrossRefIssues
Dim sLogFile

' Set paths
sFolder1 = "C:\Users\Quintin\Documents\Spectiv\3. Working\22. DMS plant\21. SSCR04 - Primary Floats D&R Screen Station"
sFolder2 = "C:\Users\Quintin\Documents\Spectiv\3. Working\22. DMS plant\22. SSCR05 - Primary Sinks D&R Screen Station"

Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set dicFolder1IPTs = CreateObject("Scripting.Dictionary")
Set dicFolder2IPTs = CreateObject("Scripting.Dictionary")
Set dicFolder1IAMs = CreateObject("Scripting.Dictionary")
Set dicFolder2IAMs = CreateObject("Scripting.Dictionary")
Set oCrossRefIssues = CreateObject("Scripting.Dictionary")

sLogFile = oFSO.GetParentFolderName(WScript.ScriptFullName) & "\CrossRef_Report.txt"

' Collect all IPT/IAM files in each folder
WScript.Echo "Collecting file lists..."
CollectFiles sFolder1, dicFolder1IPTs, dicFolder1IAMs
CollectFiles sFolder2, dicFolder2IPTs, dicFolder2IAMs

WScript.Echo "Folder 1 (SSCR04): " & dicFolder1IPTs.Count & " IPTs, " & dicFolder1IAMs.Count & " IAMs"
WScript.Echo "Folder 2 (SSCR05): " & dicFolder2IPTs.Count & " IPTs, " & dicFolder2IAMs.Count & " IAMs"

' Connect to Inventor
WScript.Echo vbCrLf & "Connecting to Inventor..."
On Error Resume Next
Set oInventor = GetObject(, "Inventor.Application")
If Err.Number <> 0 Then
    Err.Clear
    Set oInventor = CreateObject("Inventor.Application")
    oInventor.Visible = True
End If
On Error GoTo 0

If oInventor Is Nothing Then
    WScript.Echo "ERROR: Cannot connect to Inventor!"
    WScript.Quit 1
End If

WScript.Echo "Connected to Inventor successfully!"

' Scan IDWs in both folders
Dim oLogFile
Set oLogFile = oFSO.CreateTextFile(sLogFile, True)
oLogFile.WriteLine "IDW Cross-Reference Report"
oLogFile.WriteLine "Generated: " & Now()
oLogFile.WriteLine "=" & String(60, "=")
oLogFile.WriteLine ""

' Scan Folder 1 IDWs for references to Folder 2
oLogFile.WriteLine "FOLDER 1 (SSCR04) IDWs with references to FOLDER 2 (SSCR05):"
oLogFile.WriteLine "-" & String(60, "-")
WScript.Echo vbCrLf & "Scanning Folder 1 IDWs for cross-references to Folder 2..."
ScanIDWsForCrossRefs sFolder1, sFolder2, dicFolder2IPTs, oLogFile, "SSCR05"

' Scan Folder 2 IDWs for references to Folder 1  
oLogFile.WriteLine ""
oLogFile.WriteLine "FOLDER 2 (SSCR05) IDWs with references to FOLDER 1 (SSCR04):"
oLogFile.WriteLine "-" & String(60, "-")
WScript.Echo vbCrLf & "Scanning Folder 2 IDWs for cross-references to Folder 1..."
ScanIDWsForCrossRefs sFolder2, sFolder1, dicFolder1IPTs, oLogFile, "SSCR04"

oLogFile.WriteLine ""
oLogFile.WriteLine "=" & String(60, "=")
oLogFile.WriteLine "SCAN COMPLETE"
oLogFile.Close

WScript.Echo vbCrLf & "Report saved to: " & sLogFile
WScript.Echo "Done!"

' ============================================
' FUNCTIONS
' ============================================

Sub CollectFiles(sFolder, dicIPTs, dicIAMs)
    Dim oFolder, oFile, oSubFolder
    
    If Not oFSO.FolderExists(sFolder) Then Exit Sub
    Set oFolder = oFSO.GetFolder(sFolder)
    
    ' Skip OldVersions folders
    If LCase(oFolder.Name) = "oldversions" Then Exit Sub
    
    For Each oFile In oFolder.Files
        If LCase(oFSO.GetExtensionName(oFile.Name)) = "ipt" Then
            dicIPTs(LCase(oFile.Path)) = oFile.Path
        ElseIf LCase(oFSO.GetExtensionName(oFile.Name)) = "iam" Then
            dicIAMs(LCase(oFile.Path)) = oFile.Path
        End If
    Next
    
    For Each oSubFolder In oFolder.SubFolders
        CollectFiles oSubFolder.Path, dicIPTs, dicIAMs
    Next
End Sub

Sub ScanIDWsForCrossRefs(sSourceFolder, sTargetFolder, dicTargetFiles, oLog, sTargetName)
    Dim oFolder, dicIDWs, sKey
    Dim oDoc, oRefDoc, sRefPath
    Dim nIDWCount, nCrossRefCount, nIDWsWithIssues
    
    Set dicIDWs = CreateObject("Scripting.Dictionary")
    CollectIDWFiles sSourceFolder, dicIDWs
    
    nIDWCount = 0
    nCrossRefCount = 0
    nIDWsWithIssues = 0
    
    For Each sKey In dicIDWs.Keys
        nIDWCount = nIDWCount + 1
        Dim bHasCrossRef
        bHasCrossRef = False
        
        WScript.Echo "  [" & nIDWCount & "/" & dicIDWs.Count & "] " & oFSO.GetFileName(sKey)
        
        On Error Resume Next
        Set oDoc = oInventor.Documents.Open(sKey, False) ' Open invisible
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            oLog.WriteLine "  SKIPPED (open error): " & sKey
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
                    ' Check if this reference points to the target folder
                    If InStr(1, LCase(sRefPath), LCase(oFSO.GetFileName(oFSO.GetParentFolderName(sTargetFolder))), vbTextCompare) > 0 Then
                        If Not bHasCrossRef Then
                            oLog.WriteLine ""
                            oLog.WriteLine "  IDW: " & sKey
                            nIDWsWithIssues = nIDWsWithIssues + 1
                        End If
                        bHasCrossRef = True
                        nCrossRefCount = nCrossRefCount + 1
                        oLog.WriteLine "    -> CROSS-REF to " & sTargetName & ": " & sRefPath
                    End If
                End If
            Next
            
            oDoc.Close True ' Close without saving
        End If
    Next
    
    oLog.WriteLine ""
    oLog.WriteLine "  SUMMARY: " & nIDWsWithIssues & " IDWs with cross-references, " & nCrossRefCount & " total cross-refs"
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
