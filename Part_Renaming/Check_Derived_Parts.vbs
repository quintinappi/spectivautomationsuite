' Check_Derived_Parts.vbs
' Experimental script to detect derived parts in the current open assembly
' and report where their base components are located.
'
' Run with Inventor open and an assembly (.iam) active
'
Option Explicit

Dim invApp, activeDoc
Dim fso, logFile, logPath
Dim derivedCount, partCount

' Initialize
Set fso = CreateObject("Scripting.FileSystemObject")
logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\DerivedParts_Report.txt"

On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Then
    MsgBox "Inventor is not running. Please open Inventor with an assembly first.", vbCritical, "Error"
    WScript.Quit
End If
On Error GoTo 0

Set activeDoc = invApp.ActiveDocument
If activeDoc Is Nothing Then
    MsgBox "No document is open in Inventor.", vbCritical, "Error"
    WScript.Quit
End If

If activeDoc.DocumentType <> 12291 Then ' kAssemblyDocumentObject = 12291
    MsgBox "Active document is not an assembly (.iam). Please open an assembly.", vbCritical, "Error"
    WScript.Quit
End If

' Create log file
Set logFile = fso.CreateTextFile(logPath, True)
logFile.WriteLine "=========================================="
logFile.WriteLine " DERIVED PARTS DETECTION REPORT"
logFile.WriteLine "=========================================="
logFile.WriteLine ""
logFile.WriteLine "Assembly: " & activeDoc.FullFileName
logFile.WriteLine "Scan Time: " & Now()
logFile.WriteLine ""
logFile.WriteLine "------------------------------------------"

derivedCount = 0
partCount = 0

' Scan all referenced documents
Dim refDoc, refDocs
Set refDocs = activeDoc.AllReferencedDocuments

logFile.WriteLine "Total Referenced Documents: " & refDocs.Count
logFile.WriteLine "------------------------------------------"
logFile.WriteLine ""

Dim doc
For Each doc In refDocs
    ' Only check part documents
    If doc.DocumentType = 12290 Then ' kPartDocumentObject = 12290
        partCount = partCount + 1
        CheckPartForDerived doc, logFile
    End If
Next

logFile.WriteLine ""
logFile.WriteLine "=========================================="
logFile.WriteLine " SUMMARY"
logFile.WriteLine "=========================================="
logFile.WriteLine "Parts Scanned: " & partCount
logFile.WriteLine "Derived Parts Found: " & derivedCount
logFile.WriteLine "=========================================="

logFile.Close

MsgBox "Scan complete!" & vbCrLf & vbCrLf & _
       "Parts Scanned: " & partCount & vbCrLf & _
       "Derived Parts Found: " & derivedCount & vbCrLf & vbCrLf & _
       "Full report saved to:" & vbCrLf & logPath, _
       vbInformation, "Derived Parts Check"

' Open the log file
Dim shell
Set shell = CreateObject("WScript.Shell")
shell.Run """" & logPath & """", 1, False

WScript.Quit

'------------------------------------------------------------------------------
Sub CheckPartForDerived(partDoc, log)
    On Error Resume Next
    
    Dim partDef, refComps, derivedParts
    Dim derivedComp, docDesc, basePath
    Dim partName, partPath
    
    partPath = partDoc.FullFileName
    partName = fso.GetFileName(partPath)
    
    Set partDef = partDoc.ComponentDefinition
    If Err.Number <> 0 Then
        log.WriteLine "[ERROR] Could not get ComponentDefinition for: " & partName
        Err.Clear
        Exit Sub
    End If
    
    Set refComps = partDef.ReferenceComponents
    If Err.Number <> 0 Then
        log.WriteLine "[ERROR] Could not get ReferenceComponents for: " & partName
        Err.Clear
        Exit Sub
    End If
    
    Set derivedParts = refComps.DerivedPartComponents
    If Err.Number <> 0 Then
        ' No DerivedPartComponents - this is normal for regular parts
        Err.Clear
        Exit Sub
    End If
    
    If derivedParts.Count = 0 Then
        Exit Sub
    End If
    
    ' This part has derived components!
    log.WriteLine "DERIVED PART FOUND: " & partName
    log.WriteLine "  Part Location: " & partPath
    log.WriteLine "  Derived Component Count: " & derivedParts.Count
    
    Dim i
    For i = 1 To derivedParts.Count
        Set derivedComp = derivedParts.Item(i)
        
        log.WriteLine ""
        log.WriteLine "  [Derived #" & i & "]"
        log.WriteLine "    Name: " & derivedComp.Name
        
        ' Check if still linked
        If derivedComp.LinkedToFile Then
            log.WriteLine "    Linked: YES"
            
            ' Get the base file path
            Set docDesc = derivedComp.ReferencedDocumentDescriptor
            If Not docDesc Is Nothing Then
                basePath = docDesc.FullDocumentName
                log.WriteLine "    Base Component Path: " & basePath
                
                ' Check if base file exists
                If fso.FileExists(basePath) Then
                    log.WriteLine "    Base File Exists: YES"
                Else
                    log.WriteLine "    Base File Exists: NO (MISSING!)"
                End If
                
                ' Show relative location
                Dim partFolder, baseFolder
                partFolder = fso.GetParentFolderName(partPath)
                baseFolder = fso.GetParentFolderName(basePath)
                
                If LCase(partFolder) = LCase(baseFolder) Then
                    log.WriteLine "    Relative Location: SAME FOLDER"
                ElseIf InStr(1, LCase(basePath), LCase(partFolder), vbTextCompare) > 0 Then
                    log.WriteLine "    Relative Location: SUBFOLDER of part"
                ElseIf InStr(1, LCase(partPath), LCase(baseFolder), vbTextCompare) > 0 Then
                    log.WriteLine "    Relative Location: PARENT folder of part"
                Else
                    log.WriteLine "    Relative Location: DIFFERENT folder tree"
                End If
            Else
                log.WriteLine "    Base Component Path: [Could not retrieve]"
            End If
        Else
            log.WriteLine "    Linked: NO (link broken)"
        End If
        
        derivedCount = derivedCount + 1
    Next
    
    log.WriteLine ""
    log.WriteLine "------------------------------------------"
    
    On Error GoTo 0
End Sub
