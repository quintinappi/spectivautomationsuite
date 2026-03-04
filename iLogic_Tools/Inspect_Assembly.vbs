Option Explicit
On Error Resume Next
Dim invApp
Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Or invApp Is Nothing Then
    WScript.Echo "ERROR: Inventor not running"
    WScript.Quit 1
End If
Err.Clear

Dim asmDoc
Set asmDoc = Nothing
If Not invApp.ActiveDocument Is Nothing And invApp.ActiveDocument.DocumentType = 12291 Then
    Set asmDoc = invApp.ActiveDocument
Else
    Dim d
    For Each d In invApp.Documents
        If d.DocumentType = 12291 Then Set asmDoc = d: Exit For
    Next
End If

If asmDoc Is Nothing Then
    WScript.Echo "No assembly found"
    WScript.Quit 0
End If

WScript.Echo "Assembly: " & asmDoc.DisplayName
WScript.Echo "Path: " & asmDoc.FullFileName
WScript.Echo "DocumentType: " & asmDoc.DocumentType
On Error Resume Next
WScript.Echo "SubType: " & asmDoc.SubType
If Err.Number <> 0 Then Err.Clear

On Error Resume Next
Dim occCount
occCount = asmDoc.ComponentDefinition.Occurrences.Count
WScript.Echo "Occurrences: " & occCount
If Err.Number <> 0 Then
    WScript.Echo "Could not read ComponentDefinition.Occurrences - " & Err.Description
    Err.Clear
End If

' Check representations
On Error Resume Next
Dim repMgr
Set repMgr = asmDoc.RepresentationsManager
If Err.Number = 0 And Not repMgr Is Nothing Then
    On Error Resume Next
    Dim mrCount
    mrCount = repMgr.ModelRepresentations.Count
    If Err.Number = 0 Then
        WScript.Echo "ModelRepresentations.Count: " & mrCount
    Else
        Err.Clear
        WScript.Echo "ModelRepresentations.Count not available"
    End If
Else
    WScript.Echo "No RepresentationsManager available"
End If


WScript.Echo "Assembly inspection complete."