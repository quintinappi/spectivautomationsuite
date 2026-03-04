On Error Resume Next

' Connect to Inventor
Dim invApp
Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Inventor not running or not accessible"
    WScript.Quit 1
End If

' Get active document
Dim activeDoc
Set activeDoc = invApp.ActiveDocument
If activeDoc Is Nothing Then
    WScript.Echo "ERROR: No active document"
    WScript.Quit 1
End If

WScript.Echo "SCANNING OPEN ASSEMBLY: " & activeDoc.FullFileName
WScript.Echo "Document Type: " & activeDoc.DocumentType
WScript.Echo ""

' Check ReferencedFileDescriptors
Dim fds
Set fds = activeDoc.File.ReferencedFileDescriptors
WScript.Echo "REFERENCED FILE DESCRIPTORS (" & fds.Count & " total):"
Dim fd
For Each fd In fds
    WScript.Echo "  " & fd.FullFileName & " -> " & fd.ReferencedFileDescriptorStatus
Next
WScript.Echo ""

' Check ComponentDefinition.Occurrences
Dim occs
Set occs = activeDoc.ComponentDefinition.Occurrences
WScript.Echo "COMPONENT OCCURRENCES (" & occs.Count & " total):"
Dim occ
For Each occ In occs
    Dim defDoc
    Set defDoc = Nothing
    Set defDoc = occ.Definition.Document
    If Err.Number <> 0 Then
        WScript.Echo "  " & occ.Name & " -> UNRESOLVED (Error: " & Err.Description & ")"
        Err.Clear
    Else
        WScript.Echo "  " & occ.Name & " -> " & defDoc.FullFileName & " (" & defDoc.DocumentType & ")"
    End If
Next

WScript.Echo ""
WScript.Echo "SCAN COMPLETE"