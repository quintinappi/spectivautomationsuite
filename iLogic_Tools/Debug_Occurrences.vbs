' Debug Part Scanner - Show occurrences
Option Explicit
Const kAssemblyDocumentObject = 12291

Dim invApp, asmDoc

On Error Resume Next

Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Or invApp Is Nothing Then
    WScript.Echo "ERROR: Inventor not running"
    WScript.Quit 1
End If

Set asmDoc = invApp.ActiveDocument
If Err.Number <> 0 Or asmDoc Is Nothing Then
    WScript.Echo "ERROR: No active document"
    WScript.Quit 1
End If

WScript.Echo "Assembly: " & asmDoc.DisplayName
WScript.Echo "Document Type: " & asmDoc.DocumentType

Dim compDef, occs, i, occ
Err.Clear
Set compDef = asmDoc.ComponentDefinition
WScript.Echo "ComponentDefinition retrieved. Error: " & Err.Number

Err.Clear
Set occs = compDef.Occurrences
WScript.Echo "Occurrences retrieved. Count: " & occs.Count & " Error: " & Err.Number

For i = 1 To occs.Count
    Err.Clear
    Set occ = occs.Item(i)
    If Err.Number = 0 Then
        Dim doc
        Set doc = occ.Definition.Document
        Dim fname
        fname = Mid(doc.FullFileName, InStrRev(doc.FullFileName, "\") + 1)
        WScript.Echo i & ". " & fname & " (Suppressed: " & occ.Suppressed & ")"
    Else
        WScript.Echo i & ". ERROR getting occurrence: " & Err.Number
    End If
Next
