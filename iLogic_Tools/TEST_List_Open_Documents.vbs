' TEST_List_Open_Documents.vbs
' List all open documents and their types
Option Explicit
On Error Resume Next

Dim invApp, doc

WScript.Echo "=== OPEN DOCUMENTS IN INVENTOR ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Then
    WScript.Echo "Cannot connect to Inventor"
    WScript.Quit
End If

WScript.Echo "Total open documents: " & invApp.Documents.Count
WScript.Echo ""

Dim i
i = 0
For Each doc In invApp.Documents
    i = i + 1
    WScript.Echo "--- Document " & i & " ---"
    WScript.Echo "Display Name: " & doc.DisplayName
    WScript.Echo "Full Path: " & doc.FullFileName
    WScript.Echo "Type: " & doc.DocumentType
    
    Select Case doc.DocumentType
        Case 12291
            WScript.Echo "       (Part)"
        Case 12290
            WScript.Echo "       (Assembly)"
        Case 12292
            WScript.Echo "       (Drawing)"
    End Select
    
    If doc.DocumentType = 12291 Then
        Dim compDef
        Set compDef = doc.ComponentDefinition
        WScript.Echo "SubType: " & compDef.SubType
        
        If compDef.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
            WScript.Echo "       *** SHEET METAL ***"
            WScript.Echo "HasFlatPattern: " & compDef.HasFlatPattern
        ElseIf compDef.SubType = "{4D29B490-49B2-11D0-93C3-7E0706000000}" Then
            WScript.Echo "       (Standard Part)"
        End If
        Err.Clear
    End If
    
    WScript.Echo ""
Next

WScript.Echo "Active Document: " & invApp.ActiveDocument.DisplayName
WScript.Echo ""
WScript.Echo "=== DONE ==="
