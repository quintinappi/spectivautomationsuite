' TEST_Check_Part_Status.vbs
' Check current status of part
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef

WScript.Echo "=== CHECK PART STATUS ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Then
    WScript.Echo "Cannot connect to Inventor"
    WScript.Quit
End If

Set partDoc = invApp.ActiveDocument
If partDoc Is Nothing Then
    WScript.Echo "No active document"
    WScript.Quit
End If

WScript.Echo "Document: " & partDoc.DisplayName
WScript.Echo "Full path: " & partDoc.FullFileName
WScript.Echo "Document type: " & partDoc.DocumentType
WScript.Echo "  (12291 = Part, 12290 = Assembly, 12292 = Drawing)"

If partDoc.DocumentType = 12291 Then
    Set compDef = partDoc.ComponentDefinition
    WScript.Echo ""
    WScript.Echo "Component Definition Type: " & compDef.Type
    WScript.Echo "  99588099 = SheetMetalComponentDefinition"
    WScript.Echo "  150995200 = PartComponentDefinition"
    
    ' Check SubType
    WScript.Echo ""
    WScript.Echo "SubType: " & compDef.SubType
    WScript.Echo ""
    
    Select Case compDef.SubType
        Case "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
            WScript.Echo "*** SHEET METAL PART ***"
        Case "{4D29B490-49B2-11D0-93C3-7E0706000000}"
            WScript.Echo "*** STANDARD PART ***"
        Case Else
            WScript.Echo "Other part type"
    End Select
    
    ' Try to access sheet metal features
    WScript.Echo ""
    WScript.Echo "Checking for sheet metal features..."
    
    Dim smDef
    Set smDef = compDef
    
    WScript.Echo "HasFlatPattern: " & smDef.HasFlatPattern
    If Err.Number <> 0 Then
        WScript.Echo "  (Error: " & Err.Description & ")"
        Err.Clear
    End If
    
    WScript.Echo "Thickness: " & smDef.Thickness.Value * 10 & " mm"
    If Err.Number <> 0 Then
        WScript.Echo "  (Error: " & Err.Description & ")"
        Err.Clear
    End If
    
ElseIf partDoc.DocumentType = 12290 Then
    WScript.Echo "This is an Assembly document"
End If

WScript.Echo ""
WScript.Echo "=== DONE ==="
