' TEST_Active_Doc_Details.vbs
' Get full details about active document
Option Explicit
On Error Resume Next

Dim invApp, doc, compDef

WScript.Echo "=== ACTIVE DOCUMENT DETAILS ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Then
    WScript.Echo "Cannot connect to Inventor"
    WScript.Quit
End If

Set doc = invApp.ActiveDocument
WScript.Echo "DisplayName: " & doc.DisplayName
WScript.Echo "FullFileName: " & doc.FullFileName
WScript.Echo ""
WScript.Echo "DocumentType: " & doc.DocumentType
WScript.Echo "DocumentTypeID: " & doc.DocumentTypeID

WScript.Echo ""
WScript.Echo "--- Trying to get ComponentDefinition ---"

Set compDef = doc.ComponentDefinition
If Err.Number <> 0 Then
    WScript.Echo "Error getting ComponentDefinition: " & Err.Description
    Err.Clear
Else
    WScript.Echo "ComponentDefinition.Type: " & compDef.Type
    WScript.Echo ""
    WScript.Echo "Type constants:"
    WScript.Echo "  99588099 = kSheetMetalComponentDefinitionObject"
    WScript.Echo "  150995200 = kAssemblyComponentDefinitionObject"
    WScript.Echo "  100675072 = kPartComponentDefinitionObject"
    
    WScript.Echo ""
    WScript.Echo "SubType GUID: " & compDef.SubType
    
    ' Interpret SubType
    Select Case compDef.SubType
        Case "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
            WScript.Echo "  >>> SHEET METAL <<<"
        Case "{4D29B490-49B2-11D0-93C3-7E0706000000}"
            WScript.Echo "  >>> STANDARD PART <<<"
        Case Else
            WScript.Echo "  (Unknown SubType)"
    End Select
End If

' Try to check if it has flat pattern capability
WScript.Echo ""
WScript.Echo "--- Sheet Metal checks ---"

Dim smDef
Set smDef = compDef
WScript.Echo "HasFlatPattern: " & smDef.HasFlatPattern
If Err.Number <> 0 Then
    WScript.Echo "  Error: " & Err.Description
    Err.Clear
End If

WScript.Echo "Thickness: " & smDef.Thickness.Value * 10 & " mm"
If Err.Number <> 0 Then
    WScript.Echo "  Error: " & Err.Description
    Err.Clear
End If

' Check Features
WScript.Echo ""
WScript.Echo "--- Features ---"
WScript.Echo "Features count: " & compDef.Features.Count
If Err.Number <> 0 Then
    WScript.Echo "  Error: " & Err.Description
    Err.Clear
End If

' Look for SheetMetal features
Dim feat
For Each feat In compDef.Features
    WScript.Echo "  " & TypeName(feat) & ": " & feat.Name
    If Err.Number <> 0 Then
        Err.Clear
    End If
Next

WScript.Echo ""
WScript.Echo "=== DONE ==="
