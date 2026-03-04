' TEST_Fresh_CompDef.vbs
' Re-get the ComponentDefinition after reverting
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, cmdMgr

WScript.Echo "=== FRESH COMPONENT DEFINITION TEST ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
Set partDoc = invApp.ActiveDocument

WScript.Echo "Part: " & partDoc.DisplayName
WScript.Echo "Document SubType: " & partDoc.SubType
WScript.Echo ""

' Get the compdef fresh
Set compDef = partDoc.ComponentDefinition
WScript.Echo "ComponentDefinition Type: " & TypeName(compDef)
WScript.Echo ""

' If it's sheet metal, revert
If partDoc.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
    WScript.Echo "=== REVERTING ==="
    
    Set cmdMgr = invApp.CommandManager
    
    ' Delete flat pattern if exists
    If compDef.HasFlatPattern Then
        compDef.FlatPattern.Delete
        partDoc.Update
        WScript.Echo "Flat pattern deleted"
    End If
    
    Dim revertCmd
    Set revertCmd = cmdMgr.ControlDefinitions.Item("PartConvertToStandardPartCmd")
    
    If revertCmd.Enabled Then
        revertCmd.Execute
        WScript.Sleep 1500
        
        ' IMPORTANT: Re-open or re-get the document
        partDoc.Save
        WScript.Echo "Document saved"
        
        ' Force update
        partDoc.Update
        
        ' Re-get component definition
        Set compDef = Nothing
        Set compDef = partDoc.ComponentDefinition
        
        WScript.Echo ""
        WScript.Echo "After revert:"
        WScript.Echo "  Document SubType: " & partDoc.SubType
        WScript.Echo "  ComponentDef Type: " & TypeName(compDef)
    End If
    WScript.Echo ""
End If

' Now check for ConvertToSheetMetalFeatures
WScript.Echo "=== CHECKING API ACCESS ==="

Dim features
Set features = compDef.Features

WScript.Echo "Features Type: " & TypeName(features)
WScript.Echo "Features Count: " & features.Count
WScript.Echo ""

' List all available feature collections
WScript.Echo "Available feature collections:"

Dim ctsm
Set ctsm = features.ConvertToSheetMetalFeatures
If Err.Number = 0 Then
    WScript.Echo "  ConvertToSheetMetalFeatures: " & TypeName(ctsm)
Else
    WScript.Echo "  ConvertToSheetMetalFeatures: ERROR - " & Err.Description
    Err.Clear
End If

Dim extf
Set extf = features.ExtrudeFeatures  
If Err.Number = 0 Then
    WScript.Echo "  ExtrudeFeatures: " & TypeName(extf) & " (Count=" & extf.Count & ")"
Else
    WScript.Echo "  ExtrudeFeatures: ERROR - " & Err.Description
    Err.Clear
End If

Dim ff
Set ff = features.FaceFeatures
If Err.Number = 0 Then
    WScript.Echo "  FaceFeatures: " & TypeName(ff) & " (Count=" & ff.Count & ")"
Else
    WScript.Echo "  FaceFeatures: ERROR - " & Err.Description  
    Err.Clear
End If

' Check if it's really a PartComponentDefinition
WScript.Echo ""
WScript.Echo "=== COMPONENT DEFINITION PROPERTIES ==="

Dim smcd
Set smcd = compDef
WScript.Echo "SmCompDef.SubType: " & smcd.SubType
If Err.Number <> 0 Then Err.Clear

' Try to use SheetMetalComponentDefinition methods
WScript.Echo ""
WScript.Echo "Testing as SheetMetalComponentDefinition..."
Dim testFp
testFp = smcd.HasFlatPattern
If Err.Number = 0 Then
    WScript.Echo "  HasFlatPattern: " & testFp & " (SheetMetal method works)"
Else
    WScript.Echo "  HasFlatPattern: ERROR - " & Err.Description
    Err.Clear
End If

' Check for SheetMetalStyles
Dim smStyles  
Set smStyles = smcd.SheetMetalStyles
If Err.Number = 0 Then
    WScript.Echo "  SheetMetalStyles: " & TypeName(smStyles)
Else
    WScript.Echo "  SheetMetalStyles: ERROR - " & Err.Description
    Err.Clear
End If

WScript.Echo ""
WScript.Echo "=== DONE ==="
