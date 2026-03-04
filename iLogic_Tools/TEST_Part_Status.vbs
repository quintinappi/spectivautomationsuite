' TEST_Part_Status.vbs
' Check current part status in detail
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef

WScript.Echo "=== PART STATUS CHECK ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
Set partDoc = invApp.ActiveDocument
Set compDef = partDoc.ComponentDefinition

WScript.Echo "Document:"
WScript.Echo "  FullFileName: " & partDoc.FullFileName
WScript.Echo "  DisplayName: " & partDoc.DisplayName
WScript.Echo "  DocumentType: " & partDoc.DocumentType
WScript.Echo "  IsModifiable: " & partDoc.IsModifiable
WScript.Echo "  SubType: " & partDoc.SubType
WScript.Echo ""

WScript.Echo "ComponentDefinition:"
WScript.Echo "  Type: " & TypeName(compDef)
WScript.Echo "  SubType: " & compDef.SubType
WScript.Echo "  SubType GUID: [" & compDef.SubType & "]"

' Check for specific subtypes
Dim guid
guid = compDef.SubType

WScript.Echo ""
WScript.Echo "Interpreting SubType:"
If guid = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
    WScript.Echo "  -> Sheet Metal Part"
ElseIf guid = "{4D29B490-49B2-11D0-93C3-7E0706000000}" Then
    WScript.Echo "  -> Standard Part"
ElseIf guid = "" Or IsEmpty(guid) Or IsNull(guid) Then
    WScript.Echo "  -> EMPTY/NULL (possible derived or special part)"
Else
    WScript.Echo "  -> Other: " & guid
End If

WScript.Echo ""
WScript.Echo "Features:"
Dim feat
For Each feat In compDef.Features
    WScript.Echo "  " & feat.Name & " (" & TypeName(feat) & ")"
Next

WScript.Echo ""
WScript.Echo "SurfaceBodies:"
WScript.Echo "  Count: " & compDef.SurfaceBodies.Count

If compDef.SurfaceBodies.Count > 0 Then
    Dim body
    Set body = compDef.SurfaceBodies.Item(1)
    WScript.Echo "  Body 1 Faces: " & body.Faces.Count
End If

WScript.Echo ""
WScript.Echo "Derived Part Check:"
WScript.Echo "  IsReferencedDocumentComplete: " & partDoc.ReferencedDocumentDescriptors.Count

Dim refDoc
For Each refDoc In partDoc.ReferencedDocumentDescriptors
    WScript.Echo "  -> " & refDoc.DisplayName
Next

WScript.Echo ""
WScript.Echo "=== COMMAND AVAILABILITY ==="

Dim cmdMgr, ctrlDef
Set cmdMgr = invApp.CommandManager

Set ctrlDef = cmdMgr.ControlDefinitions.Item("PartConvertToSheetMetalCmd")
WScript.Echo "PartConvertToSheetMetalCmd: Enabled=" & ctrlDef.Enabled

Set ctrlDef = cmdMgr.ControlDefinitions.Item("PartConvertToStandardPartCmd")
WScript.Echo "PartConvertToStandardPartCmd: Enabled=" & ctrlDef.Enabled

WScript.Echo ""
WScript.Echo "=== DONE ==="
