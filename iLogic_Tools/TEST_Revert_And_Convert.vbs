' TEST_Revert_And_Convert.vbs
' Revert sheet metal part and convert with correct face
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, selectSet, cmdMgr

WScript.Echo "=== REVERT AND CONVERT ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
Set partDoc = invApp.ActiveDocument
Set compDef = partDoc.ComponentDefinition
Set selectSet = partDoc.SelectSet
Set cmdMgr = invApp.CommandManager

WScript.Echo "Part: " & partDoc.DisplayName
WScript.Echo "SubType: " & compDef.SubType
WScript.Echo ""

' Method: Use Undo to revert
WScript.Echo "=== CHECKING UNDO AVAILABILITY ==="

Dim undoMgr
Set undoMgr = partDoc.UnitsOfMeasure
' No undo manager in VBScript...

' Try the revert command
WScript.Echo ""
WScript.Echo "=== LOOKING FOR REVERT COMMAND ==="

Dim ctrlDefs, ctrlDef
Set ctrlDefs = cmdMgr.ControlDefinitions

' List sheet metal related commands
Dim cmdNames
cmdNames = Array("PartRevertToStandardDocCmd", "SMRevertToStandardDocCmd", "RevertToStandardDocCmd", "PartSheetMetalRevertCmd")

Dim cmdName
For Each cmdName In cmdNames
    Set ctrlDef = ctrlDefs.Item(cmdName)
    If Err.Number = 0 Then
        WScript.Echo cmdName & ": " & ctrlDef.DisplayName & " (Enabled=" & ctrlDef.Enabled & ")"
    Else
        ' WScript.Echo cmdName & ": Not found"
    End If
    Err.Clear
Next

' Search for any command with "revert" in the name
WScript.Echo ""
WScript.Echo "=== SEARCHING FOR REVERT COMMANDS ==="

Dim foundCount
foundCount = 0

For Each ctrlDef In ctrlDefs
    Dim lowerName
    lowerName = LCase(ctrlDef.InternalName)
    If InStr(lowerName, "revert") > 0 Or InStr(lowerName, "standard") > 0 Then
        WScript.Echo ctrlDef.InternalName & ": " & ctrlDef.DisplayName
        foundCount = foundCount + 1
        If foundCount > 20 Then Exit For
    End If
    Err.Clear
Next

' Try to find and delete the ConvertToSheetMetal feature
WScript.Echo ""
WScript.Echo "=== FINDING SHEET METAL FEATURES ==="

Dim features, feat
Set features = compDef.Features

WScript.Echo "Total features: " & features.Count

For Each feat In features
    Dim featType
    featType = TypeName(feat)
    If InStr(LCase(featType), "sheet") > 0 Or InStr(LCase(featType), "convert") > 0 Then
        WScript.Echo "Found: " & feat.Name & " (" & featType & ")"
    End If
    Err.Clear
Next

' List all feature types
WScript.Echo ""
WScript.Echo "=== ALL FEATURE TYPES ==="
Dim featTypes
Set featTypes = CreateObject("Scripting.Dictionary")

For Each feat In features
    Dim ft
    ft = TypeName(feat)
    If Not featTypes.Exists(ft) Then
        featTypes.Add ft, 1
    End If
    Err.Clear
Next

Dim key
For Each key In featTypes.Keys
    WScript.Echo "  " & key
Next

WScript.Echo ""
WScript.Echo "=== DONE ==="
