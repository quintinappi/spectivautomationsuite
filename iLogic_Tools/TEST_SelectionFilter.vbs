' TEST_SelectionFilter.vbs
' Try using selection filter to ensure correct face is selected for command
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, cmdMgr, intEvents

WScript.Echo "=== SELECTION FILTER APPROACH ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
Set partDoc = invApp.ActiveDocument
Set compDef = partDoc.ComponentDefinition
Set cmdMgr = invApp.CommandManager

WScript.Echo "Part: " & partDoc.DisplayName

' Find largest face
Dim body, face, largestFace, largestArea, largestFaceIdx
largestArea = 0
largestFaceIdx = 0
Dim idx
idx = 0
Set body = compDef.SurfaceBodies.Item(1)
For Each face In body.Faces
    idx = idx + 1
    Dim area
    area = face.Evaluator.Area * 100
    If area > largestArea Then
        largestArea = area
        Set largestFace = face
        largestFaceIdx = idx
    End If
    Err.Clear
Next
WScript.Echo "Largest face: #" & largestFaceIdx & " (" & FormatNumber(largestArea, 0) & " mm²)"

' Check current sheet metal status and flatten if needed
WScript.Echo ""
WScript.Echo "=== CHECK STATUS ==="

WScript.Echo "SubType: " & compDef.SubType
Dim isSheetMetal
isSheetMetal = (compDef.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
WScript.Echo "Is Sheet Metal: " & isSheetMetal

If Not isSheetMetal Then
    ' Convert to sheet metal via SubType
    WScript.Echo "Converting to sheet metal..."
    compDef.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
    partDoc.Update
    WScript.Echo "Converted!"
    Err.Clear
End If

' Delete any flat pattern
If compDef.HasFlatPattern Then
    compDef.FlatPattern.Delete
    partDoc.Update
    WScript.Echo "Deleted old flat pattern"
End If

' Try to use NameValueMap to pass face info
WScript.Echo ""
WScript.Echo "=== NAMEVALUE MAP APPROACH ==="

Dim nvm
Set nvm = invApp.TransientObjects.CreateNameValueMap

' Try adding face to NVM
nvm.Add "BaseFace", largestFace
If Err.Number = 0 Then
    WScript.Echo "Added face to NameValueMap"
Else
    WScript.Echo "NVM add failed: " & Err.Description
    Err.Clear
End If

' Look for ExecuteCommand with parameters
WScript.Echo ""
WScript.Echo "=== EXECUTE WITH PARAMS ==="

Dim convertCmd
Set convertCmd = cmdMgr.ControlDefinitions.Item("PartConvertToSheetMetalCmd")
WScript.Echo "PartConvertToSheetMetalCmd - Enabled: " & convertCmd.Enabled
Err.Clear

' Try Execute2 or ExecuteWithArgs
WScript.Echo ""
WScript.Echo "Trying Execute2(nvm)..."
convertCmd.Execute2 nvm
WScript.Echo "  Result: " & Err.Description
Err.Clear

' Try direct ControlDefinition.Execute with parameter
WScript.Echo "Trying Execute nvm..."
convertCmd.Execute nvm
WScript.Echo "  Result: " & Err.Description
Err.Clear

' Alternative: check for CommandInput
WScript.Echo ""
WScript.Echo "=== COMMANDINPUT ==="

Dim ci
Set ci = cmdMgr.ControlDefinitions.CommandInput
If Not ci Is Nothing Then
    WScript.Echo "CommandInput found!"
Else
    WScript.Echo "No CommandInput"
End If
Err.Clear

' Check if we can use the Face as the starting selection
WScript.Echo ""
WScript.Echo "=== FACE ATTRIBUTE APPROACH ==="

' Add a temporary attribute to mark the face
Dim attribSet
Set attribSet = largestFace.AttributeSets.Add("TempBaseFace")
If Err.Number = 0 Then
    WScript.Echo "Attribute added to largest face"
    
    ' Now try to find face by attribute
    Dim foundFace
    For Each face In body.Faces
        If face.AttributeSets.NameIsUsed("TempBaseFace") Then
            Set foundFace = face
            WScript.Echo "Found face with attribute!"
            Exit For
        End If
    Next
    
    ' Clean up
    largestFace.AttributeSets.Item("TempBaseFace").Delete
Else
    WScript.Echo "Attribute failed: " & Err.Description
    Err.Clear
End If

' Try the simple unfold and see if we can post-process
WScript.Echo ""
WScript.Echo "=== CREATE AND CHECK ==="

WScript.Echo "Creating flat pattern..."
compDef.Unfold
partDoc.Update

If compDef.HasFlatPattern Then
    Dim fp
    Set fp = compDef.FlatPattern
    WScript.Echo "Flat pattern: " & FormatNumber(fp.Length * 10, 1) & " x " & FormatNumber(fp.Width * 10, 1) & " mm"
    WScript.Echo "BaseFace area: " & FormatNumber(fp.BaseFace.Evaluator.Area * 100, 0) & " mm²"
End If

WScript.Echo ""
WScript.Echo "=== DONE ==="
