' TEST_FlatPattern_Methods.vbs
' Enumerate ALL methods and properties of FlatPattern object
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, smDef, fp

WScript.Echo "=== FLATPATTERN METHODS AND PROPERTIES ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
Set partDoc = invApp.ActiveDocument
Set compDef = partDoc.ComponentDefinition
Set smDef = compDef

WScript.Echo "Part: " & partDoc.DisplayName

' Ensure flat pattern exists
If Not smDef.HasFlatPattern Then
    smDef.Unfold
    partDoc.Update
End If

Set fp = smDef.FlatPattern

WScript.Echo ""
WScript.Echo "=== KNOWN FLATPATTERN PROPERTIES ==="

WScript.Echo "Length: " & fp.Length
WScript.Echo "Width: " & fp.Width
WScript.Echo "Area: " & fp.Area
Err.Clear
WScript.Echo "BaseFace: " & TypeName(fp.BaseFace)
Err.Clear
WScript.Echo "Parent: " & TypeName(fp.Parent)
Err.Clear
WScript.Echo "Sketch: " & TypeName(fp.Sketch)
Err.Clear
WScript.Echo "TopFace: " & TypeName(fp.TopFace)
Err.Clear
WScript.Echo "BottomFace: " & TypeName(fp.BottomFace)
Err.Clear
WScript.Echo "Solid: " & TypeName(fp.Solid)
Err.Clear
WScript.Echo "SurfaceBody: " & TypeName(fp.SurfaceBody)
Err.Clear

' Check if there's a SetBaseFace with specific signature
WScript.Echo ""
WScript.Echo "=== TRYING SetBaseFace VARIATIONS ==="

Dim largestFace, body, face, largestArea
largestArea = 0
Set body = smDef.SurfaceBodies.Item(1)
For Each face In body.Faces
    Dim area
    area = face.Evaluator.Area * 100
    If area > largestArea Then
        largestArea = area
        Set largestFace = face
    End If
    Err.Clear
Next

WScript.Echo "Largest face: " & FormatNumber(largestArea, 0) & " mm²"

' Try SetBaseFace with different parameter counts
WScript.Echo ""
WScript.Echo "fp.SetBaseFace largestFace, True..."
fp.SetBaseFace largestFace, True
WScript.Echo "  Result: " & Err.Description
Err.Clear

WScript.Echo "fp.SetBaseFace largestFace, False..."
fp.SetBaseFace largestFace, False
WScript.Echo "  Result: " & Err.Description
Err.Clear

' Look at edit functionality
WScript.Echo ""
WScript.Echo "=== EDIT MODE ==="

WScript.Echo "Entering edit mode..."
fp.Edit
If Err.Number = 0 Then
    WScript.Echo "In edit mode!"
    
    ' Try to access more properties in edit mode
    WScript.Echo "fp.BaseFace in edit: " & TypeName(fp.BaseFace)
    Err.Clear
    
    WScript.Echo ""
    WScript.Echo "Looking for definition/parameters..."
    
    ' Try to get definition object
    Dim def
    Set def = fp.Definition
    WScript.Echo "fp.Definition: " & TypeName(def)
    Err.Clear
    
    ' Try parameters
    Dim params
    Set params = fp.Parameters
    WScript.Echo "fp.Parameters: " & TypeName(params)
    Err.Clear
    
    fp.ExitEdit
    WScript.Echo "Exited edit mode"
Else
    WScript.Echo "  Error: " & Err.Description
    Err.Clear
End If

' Look at the SheetMetalComponentDefinition for alternate approaches
WScript.Echo ""
WScript.Echo "=== SHEETMETAL COMPONENT DEF ==="

WScript.Echo "smDef.FlatPattern: " & TypeName(smDef.FlatPattern)

' Try creating flat pattern with definition
WScript.Echo ""
WScript.Echo "Deleting and recreating flat pattern..."

fp.Delete
partDoc.Update
WScript.Echo "Deleted"

' Check for CreateFlatPattern method with face parameter
WScript.Echo ""
WScript.Echo "smDef.CreateFlatPattern largestFace..."
smDef.CreateFlatPattern largestFace
WScript.Echo "  Result: " & Err.Description
Err.Clear

WScript.Echo "smDef.Unfold largestFace..."
smDef.Unfold largestFace
WScript.Echo "  Result: " & Err.Description
Err.Clear

' Just create simple flat pattern
WScript.Echo "smDef.Unfold (no params)..."
smDef.Unfold
WScript.Echo "  Result: " & Err.Description
If Err.Number = 0 Then
    WScript.Echo "  Created!"
End If
Err.Clear

partDoc.Update

' Final check
WScript.Echo ""
WScript.Echo "=== FINAL ==="
If smDef.HasFlatPattern Then
    Set fp = smDef.FlatPattern
    WScript.Echo "Flat pattern: " & FormatNumber(fp.Length * 10, 1) & " x " & FormatNumber(fp.Width * 10, 1) & " mm"
End If

WScript.Echo ""
WScript.Echo "=== DONE ==="
