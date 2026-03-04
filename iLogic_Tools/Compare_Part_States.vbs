On Error Resume Next

Dim invApp
Set invApp = GetObject(, "Inventor.Application")
Dim doc
Set doc = invApp.ActiveDocument

WScript.Echo "=== PART STATE COMPARISON ==="
WScript.Echo "Part: " & doc.DisplayName
WScript.Echo ""

WScript.Echo "Document Properties:"
WScript.Echo "  Full Path: " & doc.FullFileName
WScript.Echo "  Is Dirty: " & doc.Dirty
WScript.Echo "  Document Type: " & doc.DocumentType
WScript.Echo "  SubType: " & doc.SubType
WScript.Echo "  File Save Counter: " & doc.FileSaveCounter
WScript.Echo ""

Dim compDef
Set compDef = doc.ComponentDefinition

WScript.Echo "Component Definition:"
WScript.Echo "  SubType: " & compDef.DocumentSubType.UniqueID
WScript.Echo "  Surface Bodies Count: " & compDef.SurfaceBodies.Count
WScript.Echo ""

If compDef.SurfaceBodies.Count > 0 Then
    Dim body
    Set body = compDef.SurfaceBodies.Item(1)
    
    WScript.Echo "Surface Body 1:"
    WScript.Echo "  Type: " & TypeName(body)
    WScript.Echo "  Faces Count: " & body.Faces.Count
    WScript.Echo ""
    
    Dim face
    Set face = body.Faces.Item(5)
    
    WScript.Echo "Face 5 (large face):"
    WScript.Echo "  Type: " & TypeName(face)
    WScript.Echo "  Surface Type: " & face.SurfaceType
    WScript.Echo "  Is Parametric: " & face.IsParametric
    
    ' Check if it's a proxy object
    On Error Resume Next
    Dim obj
    Set obj = face.ContainingOccurrence
    If Err.Number = 0 And Not obj Is Nothing Then
        WScript.Echo "  Is from Occurrence: YES (this is a proxy!)"
        WScript.Echo "  Occurrence: " & obj.Name
    Else
        WScript.Echo "  Is from Occurrence: NO (direct face)"
    End If
    Err.Clear
End If

WScript.Echo ""
WScript.Echo "=== TRYING SELECT WITH MODEL STATE UPDATE ==="

doc.Update

Dim selectSet
Set selectSet = doc.SelectSet
selectSet.Clear

WScript.Echo "Attempting to select face 5..."

Dim largestFace
Set largestFace = compDef.SurfaceBodies.Item(1).Faces.Item(5)

Err.Clear
selectSet.Select largestFace

WScript.Echo "Error: " & Err.Number & " - " & Err.Description
WScript.Echo "SelectSet.Count: " & selectSet.Count
