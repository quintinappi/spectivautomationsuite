' TEST_FlatPattern_Features.vbs
' Look at FlatPatternFeatures collection and definitions
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, smDef, fp

WScript.Echo "=== FLATPATTERN FEATURES AND DEFINITIONS ==="
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

WScript.Echo "Current: " & FormatNumber(fp.Length * 10, 1) & " x " & FormatNumber(fp.Width * 10, 1) & " mm"
WScript.Echo "BaseFace area: " & FormatNumber(fp.BaseFace.Evaluator.Area * 100, 0) & " mm²"
WScript.Echo "TopFace area: " & FormatNumber(fp.TopFace.Evaluator.Area * 100, 0) & " mm²"
Err.Clear
WScript.Echo "BottomFace: " & TypeName(fp.BottomFace)
If Err.Number = 0 And Not fp.BottomFace Is Nothing Then
    WScript.Echo "BottomFace area: " & FormatNumber(fp.BottomFace.Evaluator.Area * 100, 0) & " mm²"
End If
Err.Clear

' Check FlatPatternFeatures collection
WScript.Echo ""
WScript.Echo "=== smDef.Features.FlatPatternFeatures ==="

Dim fpFeatures
Set fpFeatures = smDef.Features.FlatPatternFeatures
If Err.Number = 0 And Not fpFeatures Is Nothing Then
    WScript.Echo "Found! Count: " & fpFeatures.Count
    
    Dim fpFeat, i
    i = 0
    For Each fpFeat In fpFeatures
        i = i + 1
        WScript.Echo ""
        WScript.Echo "FlatPatternFeature " & i & ":"
        WScript.Echo "  Name: " & fpFeat.Name
        WScript.Echo "  Type: " & TypeName(fpFeat)
        
        ' Try Definition
        Dim def
        Set def = fpFeat.Definition
        WScript.Echo "  Definition: " & TypeName(def)
        Err.Clear
        
        If Not def Is Nothing Then
            ' Try properties of definition
            WScript.Echo "    def.BaseFace: " & TypeName(def.BaseFace)
            Err.Clear
            WScript.Echo "    def.StaticFace: " & TypeName(def.StaticFace)
            Err.Clear
        End If
    Next
Else
    WScript.Echo "Not found: " & Err.Description
    Err.Clear
End If

' Check for FlatPatternDefinition
WScript.Echo ""
WScript.Echo "=== SheetMetalFeatures ==="

Dim smFeatures
Set smFeatures = smDef.Features.SheetMetalFeatures
If Err.Number = 0 And Not smFeatures Is Nothing Then
    WScript.Echo "Found SheetMetalFeatures! Count: " & smFeatures.Count
Else
    WScript.Echo "Not found: " & Err.Description
    Err.Clear
End If

' Try to find FlatPatternDefinition through Features
WScript.Echo ""
WScript.Echo "=== All Features ==="

Dim feat
For Each feat In smDef.Features
    WScript.Echo "  " & feat.Name & " (" & TypeName(feat) & ")"
    
    ' Check if this feature has a definition with BaseFace
    Dim featDef
    Set featDef = feat.Definition
    If Not featDef Is Nothing Then
        Dim bf
        Set bf = featDef.BaseFace
        If Err.Number = 0 And Not bf Is Nothing Then
            WScript.Echo "    -> Has BaseFace! Area: " & FormatNumber(bf.Evaluator.Area * 100, 0) & " mm²"
        End If
        Err.Clear
    End If
    Err.Clear
Next

' Look at Flat Pattern Extents
WScript.Echo ""
WScript.Echo "=== fp.Extents ==="
Dim extents
Set extents = fp.Extents
If Not extents Is Nothing Then
    WScript.Echo "Found! Type: " & TypeName(extents)
End If
Err.Clear

' Look at sketch
WScript.Echo ""
WScript.Echo "=== fp.Sketch ==="
Dim sk
Set sk = fp.Sketch
If Not sk Is Nothing Then
    WScript.Echo "Found! " & sk.Name
End If
Err.Clear

' Find largest face
Dim body, face, largestFace, largestArea
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

' Try invoking through command with face selection
WScript.Echo ""
WScript.Echo "=== TRYING COMMAND WITH FACE PRE-SELECTED ==="

Dim selectSet, cmdMgr
Set selectSet = partDoc.SelectSet
Set cmdMgr = invApp.CommandManager

' Delete flat pattern first
fp.Delete
partDoc.Update
WScript.Echo "Flat pattern deleted"

' Select largest face
selectSet.Clear
selectSet.Select largestFace
WScript.Echo "Largest face selected"

' Find and execute the unfold/flatten command
Dim unfoldCmds, cmdName, cmd
unfoldCmds = Array("PartUnfoldCmd", "SheetMetalUnfoldCmd", "SMUnfoldCmd", _
                   "PartCreateFlatPatternCmd", "SheetMetalCreateFlatPatternCmd", _
                   "PartFlatPatternCmd", "CreateFlatPatternCmd")

For Each cmdName In unfoldCmds
    Set cmd = cmdMgr.ControlDefinitions.Item(cmdName)
    If Not cmd Is Nothing Then
        WScript.Echo cmdName & " - Enabled: " & cmd.Enabled
    End If
    Err.Clear
Next

WScript.Echo ""
WScript.Echo "=== DONE ==="
