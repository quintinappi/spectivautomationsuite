' TEST_List_All_Feature_Collections.vbs
' List all available feature collections on PartFeatures
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, features

WScript.Echo "=== LIST ALL FEATURE COLLECTIONS ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
Set partDoc = invApp.ActiveDocument
Set compDef = partDoc.ComponentDefinition

WScript.Echo "Part: " & partDoc.DisplayName
WScript.Echo "ComponentDef Type: " & TypeName(compDef)
WScript.Echo "SubType: " & partDoc.SubType
WScript.Echo ""

Set features = compDef.Features
WScript.Echo "Features Type: " & TypeName(features)
WScript.Echo "Features Count: " & features.Count
WScript.Echo ""

' List each known feature collection type
WScript.Echo "=== KNOWN FEATURE COLLECTIONS ==="

Dim collections
collections = Array( _
    "ExtrudeFeatures", _
    "RevolveFeatures", _
    "SweepFeatures", _
    "LoftFeatures", _
    "HoleFeatures", _
    "FilletFeatures", _
    "ChamferFeatures", _
    "ShellFeatures", _
    "MirrorFeatures", _
    "PatternFeatures", _
    "SplitFeatures", _
    "ThickenFeatures", _
    "BendFeatures", _
    "FlangeFeatures", _
    "ContourFlangeFeatures", _
    "FoldFeatures", _
    "HemFeatures", _
    "CutFeatures", _
    "CornerFeatures", _
    "PunchToolFeatures", _
    "ConvertToSheetMetalFeatures", _
    "SheetMetalFeatures", _
    "FaceFeatures", _
    "CornerRoundFeatures", _
    "CornerChamferFeatures", _
    "ReferenceFeatures", _
    "DerivedPartFeatures", _
    "ClientFeatures" _
)

Dim collName, coll
For Each collName In collections
    On Error Resume Next
    Err.Clear
    
    Select Case collName
        Case "ExtrudeFeatures"
            Set coll = features.ExtrudeFeatures
        Case "RevolveFeatures"
            Set coll = features.RevolveFeatures
        Case "SweepFeatures"
            Set coll = features.SweepFeatures
        Case "LoftFeatures"
            Set coll = features.LoftFeatures
        Case "HoleFeatures"
            Set coll = features.HoleFeatures
        Case "FilletFeatures"
            Set coll = features.FilletFeatures
        Case "ChamferFeatures"
            Set coll = features.ChamferFeatures
        Case "ShellFeatures"
            Set coll = features.ShellFeatures
        Case "MirrorFeatures"
            Set coll = features.MirrorFeatures
        Case "SplitFeatures"
            Set coll = features.SplitFeatures
        Case "ThickenFeatures"
            Set coll = features.ThickenFeatures
        Case "BendFeatures"
            Set coll = features.BendFeatures
        Case "FlangeFeatures"
            Set coll = features.FlangeFeatures
        Case "ContourFlangeFeatures"
            Set coll = features.ContourFlangeFeatures
        Case "FoldFeatures"
            Set coll = features.FoldFeatures
        Case "HemFeatures"
            Set coll = features.HemFeatures
        Case "CutFeatures"
            Set coll = features.CutFeatures
        Case "CornerFeatures"
            Set coll = features.CornerFeatures
        Case "PunchToolFeatures"
            Set coll = features.PunchToolFeatures
        Case "ConvertToSheetMetalFeatures"
            Set coll = features.ConvertToSheetMetalFeatures
        Case "SheetMetalFeatures"
            Set coll = features.SheetMetalFeatures
        Case "FaceFeatures"
            Set coll = features.FaceFeatures
        Case "CornerRoundFeatures"
            Set coll = features.CornerRoundFeatures
        Case "CornerChamferFeatures"
            Set coll = features.CornerChamferFeatures
        Case "ReferenceFeatures"
            Set coll = features.ReferenceFeatures
        Case "DerivedPartFeatures"
            Set coll = features.DerivedPartFeatures
        Case "ClientFeatures"
            Set coll = features.ClientFeatures
        Case Else
            Set coll = Nothing
    End Select
    
    If Err.Number = 0 And Not coll Is Nothing Then
        Dim cnt
        cnt = coll.Count
        If Err.Number = 0 Then
            WScript.Echo "  [OK] " & collName & " (Count=" & cnt & ")"
        Else
            WScript.Echo "  [OK] " & collName & " (no count)"
            Err.Clear
        End If
    Else
        WScript.Echo "  [--] " & collName
        Err.Clear
    End If
    Set coll = Nothing
Next

WScript.Echo ""
WScript.Echo "=== DONE ==="
