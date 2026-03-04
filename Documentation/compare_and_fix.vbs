' Compare_Correct_Vs_Wrong.vbs
' Specifically compares VIEW1 (correct) with ELEVATION (wrong) on Sheet 2

Option Explicit

Dim invApp, doc, sheet, vCorrect, vWrong
On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")
Set doc = invApp.ActiveDocument
Set sheet = doc.Sheets.Item(2) ' Sheet 2

' Find views by name
For Each v In sheet.DrawingViews
    If UCase(v.Name) = "VIEW1" Then Set vCorrect = v
    If UCase(v.Name) = "ELEVATION" Then Set vWrong = v
Next

If vCorrect Is Nothing Then WScript.Echo "Could not find VIEW1": WScript.Quit
If vWrong Is Nothing Then WScript.Echo "Could not find ELEVATION": WScript.Quit

WScript.Echo "Comparison: VIEW1 (Correct) vs ELEVATION (Wrong)"
WScript.Echo "==============================================="

WScript.Echo "Property | VIEW1 | ELEVATION"
WScript.Echo "---------|-------|----------"

' Helper to get property safely
Function GetProp(obj, propName)
    On Error Resume Next
    Err.Clear
    Dim val
    val = CallByName(obj, propName, 1) ' VBScript doesn't have CallByName, need another way
    If Err.Number <> 0 Then
        GetProp = "Error: " & Err.Description
    Else
        GetProp = val
    End If
End Function

' Manual property check
On Error Resume Next
WScript.Echo "ViewStyle | " & vCorrect.ViewStyle & " | " & vWrong.ViewStyle
WScript.Echo "Scale | " & vCorrect.Scale & " | " & vWrong.Scale
WScript.Echo "Suppress | " & vCorrect.IsSuppressed & " | " & vWrong.IsSuppressed
WScript.Echo "FromStandard | " & vCorrect.IsStyleDefault & " | " & vWrong.IsStyleDefault

' Check Layer of a Hidden Curve
Dim curvesC, curvesW
Set curvesC = vCorrect.DrawingCurves
Set curvesW = vWrong.DrawingCurves

WScript.Echo "Curve Count | " & curvesC.Count & " | " & curvesW.Count

' Try to find hidden layer name
Dim hLayerC, hLayerW
hLayerC = "None"
hLayerW = "None"

For Each c In curvesC
    If c.EdgeType = 32258 Then ' kHiddenEdge
        If c.Segments.Count > 0 Then
            hLayerC = c.Segments.Item(1).Layer.Name
            Exit For
        End If
    End If
Next

For Each c In curvesW
    If c.EdgeType = 32258 Then ' kHiddenEdge
        If c.Segments.Count > 0 Then
            hLayerW = c.Segments.Item(1).Layer.Name
            Exit For
        End If
    End If
Next

WScript.Echo "Hidden Layer | " & hLayerC & " | " & hLayerW

WScript.Echo ""
WScript.Echo "Attempting to FORCE ELEVATION to use VIEW1's Hidden Layer..."

If hLayerC <> "None" Then
    Dim targetLayer
    Set targetLayer = doc.StylesManager.Layers.Item(hLayerC)
    
    If Not targetLayer Is Nothing Then
        WScript.Echo "Found target layer: " & targetLayer.Name
        Dim count
        count = 0
        For Each c In curvesW
            If c.EdgeType = 32258 Then ' Hidden
                For Each s In c.Segments
                    s.Layer = targetLayer
                    count = count + 1
                Next
            End If
        Next
        WScript.Echo "Updated " & count & " segments in ELEVATION."
    End If
End If

doc.Update
WScript.Echo "Done."
