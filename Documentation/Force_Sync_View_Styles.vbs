' Force_Sync_View_Styles_V2.vbs
' Improved version with better error reporting and aggressive updates.

Option Explicit

Dim invApp, doc
On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Or invApp Is Nothing Then
    WScript.Echo "Error: Inventor is not running!"
    WScript.Quit
End If

Set doc = invApp.ActiveDocument
If doc Is Nothing Or doc.DocumentType <> 12292 Then
    WScript.Echo "Error: Active document is not an IDW!"
    WScript.Quit
End If

' 1. Select Target Standard
Dim stylesMgr, standards, choice, i
Set stylesMgr = doc.StylesManager
Set standards = stylesMgr.StandardStyles

Dim listText
listText = "AVAILABLE STANDARDS:" & vbCrLf
For i = 1 To standards.Count
    listText = listText & i & ": " & standards.Item(i).Name & vbCrLf
Next

WScript.Echo listText
choice = InputBox(listText & vbCrLf & "Enter number to apply:", "Force Sync Styles", "1")

If choice = "" Then WScript.Quit
If Not IsNumeric(choice) Then WScript.Quit

Dim targetStd
Set targetStd = standards.Item(CInt(choice))
WScript.Echo "Applying Target Standard: " & targetStd.Name

' 2. Switch Document Standard
stylesMgr.ActiveStandardStyle = targetStd

' Get Layers from Object Defaults
Dim objDef
Set objDef = targetStd.ActiveObjectDefaults
Dim visLayer, hidLayer, tanLayer
Set visLayer = objDef.VisibleEdgeLayer
Set hidLayer = objDef.HiddenEdgeLayer
Set tanLayer = objDef.TangentEdgeLayer

WScript.Echo "Target Visible Layer: " & visLayer.Name
WScript.Echo "Target Hidden Layer: " & hidLayer.Name

' 3. Process Views
Dim sheet, view, curve, segment, count, segmentCount
count = 0
segmentCount = 0

For Each sheet In doc.Sheets
    WScript.Echo "Sheet: " & sheet.Name
    For Each view In sheet.DrawingViews
        WScript.Echo "  Updating View: " & view.Name
        
        ' Clear any "Style" override by cycling ViewStyle
        Dim currentVS
        currentVS = view.ViewStyle
        view.ViewStyle = 32257 ' kFromBaseDrawingViewStyle
        
        ' Iterate curves to force layers
        On Error Resume Next
        Dim curves
        Set curves = view.DrawingCurves
        If Err.Number = 0 Then
            For Each curve In curves
                Dim targetL
                Set targetL = Nothing
                
                ' kVisibleEdge = 32257, kHiddenEdge = 32258, kTangentEdge = 32259
                Select Case curve.EdgeType
                    Case 32257: Set targetL = visLayer
                    Case 32258: Set targetL = hidLayer
                    Case 32259: Set targetL = tanLayer
                End Select
                
                If Not targetL Is Nothing Then
                    For Each segment In curve.Segments
                        segment.Layer = targetL
                        segmentCount = segmentCount + 1
                    Next
                End If
            Next
        Else
            WScript.Echo "    Error accessing curves: " & Err.Description
        End If
        Err.Clear
        
        count = count + 1
    Next
Next

WScript.Echo "Force synced " & count & " views."
WScript.Echo "Updated " & segmentCount & " curve segments."

' 4. Final Rebuild
doc.Update
doc.Update2(True) ' Force deep update

WScript.Echo "Process Complete."
MsgBox "Updated " & count & " views and " & segmentCount & " lines." & vbCrLf & "Successfully applied '" & targetStd.Name & "'.", vbInformation
