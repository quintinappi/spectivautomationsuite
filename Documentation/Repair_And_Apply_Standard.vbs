' Repair_And_Apply_Standard.vbs
' Reads the Active Standard's Object Defaults and forces those layers onto the view curves.
' This fixes the "All Cyan" mess and correctly applies the standard.

Option Explicit

Dim invApp, doc
On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")
Set doc = invApp.ActiveDocument

WScript.Echo "=== REPAIR AND APPLY STANDARD ==="

' 1. GET ACTIVE STANDARD DEFAULTS
Dim stylesMgr, activeStd, objDefs
Set stylesMgr = doc.StylesManager
Set activeStd = stylesMgr.ActiveStandardStyle
Set objDefs = activeStd.ActiveObjectDefaults

WScript.Echo "Active Standard: " & activeStd.Name
WScript.Echo "Object Defaults: " & objDefs.Name

' Get the correct layers from the standard
Dim layerVisible, layerHidden, layerTangent
Set layerVisible = objDefs.VisibleEdgeLayer
Set layerHidden = objDefs.HiddenEdgeLayer
Set layerTangent = objDefs.TangentEdgeLayer

WScript.Echo "Target Layers:"
WScript.Echo " - Visible: " & layerVisible.Name
WScript.Echo " - Hidden:  " & layerHidden.Name
WScript.Echo " - Tangent: " & layerTangent.Name

' 2. REPAIR VIEWS
Dim sheet
Set sheet = doc.Sheets.Item(2) ' Sheet 2

Dim targetNames, viewName
targetNames = Array("ELEVATION", "C", "D", "PLAN")

For Each viewName In targetNames
    Dim targetView
    Set targetView = Nothing
    
    ' Find view
    Dim v
    For Each v In sheet.DrawingViews
        If InStr(UCase(v.Name), viewName) > 0 Then
            Set targetView = v
            Exit For
        End If
    Next
    
    If Not targetView Is Nothing Then
        WScript.Echo "Repairing View: " & targetView.Name
        
        Dim curves
        Set curves = targetView.DrawingCurves
        
        Dim c, s
        Dim countVis, countHid, countTan, countOther
        countVis = 0
        countHid = 0
        countTan = 0
        countOther = 0
        
        For Each c In curves
            Dim targetLayer
            Set targetLayer = Nothing
            
            ' Map Edge Types to Correct Standard Layers
            Select Case c.EdgeType
                Case 32257 ' kVisibleEdge
                    Set targetLayer = layerVisible
                    countVis = countVis + 1
                Case 32258 ' kHiddenEdge
                    Set targetLayer = layerHidden
                    countHid = countHid + 1
                Case 32259 ' kTangentEdge
                    Set targetLayer = layerTangent
                    countTan = countTan + 1
                Case 82695 ' The mystery type - usually Tangent or Bend
                    Set targetLayer = layerTangent
                    countOther = countOther + 1
                Case Else
                    ' Default to Visible for unknown visible geometry
                    Set targetLayer = layerVisible
            End Select
            
            If Not targetLayer Is Nothing Then
                For Each s In c.Segments
                    s.Layer = targetLayer
                Next
            End If
        Next
        
        WScript.Echo "  -> Reset " & countVis & " Visible lines."
        WScript.Echo "  -> Reset " & countHid & " Hidden lines."
        WScript.Echo "  -> Reset " & countTan & " Tangent lines."
        WScript.Echo "  -> Reset " & countOther & " Other lines."
    End If
Next

doc.Update
WScript.Echo "Done. Views should now match the '" & activeStd.Name & "' standard."
