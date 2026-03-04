' Layer_Style_Manager.vbs
' Lists all available layers and lets you apply a chosen layer to selected views
' Similar to Change_Document_Standard but for Layers instead of Standards

Option Explicit

Dim invApp, doc, activeSheet
Dim layersManager, allLayers
Dim layerList, i, layer
Dim selectedLayerIndex, selectedLayer
Dim targetViewInput, targetView
Dim viewsUpdated, curvesUpdated

On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Or invApp Is Nothing Then
    WScript.Echo "ERROR: Inventor is not running!"
    WScript.Quit 1
End If

Set doc = invApp.ActiveDocument
If doc Is Nothing Or doc.DocumentType <> 12292 Then
    WScript.Echo "ERROR: No IDW file is open!"
    WScript.Quit 1
End If

Set activeSheet = doc.ActiveSheet

WScript.Echo "=========================================="
WScript.Echo "  LAYER STYLE MANAGER"
WScript.Echo "=========================================="
WScript.Echo "Drawing: " & doc.DisplayName
WScript.Echo "Active Sheet: " & activeSheet.Name
WScript.Echo ""

' Get styles manager and layers
Set layersManager = doc.StylesManager
Set allLayers = layersManager.Layers

If allLayers Is Nothing Or allLayers.Count = 0 Then
    WScript.Echo "ERROR: No layers found in document!"
    WScript.Quit 1
End If

' Build layer list
WScript.Echo "AVAILABLE LAYERS:"
WScript.Echo ""

For i = 1 To allLayers.Count
    Set layer = allLayers.Item(i)
    WScript.Echo i & ". " & layer.Name
Next

WScript.Echo ""
WScript.Echo "Enter the NUMBER of the layer you want to APPLY:"
WScript.Echo ""

Dim layerChoice
layerChoice = InputBox(_
    "Enter the NUMBER of the layer to apply to views:" & vbCrLf & vbCrLf & _
    "Available layers: " & allLayers.Count & vbCrLf & vbCrLf & _
    "Tip: Choose a layer like 'PEN25 Centerline (ISO)' or 'PEN25 Hidden (ISO)'", _
    "Select Layer", "")

If layerChoice = "" Then
    WScript.Echo "Cancelled."
    WScript.Quit
End If

If Not IsNumeric(layerChoice) Then
    WScript.Echo "ERROR: Please enter a number!"
    WScript.Quit 1
End If

Dim layerIndex
layerIndex = CInt(layerChoice)

If layerIndex < 1 Or layerIndex > allLayers.Count Then
    WScript.Echo "ERROR: Invalid layer number!"
    WScript.Quit 1
End If

Set selectedLayer = allLayers.Item(layerIndex)
WScript.Echo ""
WScript.Echo "Selected Layer: " & selectedLayer.Name
WScript.Echo ""

' Show available views
WScript.Echo "VIEWS ON ACTIVE SHEET:"
For Each targetView In activeSheet.DrawingViews
    WScript.Echo "  - " & targetView.Name
Next
WScript.Echo ""

' Ask which views to apply to
targetViewInput = InputBox(_
    "Enter view names to apply layer to (comma-separated):" & vbCrLf & vbCrLf & _
    "Examples:" & vbCrLf & _
    "  1, 2, 3  (specific views)" & vbCrLf & _
    "  A, B, C  (lettered views)" & vbCrLf & _
    "  *        (ALL views on sheet)", _
    "Select Target Views", "*")

If targetViewInput = "" Then
    WScript.Echo "Cancelled."
    WScript.Quit
End If

targetViewInput = Trim(targetViewInput)

' Process views
WScript.Echo ""
WScript.Echo "=========================================="
WScript.Echo "  APPLYING LAYER: " & selectedLayer.Name
WScript.Echo "=========================================="
WScript.Echo ""

viewsUpdated = 0
curvesUpdated = 0

Dim viewNames, vName, v
viewNames = Split(targetViewInput, ",")

For Each vName In viewNames
    vName = Trim(vName)
    If vName <> "" Then
    
    ' Find the view
    For Each targetView In activeSheet.DrawingViews
        If UCase(targetView.Name) = UCase(vName) OR targetViewInput = "*" Then
            WScript.Echo "Processing view: " & targetView.Name
            
            ' Apply layer to all curves in this view
            Dim curves, c, s
            Set curves = targetView.DrawingCurves
            
            Dim viewCurvesCount
            viewCurvesCount = 0
            
            For Each c In curves
                For Each s In c.Segments
                    s.Layer = selectedLayer
                    viewCurvesCount = viewCurvesCount + 1
                Next
            Next
            
            WScript.Echo "  -> Updated " & viewCurvesCount & " segments"
            curvesUpdated = curvesUpdated + viewCurvesCount
            viewsUpdated = viewsUpdated + 1
            
            ' If we matched by specific name, exit loop
            If targetViewInput <> "*" Then Exit For
        End If
    Next
    End If
Next

WScript.Echo ""
WScript.Echo "=========================================="
WScript.Echo "  SUMMARY"
WScript.Echo "=========================================="
WScript.Echo "Views updated: " & viewsUpdated
WScript.Echo "Total segments updated: " & curvesUpdated
WScript.Echo ""

If viewsUpdated > 0 Then
    WScript.Echo "Updating document..."
    doc.Update
    
    WScript.Echo "Saving document..."
    doc.Save2
    
    If Err.Number = 0 Then
        WScript.Echo "SUCCESS! Document saved."
    Else
        WScript.Echo "ERROR saving: " & Err.Description
    End If
Else
    WScript.Echo "No views were updated."
End If

WScript.Echo ""
WScript.Echo "Done!"
