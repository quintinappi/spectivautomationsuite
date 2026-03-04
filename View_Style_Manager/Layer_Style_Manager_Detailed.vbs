' Layer_Style_Manager_Detailed.vbs
' Advanced version - lets you choose specific layers for:
' - Visible lines
' - Hidden lines  
' - Center lines
' - Center marks
' Then applies them to selected views based on current layer names

Option Explicit

Dim invApp, doc, activeSheet
Dim allLayers
Dim i, layer
Dim masterVisLayer, masterHidLayer, masterCenterLineLayer, masterCenterMarkLayer
Dim targetViewInput
Dim viewsUpdated, visCount, hidCount, clCount, cmCount

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
WScript.Echo "  LAYER STYLE MANAGER - DETAILED"
WScript.Echo "=========================================="
WScript.Echo "Drawing: " & doc.DisplayName
WScript.Echo "Active Sheet: " & activeSheet.Name
WScript.Echo ""

' Get all layers
Set allLayers = doc.StylesManager.Layers

If allLayers Is Nothing Or allLayers.Count = 0 Then
    WScript.Echo "ERROR: No layers found!"
    WScript.Quit 1
End If

' Build list of layers by category
Dim visLayers, hidLayers, centerLayers, otherLayers
Set visLayers = CreateObject("Scripting.Dictionary")
Set hidLayers = CreateObject("Scripting.Dictionary")
Set centerLayers = CreateObject("Scripting.Dictionary")
Set otherLayers = CreateObject("Scripting.Dictionary")

For i = 1 To allLayers.Count
    Set layer = allLayers.Item(i)
    Dim lName
    lName = UCase(layer.Name)
    
    If InStr(lName, "HIDDEN") > 0 Then
        hidLayers.Add layer.Name, layer
    ElseIf InStr(lName, "CENTER") > 0 Then
        centerLayers.Add layer.Name, layer
    ElseIf InStr(lName, "VISIBLE") > 0 OR InStr(lName, "PEN") > 0 Then
        visLayers.Add layer.Name, layer
    Else
        otherLayers.Add layer.Name, layer
    End If
Next

' Show categorized lists and let user choose
WScript.Echo "CONFIGURE LAYERS FOR EACH TYPE:"
WScript.Echo ""

' 1. Choose Visible Layer
WScript.Echo "--- VISIBLE LINES ---"
If visLayers.Count > 0 Then
    Dim visKeys
    visKeys = visLayers.Keys
    For i = 0 To UBound(visKeys)
        WScript.Echo (i+1) & ". " & visKeys(i)
    Next
    
    Dim visChoice
    visChoice = InputBox("Enter NUMBER for VISIBLE layer:", "Visible Layer", "1")
    If IsNumeric(visChoice) Then
        Set masterVisLayer = visLayers.Item(visKeys(CInt(visChoice)-1))
    End If
End If

If masterVisLayer Is Nothing Then
    ' Let user pick from all layers
    WScript.Echo "Pick from all " & allLayers.Count & " layers:"
    For i = 1 To allLayers.Count
        WScript.Echo i & ". " & allLayers.Item(i).Name
    Next
    
    visChoice = InputBox("Enter NUMBER for VISIBLE layer:", "Visible Layer", "1")
    If IsNumeric(visChoice) Then
        Set masterVisLayer = allLayers.Item(CInt(visChoice))
    End If
End If

If masterVisLayer Is Nothing Then
    WScript.Echo "ERROR: No visible layer selected!"
    WScript.Quit 1
End If

WScript.Echo "Selected Visible Layer: " & masterVisLayer.Name
WScript.Echo ""

' 2. Choose Hidden Layer
WScript.Echo "--- HIDDEN LINES ---"
If hidLayers.Count > 0 Then
    Dim hidKeys
    hidKeys = hidLayers.Keys
    For i = 0 To UBound(hidKeys)
        WScript.Echo (i+1) & ". " & hidKeys(i)
    Next
    
    Dim hidChoice
    hidChoice = InputBox("Enter NUMBER for HIDDEN layer:", "Hidden Layer", "1")
    If IsNumeric(hidChoice) Then
        Set masterHidLayer = hidLayers.Item(hidKeys(CInt(hidChoice)-1))
    End If
End If

If masterHidLayer Is Nothing Then
    ' Pick from all
    For i = 1 To allLayers.Count
        WScript.Echo i & ". " & allLayers.Item(i).Name
    Next
    
    hidChoice = InputBox("Enter NUMBER for HIDDEN layer:", "Hidden Layer", "1")
    If IsNumeric(hidChoice) Then
        Set masterHidLayer = allLayers.Item(CInt(hidChoice))
    End If
End If

If masterHidLayer Is Nothing Then
    WScript.Echo "ERROR: No hidden layer selected!"
    WScript.Quit 1
End If

WScript.Echo "Selected Hidden Layer: " & masterHidLayer.Name
WScript.Echo ""

' 3. Choose Center Line Layer (optional)
WScript.Echo "--- CENTER LINES (Optional) ---"
If centerLayers.Count > 0 Then
    Dim centerKeys
    centerKeys = centerLayers.Keys
    For i = 0 To UBound(centerKeys)
        WScript.Echo (i+1) & ". " & centerKeys(i)
    Next
    WScript.Echo (centerLayers.Count+1) & ". [SKIP - Don't change center lines]"
    
    Dim clChoice
    clChoice = InputBox("Enter NUMBER for CENTER LINE layer (or SKIP):", "Center Line Layer", "1")
    If IsNumeric(clChoice) Then
        Dim clIndex
        clIndex = CInt(clChoice) - 1
        If clIndex >= 0 And clIndex <= UBound(centerKeys) Then
            Set masterCenterLineLayer = centerLayers.Item(centerKeys(clIndex))
        End If
    End If
End If

If Not masterCenterLineLayer Is Nothing Then
    WScript.Echo "Selected Center Line Layer: " & masterCenterLineLayer.Name
Else
    WScript.Echo "Center Lines: SKIPPED"
End If
WScript.Echo ""

' 4. Choose Center Mark Layer (optional)
WScript.Echo "--- CENTER MARKS (Optional) ---"
If centerLayers.Count > 0 Then
    centerKeys = centerLayers.Keys
    For i = 0 To UBound(centerKeys)
        WScript.Echo (i+1) & ". " & centerKeys(i)
    Next
    WScript.Echo (centerLayers.Count+1) & ". [SKIP - Don't change center marks]"
    
    Dim cmChoice
    cmChoice = InputBox("Enter NUMBER for CENTER MARK layer (or SKIP):", "Center Mark Layer", "1")
    If IsNumeric(cmChoice) Then
        Dim cmIndex
        cmIndex = CInt(cmChoice) - 1
        If cmIndex >= 0 And cmIndex <= UBound(centerKeys) Then
            Set masterCenterMarkLayer = centerLayers.Item(centerKeys(cmIndex))
        End If
    End If
End If

If Not masterCenterMarkLayer Is Nothing Then
    WScript.Echo "Selected Center Mark Layer: " & masterCenterMarkLayer.Name
Else
    WScript.Echo "Center Marks: SKIPPED"
End If
WScript.Echo ""

' Show views and ask which to process
WScript.Echo "VIEWS ON ACTIVE SHEET:"
Dim v
For Each v In activeSheet.DrawingViews
    WScript.Echo "  - " & v.Name
Next
WScript.Echo ""

targetViewInput = InputBox(_
    "Enter view names to apply layers to:" & vbCrLf & _
    "  * = ALL views" & vbCrLf & _
    "  Or comma-separated names like: 1, 2, A, B", _
    "Select Views", "*")

If targetViewInput = "" Then
    WScript.Echo "Cancelled."
    WScript.Quit
End If

' Apply to views
WScript.Echo ""
WScript.Echo "=========================================="
WScript.Echo "  APPLYING LAYERS"
WScript.Echo "=========================================="
WScript.Echo ""

viewsUpdated = 0
visCount = 0
hidCount = 0
clCount = 0
cmCount = 0

Dim viewNames, vName, targetView
viewNames = Split(targetViewInput, ",")

For Each vName In viewNames
    vName = Trim(vName)
    If vName <> "" Then
    
    For Each targetView In activeSheet.DrawingViews
        If UCase(targetView.Name) = UCase(vName) OR targetViewInput = "*" Then
            WScript.Echo "Processing: " & targetView.Name
            
            Dim curves, c, s, curLayer
            Set curves = targetView.DrawingCurves
            
            For Each c In curves
                For Each s In c.Segments
                    Set curLayer = s.Layer
                    Dim curName
                    curName = UCase(curLayer.Name)
                    
                    ' Route to correct layer based on current layer name
                    If InStr(curName, "HIDDEN") > 0 Then
                        s.Layer = masterHidLayer
                        hidCount = hidCount + 1
                    ElseIf InStr(curName, "CENTER MARK") > 0 OR InStr(curName, "CENTERMARK") > 0 Then
                        If Not masterCenterMarkLayer Is Nothing Then
                            s.Layer = masterCenterMarkLayer
                            cmCount = cmCount + 1
                        End If
                    ElseIf InStr(curName, "CENTER") > 0 Then
                        If Not masterCenterLineLayer Is Nothing Then
                            s.Layer = masterCenterLineLayer
                            clCount = clCount + 1
                        End If
                    Else
                        s.Layer = masterVisLayer
                        visCount = visCount + 1
                    End If
                Next
            Next
            
            viewsUpdated = viewsUpdated + 1
            
            If targetViewInput <> "*" Then Exit For
        End If
    Next
    End If
Next

' Summary
WScript.Echo ""
WScript.Echo "=========================================="
WScript.Echo "  SUMMARY"
WScript.Echo "=========================================="
WScript.Echo "Views updated: " & viewsUpdated
WScript.Echo "Visible lines updated: " & visCount
WScript.Echo "Hidden lines updated: " & hidCount
If Not masterCenterLineLayer Is Nothing Then
    WScript.Echo "Center lines updated: " & clCount
End If
If Not masterCenterMarkLayer Is Nothing Then
    WScript.Echo "Center marks updated: " & cmCount
End If
WScript.Echo ""

If viewsUpdated > 0 Then
    doc.Update
    doc.Save2
    WScript.Echo "SUCCESS! Document saved."
Else
    WScript.Echo "No views were updated."
End If

WScript.Echo ""
WScript.Echo "Done!"
