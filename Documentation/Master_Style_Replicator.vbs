' Master_Style_Replicator.vbs
' Interactive tool to copy view styling from a user-specified Master View to Target Views.
' NOW INCLUDES: Center Line and Center Mark layer detection from geometry!
'
' USAGE:
' 1. Run script with Inventor open and an IDW active
' 2. View all available views on the active sheet
' 3. Select which view is the Master View (to copy style FROM)
' 4. Choose target views to apply style TO:
'    - ALL views (except Master)
'    - Specific views (comma-separated list)
'    - Views matching a pattern (wildcards)

Option Explicit

' Declare ALL variables at script level (no redeclarations)
Dim invApp, doc, activeSheet
Dim allViewsCollection, targetViewsCollection
Dim masterViewName, masterView, masterSheet
Dim targetPattern, specificNames, i, choice
Dim keys, targetView, totalCount, visCount, hidCount
Dim sheet, v, l, c, s, curves, tCurves
Dim masterVisLayer, masterHidLayer, masterCenterLineLayer, masterCenterMarkLayer
Dim countViewVis, countViewHid
Dim viewName, namesArray, curLayer, vName
Dim specificNamesInput, sampleCount, colorDict, colorKey
Dim centerLineCount, centerMarkCount

On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Or invApp Is Nothing Then
    WScript.Echo "ERROR: Inventor is not running!"
    WScript.Quit 1
End If

Set doc = invApp.ActiveDocument
If doc Is Nothing Then
    WScript.Echo "ERROR: No active document!"
    WScript.Quit 1
End If

If doc.DocumentType <> 12292 Then
    WScript.Echo "ERROR: Active document is not an IDW!"
    WScript.Quit 1
End If

' Get the active sheet only
Set activeSheet = doc.ActiveSheet

WScript.Echo "=========================================="
WScript.Echo "  MASTER STYLE REPLICATOR"
WScript.Echo "=========================================="
WScript.Echo ""

' ============================================================================
' STEP 1: SCAN AND LIST ALL AVAILABLE VIEWS ON ACTIVE SHEET
' ============================================================================

WScript.Echo "=========================================="
WScript.Echo "  AVAILABLE VIEWS ON ACTIVE SHEET"
WScript.Echo "=========================================="
WScript.Echo "Sheet: " & activeSheet.Name
WScript.Echo ""

Set allViewsCollection = CreateObject("Scripting.Dictionary")

For Each v In activeSheet.DrawingViews
    If Not allViewsCollection.Exists(v.Name) Then
        allViewsCollection.Add v.Name, v
    End If
    WScript.Echo "  - " & v.Name
Next

WScript.Echo ""
WScript.Echo "=========================================="
WScript.Echo "  Found " & allViewsCollection.Count & " view(s) on active sheet"
WScript.Echo "=========================================="
WScript.Echo ""

' ============================================================================
' STEP 2: ASK USER FOR MASTER VIEW
' ============================================================================

masterViewName = InputBox( _
    "Enter the NAME of the Master View to copy style FROM:" & vbCrLf & vbCrLf & vbCrLf & _
    "See the list above for available views." & vbCrLf & vbCrLf & _
    "Examples: VIEW1, ELEVATION, PLAN, SECTION A" & vbCrLf & vbCrLf & _
    vbCrLf & _
    "This view's styling will be applied to target views.", _
    "Select Master View", _
    "VIEW1")

If masterViewName = "" Then
    WScript.Echo "Cancelled by user."
    WScript.Quit
End If

masterViewName = Trim(masterViewName)
WScript.Echo ""
WScript.Echo "=========================================="
WScript.Echo "  SELECTED MASTER VIEW: " & masterViewName
WScript.Echo "=========================================="
WScript.Echo ""

' ============================================================================
' STEP 3: FIND MASTER VIEW ON ACTIVE SHEET ONLY
' ============================================================================

For Each v In activeSheet.DrawingViews
    If UCase(v.Name) = UCase(masterViewName) Then
        Set masterView = v
        Set masterSheet = activeSheet
        Exit For
    End If
Next

If masterView Is Nothing Then
    WScript.Echo "ERROR: Master View '" & masterViewName & "' not found!"
    WScript.Echo ""
    WScript.Echo "Available views on active sheet (" & activeSheet.Name & "):"
    For Each v In activeSheet.DrawingViews
        WScript.Echo "  -> " & v.Name
    Next
    WScript.Quit 1
End If

WScript.Echo "Found Master View on sheet: " & masterSheet.Name
WScript.Echo ""

' ============================================================================
' STEP 4: ANALYZE MASTER VIEW'S LINE TYPES (DETAILED)
' ============================================================================

Set masterVisLayer = Nothing
Set masterHidLayer = Nothing
Set masterCenterLineLayer = Nothing
Set masterCenterMarkLayer = Nothing

WScript.Echo "=========================================="
WScript.Echo "  MASTER VIEW LINE TYPE ANALYSIS"
WScript.Echo "=========================================="
WScript.Echo "View: " & masterView.Name
WScript.Echo ""

' Scan ALL layers in Master View
Dim masterLayerStats
Set masterLayerStats = CreateObject("Scripting.Dictionary")

Set curves = masterView.DrawingCurves
Dim totalCurves
totalCurves = 0
For Each c In curves
    totalCurves = totalCurves + 1
    If c.Segments.Count > 0 Then
        Set l = c.Segments.Item(1).Layer
        Dim layerNameStr
        layerNameStr = l.Name & " [R" & l.Color.Red & " G" & l.Color.Green & " B" & l.Color.Blue & "]"
        If masterLayerStats.Exists(layerNameStr) Then
            masterLayerStats(layerNameStr) = masterLayerStats(layerNameStr) + 1
        Else
            masterLayerStats.Add layerNameStr, 1
        End If

        ' Identify master layers by LAYER NAME (not color)
        Dim uLayerName
        uLayerName = UCase(l.Name)
        
        If InStr(uLayerName, "HIDDEN") > 0 Then
            If masterHidLayer Is Nothing Then
                Set masterHidLayer = l
            End If
        ElseIf InStr(uLayerName, "CENTER MARK") > 0 OR InStr(uLayerName, "CENTERMARK") > 0 Then
            ' Center Mark layer
            If masterCenterMarkLayer Is Nothing Then
                Set masterCenterMarkLayer = l
            End If
        ElseIf InStr(uLayerName, "CENTER") > 0 Then
            ' Center Line layer (but not center mark)
            If masterCenterLineLayer Is Nothing Then
                Set masterCenterLineLayer = l
            End If
        Else
            ' Regular visible layer
            If masterVisLayer Is Nothing Then
                Set masterVisLayer = l
            End If
        End If
    End If
Next

WScript.Echo "Total curves found: " & totalCurves
WScript.Echo ""
WScript.Echo "Line breakdown:"
For Each layerNameStr In masterLayerStats.Keys
    WScript.Echo "  " & layerNameStr & ": " & masterLayerStats(layerNameStr) & " curves"
Next

WScript.Echo ""
WScript.Echo "--- Identified Master Layers ---"
If Not masterVisLayer Is Nothing Then
    WScript.Echo "  Visible Layer: " & masterVisLayer.Name
End If
If Not masterHidLayer Is Nothing Then
    WScript.Echo "  Hidden Layer: " & masterHidLayer.Name
End If
If Not masterCenterLineLayer Is Nothing Then
    WScript.Echo "  Center Line Layer: " & masterCenterLineLayer.Name
End If
If Not masterCenterMarkLayer Is Nothing Then
    WScript.Echo "  Center Mark Layer: " & masterCenterMarkLayer.Name
End If

If masterVisLayer Is Nothing And masterHidLayer Is Nothing Then
    WScript.Echo "WARNING: Could not identify Visible or Hidden Layers from Master View."
End If

WScript.Echo ""

' ============================================================================
' STEP 5: SELECT TARGET VIEWS
' ============================================================================

WScript.Echo "=========================================="
WScript.Echo "  SELECT TARGET VIEWS"
WScript.Echo "=========================================="
WScript.Echo ""
WScript.Echo "Choose which views to apply Master View's style TO:"
WScript.Echo "  1 - ALL views on active sheet (except Master View)"
WScript.Echo "  2 - Specific views (enter names separated by commas)"
WScript.Echo "  3 - Views matching pattern (use wildcards like ELEVATION*, SECTION*)"
WScript.Echo ""

choice = InputBox( _
    "Enter option number (1-3):" & vbCrLf & vbCrLf & _
    "1 - Apply to ALL views" & vbCrLf & _
    "2 - Apply to specific views" & vbCrLf & _
    "3 - Apply to pattern match", _
    "Select Target Views", _
    "1")

If choice = "" Then
    WScript.Echo "Cancelled by user."
    WScript.Quit
End If

' ============================================================================
' STEP 6: BUILD TARGET VIEWS COLLECTION
' ============================================================================

Set targetViewsCollection = CreateObject("Scripting.Dictionary")

Select Case choice
    Case "1"
        WScript.Echo "Selected: ALL views on active sheet (except Master)..."
        WScript.Echo ""
        For Each v In activeSheet.DrawingViews
            If UCase(v.Name) <> UCase(masterViewName) Then
                targetViewsCollection.Add v.Name, v
            End If
        Next

    Case "2"
        WScript.Echo "Selected: Specific views"
        WScript.Echo ""
        specificNames = InputBox( _
            "Enter view names separated by commas:" & vbCrLf & vbCrLf & _
            "Examples:" & vbCrLf & _
            "  VIEW1, ELEVATION, PLAN, SECTION A" & vbCrLf & _
            "  C, D, PLAN, DETAIL 1", _
            "Target Views", _
            "")

        If specificNames = "" Then
            WScript.Echo "Cancelled by user."
            WScript.Quit
        End If

        namesArray = Split(specificNames, ",")
        For i = 0 To UBound(namesArray)
            viewName = Trim(namesArray(i))
            If viewName <> "" Then
                For Each v In activeSheet.DrawingViews
                    If UCase(v.Name) = UCase(viewName) Then
                        targetViewsCollection.Add v.Name, v
                        WScript.Echo "  Added: " & v.Name
                        Exit For
                    End If
                Next
            End If
        Next

    Case "3"
        WScript.Echo "Selected: Pattern matching"
        WScript.Echo ""
        targetPattern = InputBox( _
            "Enter pattern to match views:" & vbCrLf & vbCrLf & _
            "Examples (use uppercase for best results):" & vbCrLf & vbCrLf & _
            "  ELEVATION* - matches ELEVATION, ELEVATION1, ELEVATION A" & vbCrLf & _
            "  SECTION* - matches SECTION A, SECTION B, SECTION LEFT" & vbCrLf & _
            "  *D - matches any view name ending with D (like C, D, RIGHT SIDE D)" & vbCrLf & _
            "  DETAIL* - matches DETAIL 1, DETAIL 2, DETAIL A", _
            "Pattern Match", _
            "*")

        If targetPattern = "" Then
            WScript.Echo "Cancelled by user."
            WScript.Quit
        End If

        targetPattern = UCase(targetPattern)
        For Each v In activeSheet.DrawingViews
            vName = UCase(v.Name)
            If vName <> UCase(masterViewName) Then
                If MatchPattern(vName, targetPattern) Then
                    targetViewsCollection.Add v.Name, v
                End If
            End If
        Next

    Case Else
        WScript.Echo "ERROR: Invalid choice '" & choice & "'"
        WScript.Echo "Please enter 1, 2, or 3."
        WScript.Quit 1
End Select

WScript.Echo ""
WScript.Echo "Found " & targetViewsCollection.Count & " target view(s) to update."
WScript.Echo ""

If targetViewsCollection.Count = 0 Then
    WScript.Echo "No target views found!"
    WScript.Quit 0
End If

' ============================================================================
' STEP 7: APPLY MASTER STYLE TO TARGET VIEWS
' ============================================================================

WScript.Echo "=========================================="
WScript.Echo "  APPLYING STYLE"
WScript.Echo "=========================================="
WScript.Echo ""

keys = targetViewsCollection.Keys

totalCount = 0
visCount = 0
hidCount = 0
centerLineCount = 0
centerMarkCount = 0

For i = 0 To UBound(keys)
    Set targetView = targetViewsCollection.Item(keys(i))

    WScript.Echo "=========================================="
    WScript.Echo "  TARGET VIEW: " & targetView.Name
    WScript.Echo "=========================================="

    Set tCurves = targetView.DrawingCurves

    If Err.Number = 0 And Not tCurves Is Nothing Then
        countViewVis = 0
        countViewHid = 0
        sampleCount = 0

        ' DETAILED ANALYSIS: Show ALL line types in target view
        Dim targetLayerStats
        Set targetLayerStats = CreateObject("Scripting.Dictionary")
        Dim targetTotalCurves
        targetTotalCurves = 0

        For Each c In tCurves
            targetTotalCurves = targetTotalCurves + 1
            For Each s In c.Segments
                Set curLayer = s.Layer
                colorKey = curLayer.Name & " [R" & curLayer.Color.Red & " G" & curLayer.Color.Green & " B" & curLayer.Color.Blue & "]"
                If targetLayerStats.Exists(colorKey) Then
                    targetLayerStats(colorKey) = targetLayerStats(colorKey) + 1
                Else
                    targetLayerStats.Add colorKey, 1
                End If
            Next
        Next

        WScript.Echo "Total curves found: " & targetTotalCurves
        WScript.Echo ""
        WScript.Echo "Line breakdown:"
        For Each colorKey In targetLayerStats.Keys
            WScript.Echo "  " & colorKey & ": " & targetLayerStats(colorKey) & " segments"
        Next
        WScript.Echo ""

        For Each c In tCurves
            For Each s In c.Segments
                Set curLayer = s.Layer

                ' Check layer type by name
                Dim curLayerName
                curLayerName = UCase(curLayer.Name)
                
                ' If Hidden layer --> Move to Master Hidden Layer
                If InStr(curLayerName, "HIDDEN") > 0 Then
                    If Not masterHidLayer Is Nothing Then
                        If sampleCount < 3 Then
                            WScript.Echo "  [SAMPLE] Hidden layer: " & curLayer.Name & " -> " & masterHidLayer.Name
                            sampleCount = sampleCount + 1
                        End If
                        s.Layer = masterHidLayer
                        countViewHid = countViewHid + 1
                    End If

                ' If Center Mark layer --> Move to Master Center Mark Layer
                ElseIf InStr(curLayerName, "CENTER MARK") > 0 OR InStr(curLayerName, "CENTERMARK") > 0 Then
                    If Not masterCenterMarkLayer Is Nothing Then
                        If sampleCount < 3 Then
                            WScript.Echo "  [SAMPLE] Center Mark layer: " & curLayer.Name & " -> " & masterCenterMarkLayer.Name
                            sampleCount = sampleCount + 1
                        End If
                        s.Layer = masterCenterMarkLayer
                        centerMarkCount = centerMarkCount + 1
                    End If

                ' If Center Line layer (but not center mark) --> Move to Master Center Line Layer
                ElseIf InStr(curLayerName, "CENTER") > 0 Then
                    If Not masterCenterLineLayer Is Nothing Then
                        If sampleCount < 3 Then
                            WScript.Echo "  [SAMPLE] Center Line layer: " & curLayer.Name & " -> " & masterCenterLineLayer.Name
                            sampleCount = sampleCount + 1
                        End If
                        s.Layer = masterCenterLineLayer
                        centerLineCount = centerLineCount + 1
                    End If

                ' If Visible layer --> Move to Master Visible Layer
                Else
                    If Not masterVisLayer Is Nothing Then
                        If sampleCount < 3 Then
                            WScript.Echo "  [SAMPLE] Visible layer: " & curLayer.Name & " -> " & masterVisLayer.Name
                            sampleCount = sampleCount + 1
                        End If
                        s.Layer = masterVisLayer
                        countViewVis = countViewVis + 1
                    End If
                End If
            Next
        Next

        WScript.Echo "  -> Processed " & countViewVis & " visible lines to " & masterVisLayer.Name
        WScript.Echo "  -> Processed " & countViewHid & " hidden lines to " & masterHidLayer.Name
        If centerLineCount > 0 Then
            WScript.Echo "  -> Processed " & centerLineCount & " center line segments"
        End If
        If centerMarkCount > 0 Then
            WScript.Echo "  -> Processed " & centerMarkCount & " center mark segments"
        End If
        totalCount = totalCount + 1
        visCount = visCount + countViewVis
        hidCount = hidCount + countViewHid
    Else
        WScript.Echo "  -> ERROR: Could not access drawing curves"
    End If
    Err.Clear

    WScript.Echo ""
Next

WScript.Echo ""

' ============================================================================
' STEP 8: SAVE AND UPDATE DOCUMENT
' ============================================================================

WScript.Echo "=========================================="
WScript.Echo "  SUMMARY"
WScript.Echo "=========================================="
WScript.Echo ""
WScript.Echo "Total views updated: " & totalCount
WScript.Echo "Total visible lines fixed: " & visCount
WScript.Echo "Total hidden lines fixed: " & hidCount
If centerLineCount > 0 Then
    WScript.Echo "Total center lines updated: " & centerLineCount
End If
If centerMarkCount > 0 Then
    WScript.Echo "Total center marks updated: " & centerMarkCount
End If
WScript.Echo ""

WScript.Echo "Updating document..."
On Error Resume Next
doc.Update
If Err.Number = 0 Then
    WScript.Echo "Document updated successfully!"
Else
    WScript.Echo "ERROR updating document: " & Err.Description
End If
Err.Clear

WScript.Echo ""
WScript.Echo "Saving document..."
doc.Save2
If Err.Number = 0 Then
    WScript.Echo "Document saved successfully!"
Else
    WScript.Echo "ERROR saving document: " & Err.Description
End If
Err.Clear

WScript.Echo ""
WScript.Echo "=========================================="
WScript.Echo "  COMPLETE"
WScript.Echo "=========================================="
WScript.Echo ""
WScript.Echo "Style from '" & masterViewName & "' has been applied to " & totalCount & " view(s)."
WScript.Echo ""
WScript.Echo "Please check your drawing in Inventor to verify the changes."

' ============================================================================
' HELPER FUNCTION: Simple pattern matching
' ============================================================================

Function MatchPattern(str, pattern)
    ' Simple wildcard matching (supports * only)
    ' Returns True if str matches pattern

    If pattern = "*" Then
        MatchPattern = True
        Exit Function
    End If

    ' Check for wildcard position
    starPos = InStr(pattern, "*")

    If starPos = 0 Then
        ' No wildcard, exact match required
        MatchPattern = (str = pattern)
        Exit Function
    End If

    If starPos = Len(pattern) Then
        ' Wildcard at end: prefix match
        prefix = Left(pattern, starPos - 1)
        MatchPattern = (Left(str, Len(prefix)) = prefix)
        Exit Function
    End If

    ' Wildcard not at start or end - not supported in this simple version
    MatchPattern = False
End Function
