' Auto_Balloon_From_Template.vbs
' Standalone script version of custom auto-balloon placement
' Workflow:
' 1) Open IDW
' 2) Select exactly one existing balloon (template)
' 3) Run this script
' 4) Script balloons other visible occurrences in that view with reduced crossing

Option Explicit

Dim fso, logPath, logFile
Set fso = CreateObject("Scripting.FileSystemObject")
logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\Auto_Balloon_From_Template.log"
Set logFile = fso.OpenTextFile(logPath, 8, True)

Sub Log(msg)
    On Error Resume Next
    logFile.WriteLine Now & " - " & CStr(msg)
    On Error GoTo 0
End Sub

Dim invApp
Set invApp = GetObject(, "Inventor.Application")
If invApp Is Nothing Then
    Log "ERROR: Could not connect to Inventor."
    MsgBox "Cannot connect to Inventor. Make sure Inventor is running.", vbCritical, "Auto Balloon From Template"
    logFile.Close
    WScript.Quit 1
End If

If invApp.ActiveDocument Is Nothing Then
    Log "ERROR: No active document."
    MsgBox "Open an IDW drawing first.", vbExclamation, "Auto Balloon From Template"
    logFile.Close
    WScript.Quit 1
End If

If invApp.ActiveDocument.DocumentType <> 12292 Then ' Drawing document
    Log "ERROR: Active document is not drawing. Type=" & CStr(invApp.ActiveDocument.DocumentType)
    MsgBox "Active document is not a drawing (IDW).", vbExclamation, "Auto Balloon From Template"
    logFile.Close
    WScript.Quit 1
End If

Dim idwDoc, sheet
Set idwDoc = invApp.ActiveDocument
Set sheet = idwDoc.ActiveSheet
Log "START: Drawing=" & idwDoc.DisplayName & ", Sheet=" & sheet.Name

Dim templateBalloon
Set templateBalloon = GetSelectedTemplateBalloon(idwDoc)
If templateBalloon Is Nothing Then
    Set templateBalloon = GetFirstBalloonOnSheet(sheet)
    If Not templateBalloon Is Nothing Then
        Log "No selected template balloon found. Falling back to first balloon on sheet."
    Else
        Log "ERROR: No single selected template balloon found and no balloon exists on active sheet."
        MsgBox "Select exactly ONE balloon on the active sheet, then run again." & vbCrLf & _
               "(No fallback balloon found on sheet.)", vbInformation, "Auto Balloon From Template"
        logFile.Close
        WScript.Quit 1
    End If
End If

Dim targetView
Set targetView = ResolveViewForBalloon(sheet, templateBalloon)
If targetView Is Nothing Then
    Log "ERROR: Could not resolve target view from selected balloon."
    MsgBox "Could not resolve a target view from the selected balloon.", vbExclamation, "Auto Balloon From Template"
    logFile.Close
    WScript.Quit 1
End If
Log "Target view resolved: " & targetView.Name

Dim mode
mode = MsgBox( _
    "YES = Try adding vertex leaders to selected balloon first, then fallback to new balloons." & vbCrLf & _
    "NO = Create separate balloons only (recommended).", _
    vbYesNoCancel + vbQuestion, "Leader Mode")
If mode = vbCancel Then
    Log "Cancelled by user at mode prompt."
    logFile.Close
    WScript.Quit 0
End If

Dim tryVertex
tryVertex = (mode = vbYes)
If tryVertex Then
    Log "Mode selected: Vertex-first"
Else
    Log "Mode selected: Separate-balloons-only"
End If

Dim occs
Set occs = GetViewOccurrences(targetView)
If occs Is Nothing Then
    Log "Primary occurrence fetch returned Nothing. Trying ReferencedEntities fallback."
    Set occs = GetViewOccurrencesFromReferencedEntities(targetView, invApp)
End If

If occs Is Nothing Then
    Log "ERROR: Could not get any occurrences from target view."
    MsgBox "Could not get occurrences from target view." & vbCrLf & _
           "See log: " & logPath, vbExclamation, "Auto Balloon From Template"
    logFile.Close
    WScript.Quit 1
End If

Log "Occurrences detected: " & CStr(occs.Count)

Dim balloonedKeys
Set balloonedKeys = CollectAlreadyBalloonedOccurrenceKeys(sheet, targetView)

Dim rightArr(), leftArr()
Dim rightCount, leftCount
rightCount = 0
leftCount = 0
ReDim rightArr(0)
ReDim leftArr(0)

Dim targetBounds, viewCenterX
Set targetBounds = GetViewBounds(targetView)
If targetBounds Is Nothing Then
    Log "ERROR: Could not resolve view bounds for target view."
    MsgBox "Could not resolve bounds for target view. See log: " & logPath, vbExclamation, "Auto Balloon From Template"
    logFile.Close
    WScript.Quit 1
End If
viewCenterX = (CDbl(targetBounds("minX")) + CDbl(targetBounds("maxX"))) / 2

Dim i
Dim scannedCount
scannedCount = 0
For i = 1 To occs.Count
    Dim occ, occKey
    Set occ = occs.Item(i)
    scannedCount = scannedCount + 1
    occKey = GetOccurrenceKey(occ)

    If occKey <> "" Then
        If Not balloonedKeys.Exists(occKey) Then
            If IsOccurrenceVisibleInView(occ, targetView) Then
                Dim curve, intent, attachPt
                Set curve = GetFirstDrawingCurveForOccurrence(targetView, occ)
                If Not curve Is Nothing Then
                    On Error Resume Next
                    Set intent = sheet.CreateGeometryIntent(curve)
                    If Err.Number <> 0 Then
                        Set intent = Nothing
                        Err.Clear
                    End If
                    On Error GoTo 0

                    If Not intent Is Nothing Then
                        Set attachPt = GetCurveMidPoint2d(curve, invApp)
                        If Not attachPt Is Nothing Then
                            Dim c
                            Set c = CreateObject("Scripting.Dictionary")
                            c("occ") = occ
                            c("key") = occKey
                            c("intent") = intent
                            c("x") = attachPt.X
                            c("y") = attachPt.Y
                            c("label") = GetOccurrenceLabel(occ)

                            If attachPt.X >= viewCenterX Then
                                AddCandidate rightArr, rightCount, c
                            Else
                                AddCandidate leftArr, leftCount, c
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
Next

If rightCount = 0 And leftCount = 0 Then
    Log "Occurrence-based candidate scan found 0. Trying DrawingCurves-based fallback."
    CollectCandidatesFromViewCurves targetView, sheet, balloonedKeys, viewCenterX, invApp, rightArr, rightCount, leftArr, leftCount
    Log "DrawingCurves fallback candidates: right=" & CStr(rightCount) & ", left=" & CStr(leftCount)
End If

If rightCount = 0 And leftCount = 0 Then
    Log "No candidates after filtering. Scanned=" & CStr(scannedCount) & ", Existing balloons in view=" & CStr(balloonedKeys.Count)
    MsgBox "No additional visible occurrences found to balloon in this view." & vbCrLf & _
           "Scanned occurrences: " & CStr(scannedCount) & vbCrLf & _
           "See log: " & logPath, vbInformation, "Auto Balloon From Template"
    logFile.Close
    WScript.Quit 0
End If

SortCandidatesByYDesc rightArr, rightCount
SortCandidatesByYDesc leftArr, leftCount

Dim createdCount, vertexCount
createdCount = 0
vertexCount = 0

createdCount = createdCount + PlaceSide(sheet, targetView, templateBalloon, rightArr, rightCount, True, tryVertex, vertexCount, invApp)
createdCount = createdCount + PlaceSide(sheet, targetView, templateBalloon, leftArr, leftCount, False, tryVertex, vertexCount, invApp)

On Error Resume Next
idwDoc.Save
On Error GoTo 0

MsgBox "Done." & vbCrLf & _
       "Vertex leaders added: " & CStr(vertexCount) & vbCrLf & _
    "New balloons added: " & CStr(createdCount) & vbCrLf & _
    "Log: " & logPath, _
       vbInformation, "Auto Balloon From Template"

logFile.Close

WScript.Quit 0

' ---------------------------- Helpers ----------------------------

Function GetSelectedTemplateBalloon(drawingDoc)
    Dim selectedBalloon, obj
    Set selectedBalloon = Nothing

    For Each obj In drawingDoc.SelectSet
        If TypeName(obj) = "Balloon" Then
            If Not selectedBalloon Is Nothing Then
                Set GetSelectedTemplateBalloon = Nothing
                Exit Function
            End If
            Set selectedBalloon = obj
        End If
    Next

    Set GetSelectedTemplateBalloon = selectedBalloon
End Function

Function GetFirstBalloonOnSheet(sheet)
    Set GetFirstBalloonOnSheet = Nothing
    On Error Resume Next
    If Not sheet Is Nothing Then
        If sheet.Balloons.Count > 0 Then
            Set GetFirstBalloonOnSheet = sheet.Balloons.Item(1)
        End If
    End If
    On Error GoTo 0
End Function

Function GetViewOccurrencesFromReferencedEntities(view, invApp)
    Set GetViewOccurrencesFromReferencedEntities = Nothing

    On Error Resume Next
    Dim ents
    Set ents = view.ReferencedEntities
    If Err.Number <> 0 Or ents Is Nothing Then
        Err.Clear
        Exit Function
    End If

    Dim col
    Set col = invApp.TransientObjects.CreateObjectCollection
    Dim seen
    Set seen = CreateObject("Scripting.Dictionary")

    Dim i
    For i = 1 To ents.Count
        Dim ent, p, t, key
        Set ent = ents.Item(i)
        Set p = Nothing

        On Error Resume Next
        Set p = ent.Parent
        If Err.Number <> 0 Then
            Err.Clear
            Set p = Nothing
        End If
        On Error GoTo 0

        If Not p Is Nothing Then
            t = TypeName(p)
            If t = "ComponentOccurrence" Or t = "AssemblyOccurrence" Then
                key = GetOccurrenceKey(p)
                If key <> "" Then
                    If Not seen.Exists(key) Then
                        seen.Add key, True
                        col.Add p
                    End If
                End If
            End If
        End If
    Next

    If col.Count > 0 Then
        Set GetViewOccurrencesFromReferencedEntities = col
        Log "Fallback occurrence fetch succeeded. Count=" & CStr(col.Count)
    Else
        Log "Fallback occurrence fetch found 0 occurrences."
    End If
End Function

Function ResolveViewForBalloon(sheet, balloon)
    Dim v
    Set v = TryGetBalloonAttachedView(balloon)
    If Not v Is Nothing Then
        Set ResolveViewForBalloon = v
        Exit Function
    End If

    Dim bPos
    On Error Resume Next
    Set bPos = balloon.Position
    If Err.Number <> 0 Then
        Set bPos = Nothing
        Err.Clear
    End If
    On Error GoTo 0

    If bPos Is Nothing Then
        Set ResolveViewForBalloon = Nothing
        Exit Function
    End If

    Dim bestView, bestD2
    Set bestView = Nothing
    bestD2 = 1E+99

    Dim i, view
    For i = 1 To sheet.DrawingViews.Count
        Set view = sheet.DrawingViews.Item(i)

        Dim b
        Set b = GetViewBounds(view)

        If Not b Is Nothing Then
            If bPos.X >= CDbl(b("minX")) And bPos.X <= CDbl(b("maxX")) And _
               bPos.Y >= CDbl(b("minY")) And bPos.Y <= CDbl(b("maxY")) Then
                Set ResolveViewForBalloon = view
                Exit Function
            End If

            Dim cx, cy, dx, dy, d2
            cx = (CDbl(b("minX")) + CDbl(b("maxX"))) / 2
            cy = (CDbl(b("minY")) + CDbl(b("maxY"))) / 2
            dx = bPos.X - cx
            dy = bPos.Y - cy
            d2 = (dx * dx) + (dy * dy)
            If d2 < bestD2 Then
                bestD2 = d2
                Set bestView = view
            End If
        End If
    Next

    Set ResolveViewForBalloon = bestView
End Function

Function GetViewBounds(view)
    Set GetViewBounds = Nothing

    Dim d
    Set d = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    Dim rb
    Set rb = view.RangeBox
    If Err.Number = 0 And Not rb Is Nothing Then
        d("minX") = rb.MinPoint.X
        d("maxX") = rb.MaxPoint.X
        d("minY") = rb.MinPoint.Y
        d("maxY") = rb.MaxPoint.Y
        Set GetViewBounds = d
        Exit Function
    End If

    Err.Clear
    Dim pos, w, h
    Set pos = Nothing
    Set pos = view.Position
    w = CDbl(view.Width)
    h = CDbl(view.Height)
    If Err.Number = 0 And Not pos Is Nothing Then
        d("minX") = pos.X - (w / 2)
        d("maxX") = pos.X + (w / 2)
        d("minY") = pos.Y - (h / 2)
        d("maxY") = pos.Y + (h / 2)
        Set GetViewBounds = d
        Exit Function
    End If

    Err.Clear
    On Error GoTo 0
End Function

Function TryGetBalloonAttachedView(balloon)
    On Error Resume Next
    Dim leaderObj, rootNode, childNodes
    Set leaderObj = balloon.Leader
    If Err.Number <> 0 Or leaderObj Is Nothing Then
        Set TryGetBalloonAttachedView = Nothing
        Exit Function
    End If

    Set rootNode = leaderObj.RootNode
    If Err.Number <> 0 Or rootNode Is Nothing Then
        Set TryGetBalloonAttachedView = Nothing
        Exit Function
    End If

    Set childNodes = rootNode.ChildNodes
    If Err.Number <> 0 Or childNodes Is Nothing Then
        Set TryGetBalloonAttachedView = Nothing
        Exit Function
    End If

    Dim i, node, attached, v
    For i = 1 To childNodes.Count
        Set node = childNodes.Item(i)
        Set attached = Nothing
        Set v = Nothing

        On Error Resume Next
        Set attached = node.AttachedEntity
        On Error GoTo 0

        Set v = ResolveViewFromAttachedEntity(attached)
        If Not v Is Nothing Then
            Set TryGetBalloonAttachedView = v
            Exit Function
        End If
    Next

    Set TryGetBalloonAttachedView = Nothing
End Function

Function ResolveViewFromAttachedEntity(attached)
    Set ResolveViewFromAttachedEntity = Nothing
    If attached Is Nothing Then Exit Function

    On Error Resume Next
    Dim t
    t = TypeName(attached)

    If t = "GeometryIntent" Then
        Dim g
        Set g = Nothing
        Set g = attached.Geometry
        Set ResolveViewFromAttachedEntity = ResolveViewFromAttachedEntity(g)
        Exit Function
    End If

    If t = "DrawingCurve" Then
        Set ResolveViewFromAttachedEntity = attached.Parent
        Exit Function
    End If

    If t = "DrawingCurveSegment" Then
        Set ResolveViewFromAttachedEntity = attached.Parent.Parent
        Exit Function
    End If

    On Error GoTo 0
End Function

Function GetViewOccurrences(view)
    Set GetViewOccurrences = Nothing

    On Error Resume Next
    Dim refDoc
    Set refDoc = Nothing

    If Not view.ReferencedDocumentDescriptor Is Nothing Then
        Set refDoc = view.ReferencedDocumentDescriptor.ReferencedDocument
        If Err.Number = 0 Then
            If Not refDoc Is Nothing Then
                If refDoc.DocumentType = 12291 Then
                    Set GetViewOccurrences = refDoc.ComponentDefinition.Occurrences
                    Exit Function
                End If
            End If
        Else
            Err.Clear
        End If
    End If

    Set refDoc = Nothing
    Set refDoc = view.ReferencedDocument
    If Err.Number = 0 Then
        If Not refDoc Is Nothing Then
            If refDoc.DocumentType = 12291 Then
                Set GetViewOccurrences = refDoc.ComponentDefinition.Occurrences
                Exit Function
            End If
        End If
    End If

    On Error GoTo 0
End Function

Function CollectAlreadyBalloonedOccurrenceKeys(sheet, targetView)
    Dim d
    Set d = CreateObject("Scripting.Dictionary")

    Dim i
    For i = 1 To sheet.Balloons.Count
        On Error Resume Next
        Dim b, leaderObj, rootNode, childNodes
        Set b = sheet.Balloons.Item(i)
        Set leaderObj = b.Leader
        Set rootNode = leaderObj.RootNode
        Set childNodes = rootNode.ChildNodes
        If Err.Number = 0 Then
            On Error GoTo 0

            Dim j
            For j = 1 To childNodes.Count
                Dim node, attached, v, key
                Set node = childNodes.Item(j)
                Set attached = Nothing
                On Error Resume Next
                Set attached = node.AttachedEntity
                On Error GoTo 0

                Set v = ResolveViewFromAttachedEntity(attached)
                If Not v Is Nothing Then
                    If v Is targetView Then
                        key = GetOccurrenceKeyFromAttached(attached)
                        If key <> "" Then
                            If Not d.Exists(key) Then d.Add key, True
                        End If
                    End If
                End If
            Next
        Else
            Err.Clear
            On Error GoTo 0
        End If
    Next

    Set CollectAlreadyBalloonedOccurrenceKeys = d
End Function

Function GetOccurrenceKeyFromAttached(attached)
    GetOccurrenceKeyFromAttached = ""
    If attached Is Nothing Then Exit Function

    On Error Resume Next
    Dim t
    t = TypeName(attached)

    If t = "GeometryIntent" Then
        GetOccurrenceKeyFromAttached = GetOccurrenceKeyFromAttached(attached.Geometry)
        Exit Function
    End If

    If t = "DrawingCurve" Then
        GetOccurrenceKeyFromAttached = GetOccurrenceKey(attached.ModelGeometry.Parent)
        Exit Function
    End If

    If t = "DrawingCurveSegment" Then
        GetOccurrenceKeyFromAttached = GetOccurrenceKey(attached.Parent.ModelGeometry.Parent)
        Exit Function
    End If

    On Error GoTo 0
End Function

Function GetOccurrenceKey(occ)
    GetOccurrenceKey = ""
    If occ Is Nothing Then Exit Function

    On Error Resume Next
    GetOccurrenceKey = CStr(occ.InternalName)
    If Err.Number <> 0 Then
        Err.Clear
        GetOccurrenceKey = ""
    End If

    If GetOccurrenceKey = "" Then
        On Error Resume Next
        GetOccurrenceKey = CStr(occ.Name)
        If Err.Number <> 0 Then
            Err.Clear
            GetOccurrenceKey = ""
        End If
        On Error GoTo 0
    End If
End Function

Function IsOccurrenceVisibleInView(occ, view)
    On Error Resume Next
    IsOccurrenceVisibleInView = CBool(occ.VisibleInDrawingView(view))
    If Err.Number <> 0 Then
        Err.Clear
        IsOccurrenceVisibleInView = True
    End If
    On Error GoTo 0
End Function

Function GetFirstDrawingCurveForOccurrence(view, occ)
    On Error Resume Next
    Dim curves
    Set curves = view.DrawingCurves(occ)
    If Err.Number <> 0 Then
        Err.Clear
        Set curves = Nothing
    End If

    If Not curves Is Nothing Then
        If curves.Count > 0 Then
            Set GetFirstDrawingCurveForOccurrence = curves.Item(1)
            Exit Function
        End If
    End If

    Set GetFirstDrawingCurveForOccurrence = Nothing
    On Error GoTo 0
End Function

Function GetCurveMidPoint2d(curve, invApp)
    On Error Resume Next
    Dim rb
    Set rb = curve.RangeBox
    If Err.Number <> 0 Or rb Is Nothing Then
        Err.Clear
        Set GetCurveMidPoint2d = Nothing
        Exit Function
    End If

    Dim x, y
    x = (rb.MinPoint.X + rb.MaxPoint.X) / 2
    y = (rb.MinPoint.Y + rb.MaxPoint.Y) / 2

    Set GetCurveMidPoint2d = invApp.TransientGeometry.CreatePoint2d(x, y)
    On Error GoTo 0
End Function

Function GetOccurrenceLabel(occ)
    GetOccurrenceLabel = ""

    On Error Resume Next
    Dim itemNo
    itemNo = ""
    itemNo = CStr(occ.ItemNumber)
    If Err.Number = 0 Then
        If itemNo <> "" Then
            GetOccurrenceLabel = itemNo
            Exit Function
        End If
    Else
        Err.Clear
    End If

    Dim refDoc, pn
    Set refDoc = Nothing
    Set refDoc = occ.ReferencedDocument
    If Err.Number = 0 Then
        If Not refDoc Is Nothing Then
            pn = CStr(refDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value)
            If Err.Number = 0 Then
                If pn <> "" Then
                    GetOccurrenceLabel = pn
                    Exit Function
                End If
            Else
                Err.Clear
            End If
        End If
    Else
        Err.Clear
    End If

    GetOccurrenceLabel = CStr(occ.Name)
    If Err.Number <> 0 Then
        Err.Clear
        GetOccurrenceLabel = ""
    End If
    On Error GoTo 0
End Function

Sub AddCandidate(ByRef arr, ByRef count, ByVal candidate)
    If count = 0 Then
        arr(0) = candidate
        count = 1
    Else
        ReDim Preserve arr(count)
        arr(count) = candidate
        count = count + 1
    End If
End Sub

Sub CollectCandidatesFromViewCurves(view, sheet, balloonedKeys, viewCenterX, invApp, ByRef rightArr, ByRef rightCount, ByRef leftArr, ByRef leftCount)
    On Error Resume Next
    Dim curves
    Set curves = view.DrawingCurves
    If Err.Number <> 0 Or curves Is Nothing Then
        Err.Clear
        Log "DrawingCurves fallback unavailable on this view."
        Exit Sub
    End If
    On Error GoTo 0

    Dim seen
    Set seen = CreateObject("Scripting.Dictionary")

    Dim i
    For i = 1 To curves.Count
        Dim curve, key
        Set curve = curves.Item(i)
        key = GetCurveOccurrenceKey(curve)

        If key = "" Then key = "curve_" & CStr(i)

        If Not balloonedKeys.Exists(key) Then
            If Not seen.Exists(key) Then
                seen.Add key, True

                Dim intent, attachPt, c
                Set intent = Nothing
                Set attachPt = Nothing

                On Error Resume Next
                Set intent = sheet.CreateGeometryIntent(curve)
                If Err.Number <> 0 Then
                    Err.Clear
                    Set intent = Nothing
                End If
                On Error GoTo 0

                If Not intent Is Nothing Then
                    Set attachPt = GetCurveMidPoint2d(curve, invApp)
                    If Not attachPt Is Nothing Then
                        Set c = CreateObject("Scripting.Dictionary")
                        c("occ") = Nothing
                        c("key") = key
                        c("intent") = intent
                        c("x") = attachPt.X
                        c("y") = attachPt.Y
                        c("label") = GetCurveLabel(curve)

                        If attachPt.X >= viewCenterX Then
                            AddCandidate rightArr, rightCount, c
                        Else
                            AddCandidate leftArr, leftCount, c
                        End If
                    End If
                End If
            End If
        End If
    Next
End Sub

Function GetCurveOccurrenceKey(curve)
    GetCurveOccurrenceKey = ""
    If curve Is Nothing Then Exit Function

    On Error Resume Next
    Dim mg, p
    Set mg = curve.ModelGeometry
    If Err.Number <> 0 Or mg Is Nothing Then
        Err.Clear
        Exit Function
    End If

    Set p = mg.Parent
    If Err.Number = 0 And Not p Is Nothing Then
        GetCurveOccurrenceKey = GetOccurrenceKey(p)
    Else
        Err.Clear
    End If

    If GetCurveOccurrenceKey = "" Then
        On Error Resume Next
        GetCurveOccurrenceKey = CStr(mg.InternalName)
        If Err.Number <> 0 Then
            Err.Clear
            GetCurveOccurrenceKey = ""
        End If
    End If
    On Error GoTo 0
End Function

Function GetCurveLabel(curve)
    GetCurveLabel = ""
    If curve Is Nothing Then Exit Function

    On Error Resume Next
    Dim mg, p
    Set mg = curve.ModelGeometry
    If Err.Number <> 0 Or mg Is Nothing Then
        Err.Clear
        Exit Function
    End If

    Set p = mg.Parent
    If Err.Number = 0 And Not p Is Nothing Then
        GetCurveLabel = GetOccurrenceLabel(p)
    Else
        Err.Clear
    End If
    On Error GoTo 0
End Function

Sub SortCandidatesByYDesc(ByRef arr, ByVal count)
    If count <= 1 Then Exit Sub

    Dim i, j
    For i = 0 To count - 2
        For j = i + 1 To count - 1
            If CDbl(arr(j)("y")) > CDbl(arr(i)("y")) Then
                Dim temp
                Set temp = arr(i)
                Set arr(i) = arr(j)
                Set arr(j) = temp
            End If
        Next
    Next
End Sub

Function PlaceSide(sheet, view, templateBalloon, arr, count, isRight, tryVertex, ByRef vertexCount, invApp)
    PlaceSide = 0
    If count = 0 Then Exit Function

    Dim b, topY, bottomY, span, gap, outerOffset, innerOffset
    Set b = GetViewBounds(view)
    If b Is Nothing Then Exit Function

    topY = CDbl(b("maxY"))
    bottomY = CDbl(b("minY"))
    span = topY - bottomY
    If span < 0.001 Then span = 0.001

    gap = span / (count + 2)
    If gap < 0.6 Then gap = 0.6

    outerOffset = span * 0.03
    If outerOffset < 1.2 Then outerOffset = 1.2
    innerOffset = outerOffset * 0.55
    If innerOffset < 0.6 Then innerOffset = 0.6

    Dim endpointX, elbowX
    If isRight Then
        endpointX = CDbl(b("maxX")) + outerOffset
        elbowX = CDbl(b("maxX")) + innerOffset
    Else
        endpointX = CDbl(b("minX")) - outerOffset
        elbowX = CDbl(b("minX")) - innerOffset
    End If

    Dim nextY, i
    nextY = topY - gap

    For i = 0 To count - 1
        Dim c
        Set c = arr(i)

        Dim slotY
        slotY = nextY
        If slotY > CDbl(c("y")) Then slotY = CDbl(c("y"))
        If slotY < bottomY + gap Then slotY = bottomY + gap

        Dim endPt, elbowPt
        Set endPt = invApp.TransientGeometry.CreatePoint2d(endpointX, slotY)
        Set elbowPt = invApp.TransientGeometry.CreatePoint2d(elbowX, CDbl(c("y")))

        Dim doneByVertex
        doneByVertex = False
        If tryVertex Then
            doneByVertex = TryAddVertexLeader(templateBalloon, c("intent"), endPt, elbowPt)
            If doneByVertex Then vertexCount = vertexCount + 1
        End If

        If Not doneByVertex Then
            If AddBalloonUsingTemplateStyle(sheet, templateBalloon, c, endPt, elbowPt) Then
                PlaceSide = PlaceSide + 1
            End If
        End If

        nextY = slotY - gap
    Next
End Function

Function TryAddVertexLeader(templateBalloon, attachIntent, endPt, elbowPt)
    TryAddVertexLeader = False

    On Error Resume Next
    Dim leaderObj, rootNode, childNodes
    Set leaderObj = templateBalloon.Leader
    If Err.Number <> 0 Or leaderObj Is Nothing Then
        Err.Clear
        Exit Function
    End If

    Set rootNode = leaderObj.RootNode
    If Err.Number <> 0 Or rootNode Is Nothing Then
        Err.Clear
        Exit Function
    End If

    Set childNodes = rootNode.ChildNodes
    If Err.Number <> 0 Or childNodes Is Nothing Then
        Err.Clear
        Exit Function
    End If

    Err.Clear
    Call childNodes.Add(endPt, elbowPt, attachIntent)
    If Err.Number = 0 Then
        TryAddVertexLeader = True
        Exit Function
    End If

    Err.Clear
    Call childNodes.Add(endPt, attachIntent)
    If Err.Number = 0 Then
        TryAddVertexLeader = True
        Exit Function
    End If

    Err.Clear
End Function

Function AddBalloonUsingTemplateStyle(sheet, templateBalloon, c, endPt, elbowPt)
    AddBalloonUsingTemplateStyle = False

    On Error Resume Next
    Dim leaderPoints
    Set leaderPoints = invApp.TransientObjects.CreateObjectCollection
    leaderPoints.Add endPt
    leaderPoints.Add elbowPt
    leaderPoints.Add c("intent")

    Dim b
    Set b = Nothing

    Err.Clear
    Set b = sheet.Balloons.Add(leaderPoints, , , , templateBalloon.Style)
    If Err.Number <> 0 Or b Is Nothing Then
        Err.Clear
        Set b = sheet.Balloons.Add(leaderPoints)
        If Not b Is Nothing Then
            On Error Resume Next
            b.Style = templateBalloon.Style
            Err.Clear
        End If
    End If

    If b Is Nothing Then
        Exit Function
    End If

    If CStr(c("label")) <> "" Then
        On Error Resume Next
        b.Text = CStr(c("label"))
        Err.Clear
    End If

    AddBalloonUsingTemplateStyle = True
End Function
