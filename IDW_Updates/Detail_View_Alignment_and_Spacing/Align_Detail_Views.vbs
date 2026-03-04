' Align_Detail_Views.vbs
' Aligns and spaces detail views on the active IDW sheet.
' Safe to run: checks for Inventor application and active IDW.
Option Explicit
On Error Resume Next
Dim invApp, doc, sheet, drawingViews, v

' Try to get Inventor application
Set invApp = Nothing
If Not WScript.Arguments.Named.Exists("invApp") Then
  On Error Resume Next
  Set invApp = GetObject(, "Inventor.Application")
  If Err.Number <> 0 Then
    Err.Clear
    WScript.Echo "Could not get Inventor.Application via GetObject. Exiting (no Inventor)."
    WScript.Quit 1
  End If
Else
  Set invApp = WScript.Arguments.Named.Item("invApp")
End If
On Error GoTo 0

Set doc = invApp.ActiveDocument
If doc Is Nothing Then
  WScript.Echo "No active document. Open an IDW and try again." : WScript.Quit 1
End If

If LCase(Right(doc.FullFileName,4)) <> ".idw" Then
  WScript.Echo "Active document is not an IDW: " & doc.FullFileName : WScript.Quit 1
End If

' Use first sheet
If doc.Sheets.Count = 0 Then
  WScript.Echo "No sheets in IDW." : WScript.Quit 1
End If

Set sheet = doc.Sheets.Item(1)
Set drawingViews = sheet.DrawingViews

If drawingViews.Count < 2 Then
  WScript.Echo "Not enough views to align/space (need >=2)." : WScript.Quit 1
End If

' PARAMETERS and CLI options
Dim spacingMM, layout, wThresh, colsArg, rowsArg, marginMM, previewMode, debugMode
spacingMM = 10        ' spacing in mm between views (default)
layout = "horizontal"  ' "horizontal", "vertical", or "grid"
wThresh = 100         ' size threshold to pick small views (sheet units)
colsArg = 0            ' if >0, used for grid columns
rowsArg = 0
marginMM = 10         ' margin in mm from sheet edge
previewMode = False
debugMode = False

' Parse named command-line args (examples: /spacing:8 /layout:grid /cols:3 /margin:15 /preview:true)
If WScript.Arguments.Named.Exists("spacing") Then spacingMM = CDbl(WScript.Arguments.Named.Item("spacing"))
If WScript.Arguments.Named.Exists("layout") Then layout = LCase(WScript.Arguments.Named.Item("layout"))
If WScript.Arguments.Named.Exists("wthresh") Then wThresh = CDbl(WScript.Arguments.Named.Item("wthresh"))
If WScript.Arguments.Named.Exists("cols") Then colsArg = CInt(WScript.Arguments.Named.Item("cols"))
If WScript.Arguments.Named.Exists("rows") Then rowsArg = CInt(WScript.Arguments.Named.Item("rows"))
If WScript.Arguments.Named.Exists("margin") Then marginMM = CDbl(WScript.Arguments.Named.Item("margin"))
If WScript.Arguments.Named.Exists("preview") Then previewMode = (LCase(WScript.Arguments.Named.Item("preview")) = "true" Or WScript.Arguments.Named.Item("preview") = "1")
If WScript.Arguments.Named.Exists("debug") Then debugMode = (LCase(WScript.Arguments.Named.Item("debug")) = "true" Or WScript.Arguments.Named.Item("debug") = "1")
' Force/unlock option: try to unanchor/unlock views before moving
Dim forceMode
forceMode = False
If WScript.Arguments.Named.Exists("force") Then forceMode = (LCase(WScript.Arguments.Named.Item("force")) = "true" Or WScript.Arguments.Named.Item("force") = "1")
' Recreate option: try to create replacement views if moves fail
Dim recreateMode
recreateMode = False
If WScript.Arguments.Named.Exists("recreate") Then recreateMode = (LCase(WScript.Arguments.Named.Item("recreate")) = "true" Or WScript.Arguments.Named.Item("recreate") = "1")
' Save and undo options
Dim saveMode, undoMode, undoFilePath
saveMode = False
undoMode = False
undoFilePath = ""
If WScript.Arguments.Named.Exists("save") Then saveMode = (LCase(WScript.Arguments.Named.Item("save")) = "true" Or WScript.Arguments.Named.Item("save") = "1")
If WScript.Arguments.Named.Exists("undo") Then undoMode = (LCase(WScript.Arguments.Named.Item("undo")) = "true" Or WScript.Arguments.Named.Item("undo") = "1")
If WScript.Arguments.Named.Exists("undoFile") Then undoFilePath = WScript.Arguments.Named.Item("undoFile")
' Default undo file name in same folder
If undoFilePath = "" Then
  undoFilePath = Left(WScript.ScriptFullName, Len(WScript.ScriptFullName) - Len(WScript.ScriptName)) & "Align_Detail_Views.undo.csv"
End If

' If undo mode requested, attempt to read the undo file and restore positions immediately
If undoMode Then
  Set fso = CreateObject("Scripting.FileSystemObject")
  If Not fso.FileExists(undoFilePath) Then
    WScript.Echo "Undo file not found: " & undoFilePath & ". Nothing to undo." : WScript.Quit 1
  End If
  Set undoFS = fso.OpenTextFile(undoFilePath, 1)
  WScript.Echo "Undoing positions from " & undoFilePath & "..."
  Do While Not undoFS.AtEndOfStream
    Dim line, parts, nameKey, ox, oy
    line = undoFS.ReadLine
    parts = Split(line, ",")
    If UBound(parts) >= 2 Then
      ' Support two formats:
      ' 1) name,ox,oy
      ' 2) bakName,newName,ox,oy  (recreate mode)
      If UBound(parts) = 2 Then
        nameKey = Trim(parts(0))
        ox = CDbl(Trim(parts(1)))
        oy = CDbl(Trim(parts(2)))
        ' find view with this name
        Dim foundView, j
        Set foundView = Nothing
        For j = 1 To drawingViews.Count
          If LCase(Trim(drawingViews.Item(j).Name)) = LCase(nameKey) Then
            Set foundView = drawingViews.Item(j)
            Exit For
          End If
        Next
        If Not foundView Is Nothing Then
          Dim curdx, curdy
          curdx = ox - foundView.Center.X
          curdy = oy - foundView.Center.Y
          If Abs(curdx) > 1e-9 Then foundView.Position.TranslateBy invApp.TransientGeometry.CreateVector2d(curdx, 0)
          If Abs(curdy) > 1e-9 Then foundView.Position.TranslateBy invApp.TransientGeometry.CreateVector2d(0, curdy)
          WScript.Echo "Restored " & foundView.Name & " to (" & FormatNumber(ox,3) & "," & FormatNumber(oy,3) & ")"
        Else
          WScript.Echo "Warning: could not find view named '" & nameKey & "' for undo."
        End If
      ElseIf UBound(parts) >= 3 Then
        ' Recreate undo entry: bakName,newName,ox,oy
        Dim bakName, newName
        bakName = Trim(parts(0))
        newName = Trim(parts(1))
        ox = CDbl(Trim(parts(2)))
        oy = CDbl(Trim(parts(3)))
        ' If new (recreated) view exists, delete it
        Dim recreatedView
        Set recreatedView = Nothing
        For j = 1 To drawingViews.Count
          If LCase(Trim(drawingViews.Item(j).Name)) = LCase(newName) Then
            Set recreatedView = drawingViews.Item(j)
            Exit For
          End If
        Next
        If Not recreatedView Is Nothing Then
          recreatedView.Delete
          WScript.Echo "Deleted recreated view: " & newName
        End If
        ' Find bak view and restore name + center
        Dim bakView
        Set bakView = Nothing
        For j = 1 To drawingViews.Count
          If LCase(Trim(drawingViews.Item(j).Name)) = LCase(bakName) Then
            Set bakView = drawingViews.Item(j)
            Exit For
          End If
        Next
        If Not bakView Is Nothing Then
          bakView.Name = newName
          Dim curdx2, curdy2
          curdx2 = ox - bakView.Center.X
          curdy2 = oy - bakView.Center.Y
          If Abs(curdx2) > 1e-9 Then bakView.Position.TranslateBy invApp.TransientGeometry.CreateVector2d(curdx2, 0)
          If Abs(curdy2) > 1e-9 Then bakView.Position.TranslateBy invApp.TransientGeometry.CreateVector2d(0, curdy2)
          WScript.Echo "Restored recreated view back to original: " & newName & " at (" & FormatNumber(ox,3) & "," & FormatNumber(oy,3) & ")"
        Else
          WScript.Echo "Warning: could not find backup view '" & bakName & "' to restore."
        End If
      End If
    End If
  Loop
  undoFS.Close
  doc.Update
  WScript.Echo "Undo complete. Saved positions restored where possible." : WScript.Quit 0
End If

' Convert spacing and margin to document units (mm -> document units)
Dim uom
Set uom = doc.UnitsOfMeasure
Dim spacing, margin
spacing = spacingMM
margin = marginMM
On Error Resume Next
spacing = uom.ConvertUnits(spacingMM, "mm", uom.LengthUnit)
margin = uom.ConvertUnits(marginMM, "mm", uom.LengthUnit)
On Error GoTo 0

WScript.Echo "Found " & drawingViews.Count & " views. Layout='" & layout & "' spacing=" & spacingMM & "mm margin=" & marginMM & "mm preview=" & previewMode & vbCrLf & "Active document: " & doc.FullFileName

' Build targets: preferentially small/detail views (size heuristic) unless /all specified
Dim targets()
Dim tcount
tcount = 0
Dim i
For i = 1 To drawingViews.Count
  Set v = drawingViews.Item(i)
  If v.Width < wThresh And v.Height < wThresh Then
    ReDim Preserve targets(tcount)
    Set targets(tcount) = v
    tcount = tcount + 1
  End If
Next

' If none found using size heuristic, use all views
If tcount = 0 Then
  For i = 1 To drawingViews.Count
    Set v = drawingViews.Item(i)
    ReDim Preserve targets(tcount)
    Set targets(tcount) = v
    tcount = tcount + 1
  Next
End If

If tcount < 2 Then
  WScript.Echo "Not enough target views (found " & tcount & "). Nothing to do." : WScript.Quit 1
End If

' Local temporaries (declare once to avoid redefinition errors)
Dim a, b, tmpObj, tmpObj2, desiredY, dx, dy, dx2, dy2, ddx, ddy, sumX, avgX, sumY, avgY, totalHeight, topEdge, totalW, leftmost, firstCenter, curCenterX, curY, fso, outFS, undoFS, centerNote, minCX, maxCX, minCY, maxCY, adjCX, adjCY
' Logging for moved views
Dim updates(), ucount
ucount = 0
Dim viewName, oldX, oldY, newX, newY

' Helper: get sheet extents and transient geometry
Dim sheetW, sheetH, tg
sheetW = CDbl(sheet.Width)
sheetH = CDbl(sheet.Height)
Set tg = invApp.TransientGeometry

' Helper: attempt to make a view moveable by removing anchors/locks when possible
Sub TryMakeViewMoveable(v, dbg)
  On Error Resume Next
  If dbg Then WScript.Echo "DEBUG: TryMakeViewMoveable: " & CStr(v.Name)
  ' Attempt to clear simple flags if available
  v.Anchored = False
  v.Locked = False
  ' Try to remove a view-space anchor if method exists
  On Error Resume Next
  Call v.RemoveViewSpaceAnchor
  Err.Clear
  If Not v.ComponentGraphics Is Nothing Then
    Call v.ComponentGraphics.RemoveViewSpaceAnchor
    Err.Clear
  End If
  On Error GoTo 0
End Sub

' Helper: recreate a view by renaming original to a bak name and creating a new base view at target center
Function RecreateView(oldV, targetX, targetY, dbg)
  On Error Resume Next
  Dim origName, bakName, timestamp, newV, refDoc, scale, orient
  origName = CStr(oldV.Name)
  timestamp = Replace(CStr(Timer), ".", "")
  bakName = origName & ".bak_recreate_" & timestamp
  If dbg Then WScript.Echo "DEBUG: Recreate: renaming " & origName & " -> " & bakName
  oldV.Name = bakName
  Err.Clear
  ' Try to get referenced document
  On Error Resume Next
  Set refDoc = Nothing
  Set refDoc = oldV.ReferencedDocumentDescriptor.ReferencedDocument
  If dbg Then
    If refDoc Is Nothing Then
      WScript.Echo "DEBUG: RecreateView: could not obtain ReferencedDocument for " & origName
    Else
      WScript.Echo "DEBUG: RecreateView: found referenced document: " & refDoc.FullFileName
    End If
  End If
  scale = 1
  On Error Resume Next
  scale = oldV.Scale
  If dbg Then WScript.Echo "DEBUG: RecreateView: scale=" & CStr(scale)
  Err.Clear
  ' Try to get orientation if present
  On Error Resume Next
  orient = oldV.ViewOrientationFromBase
  If dbg Then WScript.Echo "DEBUG: RecreateView: orient=" & CStr(orient)
  Err.Clear
  Set newV = Nothing
  If Not refDoc Is Nothing Then
    On Error Resume Next
    ' Try basic AddBaseView with minimal args first
    Set newV = Nothing
    On Error Resume Next
    Set newV = sheet.DrawingViews.AddBaseView(refDoc, tg.CreatePoint2d(targetX, targetY), scale)
    If Err.Number <> 0 Then
      If dbg Then WScript.Echo "DEBUG: RecreateView: basic AddBaseView failed: " & CStr(Err.Number) & " - " & Err.Description
      Err.Clear
      Set newV = Nothing
    Else
      ' Attempt to set name and orientation when possible
      On Error Resume Next
      newV.Name = origName
      If dbg Then WScript.Echo "DEBUG: RecreateView: Created base view named " & newV.Name
      Err.Clear
    End If
    On Error GoTo 0
    On Error GoTo 0
  End If

  If newV Is Nothing Then
    ' Failed to create replacement - restore old name
    On Error Resume Next
    oldV.Name = origName
    Err.Clear
    RecreateView = False
    Exit Function
  End If

  ' Success: write undo CSV entry indicating backup + new mapping
  On Error Resume Next
  Dim outFS
  If Not fso.FileExists(undoFilePath) Then
    Set outFS = fso.CreateTextFile(undoFilePath, True)
  Else
    Set outFS = fso.OpenTextFile(undoFilePath, 8)
  End If
  ' Format: bakName,newName,oldX,oldY
  outFS.WriteLine bakName & "," & origName & "," & CStr(oldV.Center.X) & "," & CStr(oldV.Center.Y)
  outFS.Close
  RecreateView = True
End Function

' Prepare undo file (start fresh when applying changes)
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(undoFilePath) Then
  On Error Resume Next
  fso.DeleteFile(undoFilePath)
  On Error GoTo 0
End If

' Helper: simple in-place sort of targets by Center.X (ascending) or Center.Y (descending)
If layout = "horizontal" Then
  For a = 0 To tcount - 2
    For b = a + 1 To tcount - 1
      If targets(b).Center.X < targets(a).Center.X Then
        Set tmpObj = targets(a)
        Set targets(a) = targets(b)
        Set targets(b) = tmpObj
      End If
    Next
  Next
ElseIf layout = "vertical" Then
  For a = 0 To tcount - 2
    For b = a + 1 To tcount - 1
      If targets(b).Center.Y > targets(a).Center.Y Then ' top-down
        Set tmpObj2 = targets(a)
        Set targets(a) = targets(b)
        Set targets(b) = tmpObj2
      End If
    Next
  Next
End If

' Compute some metrics
Dim iMaxW, iMaxH, totalWidth
iMaxW = 0: iMaxH = 0: totalWidth = 0
For i = 0 To tcount - 1
  If targets(i).Width > iMaxW Then iMaxW = targets(i).Width
  If targets(i).Height > iMaxH Then iMaxH = targets(i).Height
  totalWidth = totalWidth + targets(i).Width
Next

' Layout implementations
If layout = "grid" Then
  ' Determine columns if not specified
  Dim cols, rows, estCols
  cols = colsArg
  If cols < 1 Then
    estCols = Int((sheetW - 2 * margin + spacing) / (iMaxW + spacing))
    If estCols < 1 Then estCols = 1
    cols = estCols
  End If
  rows = rowsArg
  If rows < 1 Then rows = Int((tcount + cols - 1) / cols)

  ' Recompute with max sizes
  Dim totalGridW, topmost
  totalGridW = cols * iMaxW + (cols - 1) * spacing
  leftmost = margin
  If sheetW - 2 * margin > totalGridW Then leftmost = margin + (sheetW - 2 * margin - totalGridW) / 2
  topmost = sheetH - margin - iMaxH / 2

  WScript.Echo "Grid layout: " & cols & " cols x " & rows & " rows; maxW=" & iMaxW & ", maxH=" & iMaxH

  For i = 0 To tcount - 1
    Dim rowIdx, colIdx, centerX, centerY, idx
    rowIdx = Int(i / cols)
    colIdx = i Mod cols
    centerX = leftmost + colIdx * (iMaxW + spacing) + iMaxW / 2
    centerY = topmost - rowIdx * (iMaxH + spacing)
    ' Clamp center to keep view fully on-sheet (respect margin)
    minCX = margin + (targets(i).Width / 2)
    maxCX = sheetW - margin - (targets(i).Width / 2)
    minCY = margin + (targets(i).Height / 2)
    maxCY = sheetH - margin - (targets(i).Height / 2)
    adjCX = centerX
    adjCY = centerY
    centerNote = ""
    If adjCX < minCX Then adjCX = minCX : centerNote = " (clamped)"
    If adjCX > maxCX Then adjCX = maxCX : centerNote = " (clamped)"
    If adjCY < minCY Then adjCY = minCY : centerNote = " (clamped)"
    If adjCY > maxCY Then adjCY = maxCY : centerNote = " (clamped)"
    If previewMode Then
      WScript.Echo "Preview: view " & (i+1) & " -> (" & centerX & ", " & centerY & ")" & centerNote
    Else
      If debugMode Then
        On Error Resume Next
        Dim pAnch, pLocked, pMove
        pAnch = "n/a": pLocked = "n/a": pMove = "n/a"
        pAnch = targets(i).Anchored
        If Err.Number <> 0 Then Err.Clear : pAnch = "n/a"
        pLocked = targets(i).Locked
        If Err.Number <> 0 Then Err.Clear : pLocked = "n/a"
        pMove = targets(i).MoveableStatus
        If Err.Number <> 0 Then Err.Clear : pMove = "n/a"
        On Error GoTo 0
        WScript.Echo "DEBUG: Moving " & CStr(targets(i).Name) & " center request=(" & FormatNumber(adjCX,3) & "," & FormatNumber(adjCY,3) & ") Anchored=" & pAnch & " Locked=" & pLocked & " MoveableStatus=" & pMove
      End If
      oldX = targets(i).Center.X
      oldY = targets(i).Center.Y
      dx = adjCX - oldX
      dy = adjCY - oldY
      If Abs(dx) > 1e-9 Then targets(i).Position.TranslateBy tg.CreateVector2d(dx, 0)
      If Abs(dy) > 1e-9 Then targets(i).Position.TranslateBy tg.CreateVector2d(0, dy)
      ' Record movement and save undo record
      viewName = ""
      On Error Resume Next
      viewName = CStr(targets(i).Name)
      On Error GoTo 0
      newX = oldX + dx
      newY = oldY + dy
      ReDim Preserve updates(ucount)
      updates(ucount) = "View " & (i+1) & " [" & viewName & "] moved from (" & FormatNumber(oldX,3) & "," & FormatNumber(oldY,3) & ") to (" & FormatNumber(newX,3) & "," & FormatNumber(newY,3) & ")"
      ' Append to undo CSV
      If Not fso.FileExists(undoFilePath) Then
        Set outFS = fso.CreateTextFile(undoFilePath, True)
      Else
        Set outFS = fso.OpenTextFile(undoFilePath, 8)
      End If
      outFS.WriteLine viewName & "," & CStr(oldX) & "," & CStr(oldY)
      outFS.Close
      ' Verify move succeeded (read-back)
      Dim actualX, actualY
      actualX = targets(i).Center.X
      actualY = targets(i).Center.Y
      If Abs(actualX - newX) > 0.01 Or Abs(actualY - newY) > 0.01 Then
        ' Try one more time with TranslateBy
        If Abs(dx) > 1e-9 Then targets(i).Position.TranslateBy tg.CreateVector2d(dx, 0)
        If Abs(dy) > 1e-9 Then targets(i).Position.TranslateBy tg.CreateVector2d(0, dy)
        invApp.ActiveView.Update
        actualX = targets(i).Center.X
        actualY = targets(i).Center.Y
        If Abs(actualX - newX) > 0.01 Or Abs(actualY - newY) > 0.01 Then
          ' Fallback: try Move or SetCenter if available
          If debugMode Then WScript.Echo "DEBUG: Primary translate failed for " & viewName & ". Trying Move/SetCenter fallbacks..."
          On Error Resume Next
          Call targets(i).Move(tg.CreatePoint2d(newX, newY))
          Err.Clear
          Call targets(i).SetCenter(newX, newY)
          Err.Clear
          On Error GoTo 0
          invApp.ActiveView.Update
          actualX = targets(i).Center.X
          actualY = targets(i).Center.Y
          If Abs(actualX - newX) > 0.01 Or Abs(actualY - newY) > 0.01 Then
            ' Primary translate + fallback attempts failed
            ReDim Preserve updates(ucount)
            updates(ucount) = "FAILED to move View " & (i+1) & " [" & viewName & "] - expected (" & FormatNumber(newX,3) & "," & FormatNumber(newY,3) & ") but got (" & FormatNumber(actualX,3) & "," & FormatNumber(actualY,3) & ")"
            ucount = ucount + 1

            ' If requested, try to recreate view as a last resort
            If recreateMode Then
              If RecreateView(targets(i), newX, newY, debugMode) Then
                ReDim Preserve updates(ucount)
                updates(ucount) = "RECREATED View " & (i+1) & " [" & viewName & "] at (" & FormatNumber(newX,3) & "," & FormatNumber(newY,3) & ")"
                ucount = ucount + 1
              Else
                ReDim Preserve updates(ucount)
                updates(ucount) = "FAILED to RECREATE View " & (i+1) & " [" & viewName & "]"
                ucount = ucount + 1
              End If
            End If

          Else
            ReDim Preserve updates(ucount)
            updates(ucount) = "View " & (i+1) & " [" & viewName & "] moved (verified via fallback) to (" & FormatNumber(actualX,3) & "," & FormatNumber(actualY,3) & ")"
            ucount = ucount + 1
          End If
        Else
          ReDim Preserve updates(ucount)
          updates(ucount) = "View " & (i+1) & " [" & viewName & "] moved (verified) to (" & FormatNumber(actualX,3) & "," & FormatNumber(actualY,3) & ")"
          ucount = ucount + 1
        End If
      Else
        ReDim Preserve updates(ucount)
        updates(ucount) = "View " & (i+1) & " [" & viewName & "] moved (verified) to (" & FormatNumber(actualX,3) & "," & FormatNumber(actualY,3) & ")"
        ucount = ucount + 1
      End If
    End If
  Next

ElseIf layout = "vertical" Then
  ' Stack top-down centered horizontally
  sumX = 0
  For i = 0 To tcount - 1
    sumX = sumX + targets(i).Center.X
  Next
  avgX = sumX / tcount

  curY = 0
  ' Compute total height
  totalHeight = 0
  For i = 0 To tcount - 1
    totalHeight = totalHeight + targets(i).Height
  Next
  totalHeight = totalHeight + spacing * (tcount - 1)
  topEdge = sheetH - margin
  If sheetH - 2 * margin > totalHeight Then topEdge = margin + (sheetH - 2 * margin + totalHeight) / 2 ' center vertically
  curY = sheetH - margin - (targets(0).Height / 2)

  For i = 0 To tcount - 1
    If i = 0 Then
      desiredY = curY
    Else
      desiredY = curY - (targets(i-1).Height + targets(i).Height) / 2 - spacing
      curY = desiredY
    End If
    ' Clamp desired center into sheet bounds
    minCX = margin + (targets(i).Width / 2)
    maxCX = sheetW - margin - (targets(i).Width / 2)
    minCY = margin + (targets(i).Height / 2)
    maxCY = sheetH - margin - (targets(i).Height / 2)
    adjCX = avgX
    adjCY = desiredY
    centerNote = ""
    If adjCX < minCX Then adjCX = minCX : centerNote = " (clamped)"
    If adjCX > maxCX Then adjCX = maxCX : centerNote = " (clamped)"
    If adjCY < minCY Then adjCY = minCY : centerNote = " (clamped)"
    If adjCY > maxCY Then adjCY = maxCY : centerNote = " (clamped)"
    If previewMode Then
      WScript.Echo "Preview: view " & (i+1) & " -> (" & avgX & ", " & desiredY & ")" & centerNote
    Else
      ' If requested, attempt to unanchor/unlock/make moveable before moving
      If forceMode Then Call TryMakeViewMoveable(targets(i), debugMode)
      oldX = targets(i).Center.X
      oldY = targets(i).Center.Y
      dx2 = adjCX - oldX
      dy2 = adjCY - oldY
      If debugMode Then
        On Error Resume Next
        Dim propAnch, propLocked, propMoveStatus, propViewType
        propAnch = "n/a": propLocked = "n/a": propMoveStatus = "n/a": propViewType = "n/a"
        propAnch = targets(i).Anchored
        If Err.Number <> 0 Then Err.Clear : propAnch = "n/a"
        propLocked = targets(i).Locked
        If Err.Number <> 0 Then Err.Clear : propLocked = "n/a"
        propMoveStatus = targets(i).MoveableStatus
        If Err.Number <> 0 Then Err.Clear : propMoveStatus = "n/a"
        propViewType = targets(i).Type
        If Err.Number <> 0 Then Err.Clear : propViewType = "n/a"
        On Error GoTo 0
        WScript.Echo "DEBUG: Moving " & CStr(targets(i).Name) & " oldCenter=(" & FormatNumber(oldX,3) & "," & FormatNumber(oldY,3) & ") Width=" & FormatNumber(targets(i).Width,3) & " Height=" & FormatNumber(targets(i).Height,3) & " Anchored=" & propAnch & " Locked=" & propLocked & " MoveableStatus=" & propMoveStatus & " Type=" & propViewType
      End If
      If Abs(dx2) > 1e-9 Then targets(i).Position.TranslateBy tg.CreateVector2d(dx2, 0)
      If Abs(dy2) > 1e-9 Then targets(i).Position.TranslateBy tg.CreateVector2d(0, dy2)
      invApp.ActiveView.Update
      viewName = ""
      On Error Resume Next
      viewName = CStr(targets(i).Name)
      On Error GoTo 0
      newX = oldX + dx2
      newY = oldY + dy2
      ' Verify and fallback
      actualX = targets(i).Center.X
      actualY = targets(i).Center.Y
      If Abs(actualX - newX) > 0.01 Or Abs(actualY - newY) > 0.01 Then
        If debugMode Then WScript.Echo "DEBUG: Primary translate failed for " & viewName & ". Trying fallbacks..."
        On Error Resume Next
        Call targets(i).Move(tg.CreatePoint2d(newX, newY))
        Err.Clear
        Call targets(i).SetCenter(newX, newY)
        Err.Clear
        On Error GoTo 0
        invApp.ActiveView.Update
        actualX = targets(i).Center.X
        actualY = targets(i).Center.Y
        If Abs(actualX - newX) > 0.01 Or Abs(actualY - newY) > 0.01 Then
          ReDim Preserve updates(ucount)
          updates(ucount) = "FAILED to move View " & (i+1) & " [" & viewName & "] - expected (" & FormatNumber(newX,3) & "," & FormatNumber(newY,3) & ") but got (" & FormatNumber(actualX,3) & "," & FormatNumber(actualY,3) & ")"
        Else
          ReDim Preserve updates(ucount)
          updates(ucount) = "View " & (i+1) & " [" & viewName & "] moved (verified via fallback) to (" & FormatNumber(actualX,3) & "," & FormatNumber(actualY,3) & ")"
        End If
      Else
        ReDim Preserve updates(ucount)
        updates(ucount) = "View " & (i+1) & " [" & viewName & "] moved (verified) to (" & FormatNumber(actualX,3) & "," & FormatNumber(actualY,3) & ")"
      End If
      ucount = ucount + 1
    End If
  Next

Else
  ' Default: horizontal distribution (centered)
  ' Compute total width including spacing
  totalW = totalWidth + spacing * (tcount - 1)
  leftmost = margin
  If sheetW - 2 * margin > totalW Then leftmost = margin + (sheetW - 2 * margin - totalW) / 2
  firstCenter = leftmost + (targets(0).Width / 2)
  sumY = 0
  For i = 0 To tcount - 1
    sumY = sumY + targets(i).Center.Y
  Next
  avgY = sumY / tcount

  curCenterX = firstCenter
  For i = 0 To tcount - 1
    If i > 0 Then
      curCenterX = curCenterX + (targets(i-1).Width + targets(i).Width) / 2 + spacing
    End If
    ' Clamp curCenter into sheet bounds for this view
    minCX = margin + (targets(i).Width / 2)
    maxCX = sheetW - margin - (targets(i).Width / 2)
    minCY = margin + (targets(i).Height / 2)
    maxCY = sheetH - margin - (targets(i).Height / 2)
    adjCX = curCenterX
    adjCY = avgY
    centerNote = ""
    If adjCX < minCX Then adjCX = minCX : centerNote = " (clamped)"
    If adjCX > maxCX Then adjCX = maxCX : centerNote = " (clamped)"
    If adjCY < minCY Then adjCY = minCY : centerNote = " (clamped)"
    If adjCY > maxCY Then adjCY = maxCY : centerNote = " (clamped)"
    If previewMode Then
      WScript.Echo "Preview: view " & (i+1) & " -> (" & curCenterX & ", " & avgY & ")" & centerNote
    Else
      ' If requested, attempt to unanchor/unlock/make moveable before moving
      If forceMode Then Call TryMakeViewMoveable(targets(i), debugMode)
      oldX = targets(i).Center.X
      oldY = targets(i).Center.Y
      ddx = adjCX - oldX
      ddy = adjCY - oldY
      If debugMode Then
        On Error Resume Next
        Dim propAnch2, propLocked2
        propAnch2 = "n/a"
        propLocked2 = "n/a"
        propAnch2 = targets(i).Anchored
        If Err.Number <> 0 Then Err.Clear : propAnch2 = "n/a"
        propLocked2 = targets(i).Locked
        If Err.Number <> 0 Then Err.Clear : propLocked2 = "n/a"
        On Error GoTo 0
        WScript.Echo "DEBUG: Moving " & CStr(targets(i).Name) & " oldCenter=(" & FormatNumber(oldX,3) & "," & FormatNumber(oldY,3) & ") Anchored=" & propAnch2 & " Locked=" & propLocked2
      End If
      If Abs(ddx) > 1e-9 Then targets(i).Position.TranslateBy tg.CreateVector2d(ddx, 0)
      If Abs(ddy) > 1e-9 Then targets(i).Position.TranslateBy tg.CreateVector2d(0, ddy)
      invApp.ActiveView.Update
      viewName = ""
      On Error Resume Next
      viewName = CStr(targets(i).Name)
      On Error GoTo 0
      newX = oldX + ddx
      newY = oldY + ddy
      actualX = targets(i).Center.X
      actualY = targets(i).Center.Y
      If Abs(actualX - newX) > 0.01 Or Abs(actualY - newY) > 0.01 Then
        If debugMode Then WScript.Echo "DEBUG: Primary translate failed for " & viewName & ". Trying fallbacks..."
        On Error Resume Next
        Call targets(i).Move(tg.CreatePoint2d(newX, newY))
        Err.Clear
        Call targets(i).SetCenter(newX, newY)
        Err.Clear
        On Error GoTo 0
        invApp.ActiveView.Update
        actualX = targets(i).Center.X
        actualY = targets(i).Center.Y
        If Abs(actualX - newX) > 0.01 Or Abs(actualY - newY) > 0.01 Then
          ReDim Preserve updates(ucount)
          updates(ucount) = "FAILED to move View " & (i+1) & " [" & viewName & "] - expected (" & FormatNumber(newX,3) & "," & FormatNumber(newY,3) & ") but got (" & FormatNumber(actualX,3) & "," & FormatNumber(actualY,3) & ")"
        Else
          ReDim Preserve updates(ucount)
          updates(ucount) = "View " & (i+1) & " [" & viewName & "] moved (verified via fallback) to (" & FormatNumber(actualX,3) & "," & FormatNumber(actualY,3) & ")"
        End If
      Else
        ReDim Preserve updates(ucount)
        updates(ucount) = "View " & (i+1) & " [" & viewName & "] moved (verified) to (" & FormatNumber(actualX,3) & "," & FormatNumber(actualY,3) & ")"
      End If
      ucount = ucount + 1
    End If
  Next
End If

' Finish
If previewMode Then
  WScript.Echo "Preview complete. Computed positions for " & tcount & " views (no changes applied)."
Else
  doc.Update
  If saveMode Then
    On Error Resume Next
    doc.Save
    If Err.Number = 0 Then
      WScript.Echo "Document saved to disk." 
    Else
      WScript.Echo "Warning: could not save document: " & Err.Description
      Err.Clear
    End If
    On Error GoTo 0
  End If
  ' Force UI update and redraw
  On Error Resume Next
  invApp.ActiveView.Update
  WScript.Sleep 200
  invApp.ActiveView.Update
  On Error GoTo 0
  WScript.Echo "Alignment complete. Updated " & tcount & " views." & vbCrLf & "Sheet: " & sheet.Name & " (" & sheet.Width & " x " & sheet.Height & ")"
  If ucount = 0 Then
    WScript.Echo "No view movements were recorded."
  Else
    WScript.Echo "Moved views (" & ucount & "):"
    For i = 0 To ucount - 1
      WScript.Echo " - " & updates(i)
    Next
    WScript.Echo "Undo file written to: " & undoFilePath
  End If
End If