' Auto_Balloon_Views.vbs
' Adds balloons to visible occurrences in each placed view on the active sheet.
' - Uses BOM item number when available
' - Skips occurrences already having a balloon in that view
' - Places balloon at centroid with a small offset to reduce overlap (basic)
' - Logs actions to AutoBalloonLog.txt next to the script
'
' Usage: run from a batch file as other scripts in this folder (cscript //nologo)
' NOTE: Test on a copy of your drawing first. This is a prototype and may
' need adjustments for complex drawings or advanced leader routing.

' Debug mode check (allow headless runs): pass argument "debug" to run without GUI prompts
Dim debugMode
debugMode = False
If WScript.Arguments.Count > 0 Then
    If LCase(CStr(WScript.Arguments(0))) = "debug" Then debugMode = True
    If LCase(CStr(WScript.Arguments(0))) = "headless" Then debugMode = True
End If

' Label method selection
Dim labelChoice, labelPrompt
labelPrompt = "Label method (enter number):" & vbCrLf & "1: BOM item number (preferred) - falls back to Part Number" & vbCrLf & "2: Part Number"
If Not debugMode Then
    ' We will ask this later if not debug, but define it here
    labelChoice = 1
Else
    labelChoice = 1
End If

' Prepare logging early
Dim fso, logPath, logFile
Set fso = CreateObject("Scripting.FileSystemObject")
logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\AutoBalloonLog.txt"
Set logFile = fso.OpenTextFile(logPath, 8, True)

' Helper: write to log
Sub Log(s)
    If Not logFile Is Nothing Then logFile.WriteLine Now & " - " & s
    If debugMode Then
        WScript.Echo s
    End If
End Sub

Log "--- Script Start ---"

On Error Resume Next
Dim invApp
Set invApp = GetObject(, "Inventor.Application")
If invApp Is Nothing Then
    MsgBox "ERROR: Cannot connect to Inventor. Make sure Inventor is running.", vbCritical, "Auto Balloon Views"
    WScript.Quit 1
End If

' Find an open drawing document
Dim idwDoc
Set idwDoc = Nothing
If Not invApp.ActiveDocument Is Nothing Then
    If invApp.ActiveDocument.DocumentType = 12292 Then ' Drawing
        Set idwDoc = invApp.ActiveDocument
    End If
End If

If idwDoc Is Nothing Then
    ' Search all open docs
    Dim docs, i
    Set docs = invApp.Documents
    For i = 1 To docs.Count
        If docs.Item(i).DocumentType = 12292 Then
            Set idwDoc = docs.Item(i)
            Exit For
        End If
    Next
End If

If idwDoc Is Nothing Then
    MsgBox "No drawing document found. Open the IDW/DWG you want to modify and try again.", vbExclamation, "Auto Balloon Views"
    WScript.Quit 1
End If

' Confirm active sheet
Dim sheet
Set sheet = idwDoc.ActiveSheet
If sheet Is Nothing Then
    MsgBox "ERROR: Could not get active sheet.", vbCritical, "Auto Balloon Views"
    WScript.Quit 1
End If

' Diagnostic: probe sheet for available annotation/text collections
On Error Resume Next
Dim probeCol, hasDrawingSketches, hasDrawingNotes, hasTextBoxes, hasDrawingTextBoxes, hasNotes
hasDrawingSketches = False
hasDrawingNotes = False
hasTextBoxes = False
hasDrawingTextBoxes = False
hasNotes = False
Set probeCol = Nothing
On Error Resume Next
Set probeCol = sheet.Sketches
If Err.Number = 0 And Not probeCol Is Nothing Then hasDrawingSketches = True
Err.Clear
On Error Resume Next
Set probeCol = sheet.DrawingNotes
If Err.Number = 0 And Not probeCol Is Nothing Then hasDrawingNotes = True
Err.Clear
On Error Resume Next
Set probeCol = sheet.TextBoxes
If Err.Number = 0 And Not probeCol Is Nothing Then hasTextBoxes = True
Err.Clear
On Error Resume Next
Set probeCol = sheet.DrawingTextBoxes
If Err.Number = 0 And Not probeCol Is Nothing Then hasDrawingTextBoxes = True
Err.Clear
On Error Resume Next
Set probeCol = sheet.Notes
If Err.Number = 0 And Not probeCol Is Nothing Then hasNotes = True
Err.Clear
Log "Sheet collections available - Sketches: " & CStr(hasDrawingSketches) & ", DrawingNotes: " & CStr(hasDrawingNotes) & ", TextBoxes: " & CStr(hasTextBoxes) & ", DrawingTextBoxes: " & CStr(hasDrawingTextBoxes) & ", Notes: " & CStr(hasNotes)

' Get balloon style collection and choose current default style
Dim bStyles, defStyle
Set bStyles = idwDoc.StylesManager.BalloonStyles
If bStyles.Count = 0 Then
    MsgBox "No balloon styles found in this drawing.", vbExclamation, "Auto Balloon Views"
    WScript.Quit 1
End If

' Prompt user to select balloon style (index or exact name). Cancel => use first style
Dim listText, inputStyle, idxStyle, sName
listText = "Available balloon styles in " & idwDoc.DisplayName & ":" & vbCrLf & vbCrLf
For i = 1 To bStyles.Count
    listText = listText & CStr(i) & ": " & bStyles.Item(i).Name & vbCrLf
Next
Dim forceSilentStyleSelection
forceSilentStyleSelection = False
If debugMode Then
    ' If a second argument of 'silent' is provided, suppress the style prompt.
    If WScript.Arguments.Count > 1 Then
        If LCase(CStr(WScript.Arguments(1))) = "silent" Then forceSilentStyleSelection = True
    End If
End If
If Not debugMode Or (debugMode And Not forceSilentStyleSelection) Then
    Dim runnerIsCscript
    runnerIsCscript = False
    On Error Resume Next
    If InStr(LCase(WScript.FullName), "cscript") > 0 Then runnerIsCscript = True
    Err.Clear
    If runnerIsCscript Then
        On Error Resume Next
        WScript.StdOut.WriteLine listText
        WScript.StdOut.Write "Enter index number or exact style name (blank uses first style): "
        inputStyle = ""
        On Error Resume Next
        inputStyle = WScript.StdIn.ReadLine
        If Err.Number <> 0 Then
            inputStyle = ""
            Log "Could not read console input for style selection; using default style"
            Err.Clear
        End If
    Else
        inputStyle = InputBox(listText & vbCrLf & "Enter index number or exact style name (Cancel to use first style):", "Select Balloon Style")
    End If
Else
    inputStyle = "" ' use default (first style) when debug+silent
End If
If inputStyle = "" Then
    Set defStyle = bStyles.Item(1)
Else
    Set defStyle = Nothing
    If IsNumeric(inputStyle) Then
        idxStyle = CInt(inputStyle)
        If idxStyle >= 1 And idxStyle <= bStyles.Count Then Set defStyle = bStyles.Item(idxStyle)
    End If
    If defStyle Is Nothing Then
        sName = Trim(inputStyle)
        For i = 1 To bStyles.Count
            If UCase(bStyles.Item(i).Name) = UCase(sName) Then
                Set defStyle = bStyles.Item(i)
                Exit For
            End If
        Next
    End If
    If defStyle Is Nothing Then Set defStyle = bStyles.Item(1)
End If


Log "--- Helper functions defined ---"

' Helper: check if an occurrence already has a balloon in a given view
' In standard API, balloons are at the sheet level, but contain information about their target.
Function HasBalloonForOccurrence(view, occ)
    HasBalloonForOccurrence = False
    On Error Resume Next
    Dim sheet, b
    Set sheet = view.Parent
    For b = 1 To sheet.Balloons.Count
        Dim balloon, intent
        Set balloon = sheet.Balloons.Item(b)
        ' A balloon can have multiple leaders/points; check the first leader's intent
        If balloon.Leader.RootNode.ChildNodes.Count > 0 Then
            Set intent = balloon.Leader.RootNode.ChildNodes.Item(1).AttachedEntity
            If Not intent Is Nothing Then
                ' If intent is an occurrence, compare names/ids
                If TypeName(intent) = "ComponentOccurrence" Or TypeName(intent) = "AssemblyOccurrence" Then
                    If intent.InternalName = occ.InternalName Then
                        HasBalloonForOccurrence = True
                        Exit For
                    End If
                End If
            End If
        End If
    Next
End Function

' Helper: basic centroid of occurrence geometry projected to view (best-effort)
Function GetOccurrenceCentroidInView(view, occ, refDoc)
    On Error Resume Next
    Dim bbox
    ' ComponentOccurrence.RangeBox is already in assembly space.
    Set bbox = occ.RangeBox
    Dim minPt, maxPt
    Set minPt = bbox.MinPoint
    Set maxPt = bbox.MaxPoint
    Dim cx, cy, cz
    cx = (minPt.X + maxPt.X) / 2
    cy = (minPt.Y + maxPt.Y) / 2
    cz = (minPt.Z + maxPt.Z) / 2

    Log "    Assembly-space centroid: (" & CStr(cx) & "," & CStr(cy) & "," & CStr(cz) & ")"

    ' Return a point in the application transient geometry (all Inventor docs share same TG scale/units)
    Set GetOccurrenceCentroidInView = invApp.TransientGeometry.CreatePoint(cx, cy, cz)
End Function

' Helper: get label text for an occurrence (BOM item number or part number)
Function GetLabelForOccurrence(occ, labelChoice)
    On Error Resume Next
    GetLabelForOccurrence = ""
    Dim partDoc, prop, temp

    If labelChoice = 1 Then
        ' Try to get an item number property or exposed item number
        On Error Resume Next
        temp = ""
        temp = occ.ItemNumber
        If Err.Number = 0 And Not IsEmpty(temp) Then
            GetLabelForOccurrence = CStr(temp)
            Exit Function
        End If
        Err.Clear
    End If

    ' Fallback to referenced document Part Number iProperty
    On Error Resume Next
    If Not occ.ReferencedDocument Is Nothing Then
        Set partDoc = occ.ReferencedDocument
        On Error Resume Next
        Set prop = Nothing
        Set prop = partDoc.PropertySets("Design Tracking Properties").Item("Part Number")
        If Err.Number = 0 And Not prop Is Nothing Then
            GetLabelForOccurrence = CStr(prop.Value)
            Exit Function
        End If
        Err.Clear
    End If

    ' Last resort: use occurrence name
    GetLabelForOccurrence = occ.Name
End Function

' Iterate all placed views on the active sheet
' Used points arrays for basic collision avoidance (per view)
Dim usedPointsX(), usedPointsY(), usedCount
ReDim usedPointsX(0)
ReDim usedPointsY(0)
usedCount = 0

Function IsNearExisting(x,y,scale)
    IsNearExisting = False
    On Error Resume Next
    Dim t, dx, dy, thresh
    thresh = (scale * 6) * (scale * 6)
    If usedCount = 0 Then Exit Function
    For t = 0 To usedCount - 1
        dx = usedPointsX(t) - x
        dy = usedPointsY(t) - y
        If (dx*dx + dy*dy) < thresh Then
            IsNearExisting = True
            Exit Function
        End If
    Next
End Function

Sub AddUsedPoint(x,y)
    If usedCount = 0 Then
        usedPointsX(0) = x
        usedPointsY(0) = y
        usedCount = 1
    Else
        Dim n
        n = usedCount
        ReDim Preserve usedPointsX(n)
        ReDim Preserve usedPointsY(n)
        usedPointsX(n) = x
        usedPointsY(n) = y
        usedCount = usedCount + 1
    End If
End Sub

Dim gv, viewsAdded, proceedResp, viewCount
viewCount = sheet.DrawingViews.Count
' Select single view or all views
Dim selectViewInput, singleViewIndex, singleViewName
singleViewIndex = -1
singleViewName = ""
If Not debugMode Then
    selectViewInput = InputBox("Enter a view index (1.." & CStr(viewCount) & ") or name to process a single view, or leave blank for all views:", "Select View")
Else
    selectViewInput = "1" ' debug -> process first view only
End If
If selectViewInput <> "" Then
    If IsNumeric(selectViewInput) Then
        singleViewIndex = CInt(selectViewInput)
        If singleViewIndex < 1 Or singleViewIndex > viewCount Then
            MsgBox "Invalid view index. Aborting.", vbExclamation, "Auto Balloon Views"
            logFile.WriteLine Now & " - Invalid view index provided. Aborting."
            logFile.Close
            WScript.Quit 1
        End If
    Else
        singleViewName = Trim(selectViewInput)
    End If
End If

' Per-view occurrence limit for quick tests (N=10 default in debug)
Dim perViewLimit
If Not debugMode Then
    Dim limitInput
    limitInput = InputBox("Enter per-view occurrence processing limit (leave blank for no limit):", "Per-view Limit", "")
    If limitInput = "" Then
        perViewLimit = 0
    ElseIf IsNumeric(limitInput) Then
        perViewLimit = CInt(limitInput)
    Else
        perViewLimit = 0
    End If
Else
    perViewLimit = 1 ' single-occurrence test
End If

If Not debugMode Then
    proceedResp = MsgBox("This will attempt to balloon visible parts in " & CStr(viewCount) & " views on sheet: " & sheet.Name & vbCrLf & "Continue?", vbYesNo + vbQuestion, "Auto Balloon Views")
    If proceedResp = vbNo Then
        Log "User cancelled before processing views."
        logFile.WriteLine Now & " - Cancelled by user."
        logFile.Close
        MsgBox "Cancelled.", vbInformation, "Auto Balloon Views"
        WScript.Quit 0
    End If
Else
    Log "Debug/headless mode: auto-accepting proceed confirmation; perViewLimit=" & CStr(perViewLimit)
End If

viewsAdded = 0
For i = 1 To sheet.DrawingViews.Count
    Dim view
    Set view = sheet.DrawingViews.Item(i)

    ' Determine whether to process this view (single-view selection support)
    Dim processThisView
    processThisView = True
    If singleViewIndex <> -1 Then
        If i <> singleViewIndex Then processThisView = False
    ElseIf singleViewName <> "" Then
        If LCase(view.Name) <> LCase(singleViewName) Then processThisView = False
    End If

    If processThisView Then
        Log "Processing view (" & CStr(i) & "/" & CStr(sheet.DrawingViews.Count) & "): " & view.Name
        ' Initialize used-points for this view (simple collision avoidance)
        usedCount = 0
        ReDim usedPointsX(0)
        ReDim usedPointsY(0)

        ' Log view type
        On Error Resume Next
        Dim viewType
        viewType = "Unknown"
        viewType = CStr(view.ViewType)
        Log "    View type: " & viewType
        Err.Clear

        ' Per-view occurrence attempt counter
        Dim occAttempted
        occAttempted = 0

    ' Determine the list of visible occurrences in this view.
    ' Approach: use ReferencedOccurrence for model-level drawing views or
    ' search for ComponentOccurrences in associated assembly.
    Dim occs
    Set occs = Nothing
    On Error Resume Next
    ' If view is an assembly or presentation view, it may have a ReferencedDocument or
    ' a ReferencedEntity. We'll try to query the underlying assembly and iterate its occurrences.
    ' Try ReferencedDocumentDescriptor first (works for many view types)
    On Error Resume Next
    Dim refDoc
    If Not view.ReferencedDocumentDescriptor Is Nothing Then
        Set refDoc = view.ReferencedDocumentDescriptor.ReferencedDocument
        If Not refDoc Is Nothing Then
            Log "    ReferencedDocumentDescriptor.Type: " & CStr(refDoc.DocumentType) & " TypeName: " & TypeName(refDoc)
            If refDoc.DocumentType = 12291 Then
                Set occs = refDoc.ComponentDefinition.Occurrences
                If Not occs Is Nothing Then Log "    Got occurrences from descriptor assembly: " & CStr(occs.Count)
            End If
        End If
    End If

    ' Fallback to ReferencedDocument property
    If occs Is Nothing Then
        If Not view.ReferencedDocument Is Nothing Then
            Set refDoc = view.ReferencedDocument
            Log "    ReferencedDocument.Type: " & CStr(refDoc.DocumentType) & " TypeName: " & TypeName(refDoc)
            ' If referencing an assembly or part file, get occurrences through the assembly document
            If refDoc.DocumentType = 12291 Then ' Assembly
                ' Try to get the component occurrences collection via view.ReferencedModel or used API
                ' Fallback: get assembly file from View.ReferencedEntity
                Set occs = refDoc.ComponentDefinition.Occurrences
                If Not occs Is Nothing Then Log "    Got occurrences from referenced assembly: " & CStr(occs.Count)
            End If
        End If
    End If

    ' If occs is still Nothing, try using view.ReferencedEntities and probe parent objects
    If occs Is Nothing Then
        On Error Resume Next
        Dim ents
        Set ents = view.ReferencedEntities
        If Not ents Is Nothing Then
            ' Build a transient collection of occurrences from the referenced entities
            Set occs = invApp.TransientObjects.CreateObjectCollection()
            Dim e
            For e = 1 To ents.Count
                Dim ent
                Set ent = ents.Item(e)
                Log "    RefEntity[" & CStr(e) & "]: " & TypeName(ent)
                On Error Resume Next
                Dim parentEnt
                Set parentEnt = Nothing
                Set parentEnt = ent.Parent
                If Err.Number = 0 And Not parentEnt Is Nothing Then
                    Log "       Parent: " & TypeName(parentEnt) & " - " & CStr(parentEnt.Name)
                    If TypeName(parentEnt) = "ComponentOccurrence" Or TypeName(parentEnt) = "AssemblyOccurrence" Then
                        occs.Add parentEnt
                    End If
                End If
            Next
            Log "    Derived occurrence count from referenced entities: " & CStr(occs.Count)
        Else
            Log "    view.ReferencedEntities is Nothing"
        End If
    End If

    If occs Is Nothing Then
        Log "  Could not get occurrences for view: " & view.Name & " (skipping)"
    Else
        On Error Resume Next
        Log "  Occurrence count in view: " & CStr(occs.Count)
        Err.Clear

    ' For each occurrence, check if it is visible in THIS view (View.IsVisible or GetVisibleStatus)
    Dim j
    For j = 1 To occs.Count
        Dim occ, skipOccurrence
        Set occ = occs.Item(j)
        skipOccurrence = False
        ' Respect per-view processing limit (debug quick-test)
        occAttempted = occAttempted + 1
        If perViewLimit > 0 And occAttempted > perViewLimit Then
            Log "  Reached per-view attempt limit (" & CStr(perViewLimit) & ") - stopping view processing."
            Exit For
        End If
        Log "    Attempting occurrence (" & CStr(occAttempted) & "/" & CStr(occs.Count) & "): " & occ.Name

        ' Basic visibility check: try VisibleInDrawingView and fallback
        Dim visible
        visible = False
        On Error Resume Next
        visible = occ.VisibleInDrawingView(view)
        If Err.Number <> 0 Then
            Err.Clear
            ' Could not call VisibleInDrawingView - assume visible (best-effort)
            visible = True
        End If
        Log "    Occurrence: " & occ.Name & " Visible: " & CStr(visible)

        If Not visible Then
            ' Skip non-visible occurrences
            skipOccurrence = True
        End If

        ' Skip if already has a balloon in this view
        If HasBalloonForOccurrence(view, occ) Then
            Log "  Skipping already-ballooned occurrence: " & occ.Name
            skipOccurrence = True
        End If

        ' Quick attempt: try to add a balloon directly using the occurrence (some view types accept this)
        On Error Resume Next
        Dim annDirect
        Set annDirect = Nothing
        Set annDirect = view.Annotations.AddBalloon(occ)
        If Err.Number = 0 And Not annDirect Is Nothing Then
            Log "    Annotations.AddBalloon(occ) succeeded for: " & occ.Name
            Dim directLabel
            directLabel = GetLabelForOccurrence(occ, labelChoice)
            On Error Resume Next
            If directLabel <> "" Then annDirect.Text = directLabel
            If Err.Number <> 0 Then
                Log "    Failed to set direct balloon text for: " & occ.Name & " - " & CStr(Err.Description)
                Err.Clear
            End If
            viewsAdded = viewsAdded + 1
            skipOccurrence = True
        Else
            Err.Clear
        End If

        ' Compute centroid and offset to place balloon leader origin
        Dim centroid3d
        Set centroid3d = GetOccurrenceCentroidInView(view, occ, refDoc)
        If centroid3d Is Nothing Then
            Log "  Could not compute centroid for occurrence: " & occ.Name & " (skipping)"
            skipOccurrence = True
        End If

        ' Project 3D centroid to drawing view coordinates
        Dim projPt, attempt
        Set projPt = Nothing
        attempt = 0
        On Error Resume Next
        Set projPt = view.ModelToSheetSpace(centroid3d)
        If Err.Number <> 0 Or projPt Is Nothing Then
            Log "  Initial ModelToSheetSpace failed for: " & occ.Name & " - " & CStr(Err.Description)
            Err.Clear
            ' Try alternative: create a point in the view's referenced document transient geometry with the same coordinates
            attempt = attempt + 1
            If Not refDoc Is Nothing Then
                On Error Resume Next
                Dim altPt
                Set altPt = Nothing
                If Not refDoc.ComponentDefinition Is Nothing Then
                    Set altPt = refDoc.ComponentDefinition.TransientGeometry.CreatePoint(centroid3d.X, centroid3d.Y, centroid3d.Z)
                Else
                    Set altPt = refDoc.TransientGeometry.CreatePoint(centroid3d.X, centroid3d.Y, centroid3d.Z)
                End If
                If Not altPt Is Nothing Then
                    On Error Resume Next
                    Set projPt = view.ModelToSheetSpace(altPt)
                    If Err.Number <> 0 Or projPt Is Nothing Then
                        Log "  Alt ModelToSheetSpace attempt #" & CStr(attempt) & " failed for: " & occ.Name & " - " & CStr(Err.Description)
                        Err.Clear
                    Else
                        Log "    Alt projection succeeded (refDoc point)"
                    End If
                End If
            End If
        Else
            Log "    Projected centroid to sheet coordinates: (" & CStr(projPt.X) & "," & CStr(projPt.Y) & ")"
        End If

        ' Final fallback: try with application transient geometry
        If projPt Is Nothing Then
            attempt = attempt + 1
            On Error Resume Next
            Dim altPt2
            Set altPt2 = invApp.TransientGeometry.CreatePoint(centroid3d.X, centroid3d.Y, centroid3d.Z)
            Set projPt = view.ModelToSheetSpace(altPt2)
            If Err.Number <> 0 Or projPt Is Nothing Then
                Log "  Final projection attempt failed for: " & occ.Name & " - " & CStr(Err.Description)
                Err.Clear
            Else
                Log "    Final projection succeeded (app point): (" & CStr(projPt.X) & "," & CStr(projPt.Y) & ")"
            End If
        End If

        ' If projection is still missing, try to derive 2D location from view.ReferencedEntities (if any)
        If projPt Is Nothing Then
            On Error Resume Next
            If Not view.ReferencedEntities Is Nothing Then
                foundRefEnt = False
                For idxRef = 1 To view.ReferencedEntities.Count
                    On Error Resume Next
                    Dim refEntity
                    Set refEntity = view.ReferencedEntities.Item(idxRef)
                    If Err.Number = 0 And Not refEntity Is Nothing Then
                        On Error Resume Next
                        Dim pParent
                        Set pParent = Nothing
                        Set pParent = refEntity.Parent
                        If Err.Number = 0 And Not pParent Is Nothing Then
                            If TypeName(pParent) = "ComponentOccurrence" Or TypeName(pParent) = "AssemblyOccurrence" Then
                                If pParent.InternalName = occ.InternalName Then
                                    ' Try to use the referenced entity rangebox midpoint as drawing coords
                                    On Error Resume Next
                                    Dim rbb, rmin, rmax, midX, midY
                                    Set rbb = refEntity.RangeBox
                                    If Err.Number = 0 And Not rbb Is Nothing Then
                                        Set rmin = rbb.MinPoint
                                        Set rmax = rbb.MaxPoint
                                        midX = (rmin.X + rmax.X) / 2
                                        midY = (rmin.Y + rmax.Y) / 2
                                        Log "    Using ReferencedEntity rangebox midpoint for projection: (" & CStr(midX) & "," & CStr(midY) & ")"
                                        Set projPt = invApp.TransientGeometry.CreatePoint(midX, midY, 0)
                                        foundRefEnt = True
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
                If Not foundRefEnt Then
                    Log "    No referenced entity rangebox found for occurrence: " & occ.Name
                End If
            Else
                Log "    view.ReferencedEntities is Nothing (cannot derive 2D point)"
            End If
        End If

        ' Last resort: place balloons in a grid near the view's position on the sheet
        If projPt Is Nothing Then
            On Error Resume Next
            Dim vPos
            Set vPos = Nothing
            Set vPos = view.Position
            Dim baseX, baseY
            If Err.Number <> 0 Or vPos Is Nothing Then
                Err.Clear
                ' Use sheet origin fallback
                baseX = 10 * view.Scale
                baseY = -10 * view.Scale
                Log "    Could not get view.Position; using sheet-based fallback coords"
            Else
                baseX = vPos.X
                baseY = vPos.Y
                Log "    Using view.Position as base for fallback grid: (" & CStr(baseX) & "," & CStr(baseY) & ")"
            End If
            ' Grid placement: columns of 10 items
            Dim fallbackIdx, fallbackCol, fallbackRow, fallbackGridX, fallbackGridY
            fallbackIdx = occAttempted - 1 ' zero-based
            fallbackCol = (fallbackIdx Mod 10)
            fallbackRow = Int(fallbackIdx / 10)
            fallbackGridX = baseX + (fallbackCol * 12 * view.Scale)
            fallbackGridY = baseY - (fallbackRow * 12 * view.Scale)
            Set projPt = invApp.TransientGeometry.CreatePoint(fallbackGridX, fallbackGridY, 0)
            Log "    Fallback grid placement for occurrence: (" & CStr(fallbackGridX) & "," & CStr(fallbackGridY) & ")"
        End If

        ' If we still have no projPt, skip occurrence
        If projPt Is Nothing Then
            Log "  Could not find any way to determine 2D projection for: " & occ.Name & " (skipping)"
            skipOccurrence = True
        Else
            ' If projPt was created as a 3D point but in drawing space, ensure we use X/Y
            Log "    Final projected coordinates (drawing space): (" & CStr(projPt.X) & "," & CStr(projPt.Y) & ")"
        End If

        If Not skipOccurrence Then
            Dim offsetX, offsetY, baseOffset, putX, putY, tries
            baseOffset = view.Scale * 2.0 ' small base offset, scales with view scale
            offsetX = baseOffset
            offsetY = 0
            putX = projPt.X
            putY = projPt.Y
        tries = 0
        Do While IsNearExisting(putX + offsetX, putY, view.Scale) And tries < 6
            tries = tries + 1
            offsetX = baseOffset * (1 + tries)
        Loop
        Dim balloonPt
        On Error Resume Next
        Dim balloonPt2d
        Set balloonPt2d = Nothing
        ' Prefer document's transient geometry for drawing 2D points
        On Error Resume Next
        If Not idwDoc.TransientGeometry Is Nothing Then
            Set balloonPt2d = idwDoc.TransientGeometry.CreatePoint2d(putX + offsetX, putY + offsetY)
        End If
        If Err.Number <> 0 Or balloonPt2d Is Nothing Then
            Err.Clear
            Set balloonPt2d = invApp.TransientGeometry.CreatePoint2d(putX + offsetX, putY + offsetY)
        End If
        If Not balloonPt2d Is Nothing Then
            Set balloonPt = balloonPt2d
            Log "    Created Point2d for balloon placement (doc or app tg)"
        Else
            Err.Clear
            Set balloonPt = invApp.TransientGeometry.CreatePoint(putX + offsetX, putY + offsetY, 0)
            Log "    Falling back to Point3d for balloon placement"
        End If
        ' remember this placement
        AddUsedPoint putX + offsetX, putY

        ' Create balloon annotation - COM API style
        On Error Resume Next
        Set ann = Nothing

        Dim leaderPoints
        Set leaderPoints = invApp.TransientObjects.CreateObjectCollection

        ' Add the balloon position point FIRST
        leaderPoints.Add balloonPt

        ' Try to get a geometry intent for attachment to the occurrence
        Dim curves, intent
        Set curves = view.DrawingCurves(occ)
        If Err.Number = 0 And Not curves Is Nothing Then
            If curves.Count > 0 Then
                Log "    Found " & curves.Count & " drawing curves for occurrence"
                Set intent = sheet.CreateGeometryIntent(curves.Item(1))
                If Not intent Is Nothing Then
                    Log "    Successfully created GeometryIntent for balloon"
                    ' Add intent LAST
                    leaderPoints.Add intent
                End If
            End If
        Else
            Log "    Failed to get drawing curves for occurrence: " & Err.Description
        End If
        Err.Clear

        ' Add the balloon
        On Error Resume Next
        Err.Clear
        Log "    Calling Balloons.Add with " & leaderPoints.Count & " points/intents"
        ' Signature: Add(LeaderPoints, [Target], [Level], [NumberingScheme], [BalloonStyle], [Layer])
        ' Attempt 1: Full arguments
        Set ann = sheet.Balloons.Add(leaderPoints, , , , defStyle)
        
        If Err.Number <> 0 Or ann Is Nothing Then
            Log "    Balloons.Add (Attempt 1) failed: " & Err.Description & " (Code: " & Hex(Err.Number) & ")"
            Err.Clear

            ' Attempt 2: Minimal arguments
            Log "    Balloons.Add (Attempt 2 - Minimal) starting..."
            Set ann = sheet.Balloons.Add(leaderPoints)
            If Not ann Is Nothing Then
                Log "    Balloons.Add (Attempt 2) succeeded"
                On Error Resume Next
                If Not defStyle Is Nothing Then ann.Style = defStyle
            Else
                Log "    Balloons.Add (Attempt 2) failed: " & Err.Description
                Err.Clear
            End If
        Else
            Log "    Balloons.Add (Attempt 1) succeeded"
        End If
        Err.Clear

        ' If we could not create a balloon, try TextBox fallback
        If ann Is Nothing Then
            Log "    Could not create a standard balloon for: " & occ.Name & " - attempting TextBox fallback"
            On Error Resume Next
            lbl = GetLabelForOccurrence(occ, labelChoice)
            If lbl = "" Then lbl = occ.Name
            Dim sketches, dsketch, textBox
            Set sketches = Nothing
            Set dsketch = Nothing
            Set textBox = Nothing
            On Error Resume Next
            Set sketches = sheet.Sketches
            If Err.Number = 0 And Not sketches Is Nothing Then
                On Error Resume Next
                Set dsketch = sketches.Add()
                If Err.Number = 0 And Not dsketch Is Nothing Then
                    On Error Resume Next
                    ' Use the sketch TextBoxes collection to add fitted text at the 2D point
                    Set textBox = dsketch.TextBoxes.AddFitted(balloonPt2d, lbl)
                    If Err.Number = 0 And Not textBox Is Nothing Then
                        Log "    TextBox.AddFitted succeeded for: " & occ.Name & " (label: " & lbl & ")"
                        viewsAdded = viewsAdded + 1
                        skipOccurrence = True
                    Else
                        Log "    TextBox.AddFitted failed: " & CStr(Err.Description)
                        Err.Clear
                    End If
                Else
                    Log "    Could not create DrawingSketch: " & CStr(Err.Description)
                    Err.Clear
                End If
            Else
                Log "    DrawingSketches collection unavailable on sheet (cannot add TextBox fallback)"
            End If
        Else
            ' Assign label text
            On Error Resume Next
            lbl = GetLabelForOccurrence(occ, labelChoice)
            If lbl <> "" Then
                On Error Resume Next
                ann.Text = lbl
                If Err.Number <> 0 Then
                    Log "    Failed to set balloon text for: " & occ.Name & " - " & CStr(Err.Description)
                    Err.Clear
                Else
                    Log "    Set balloon text: " & lbl
                End If
            End If

            Log "  Created balloon for occurrence: " & occ.Name & " (label: " & lbl & ") in view: " & view.Name
            viewsAdded = viewsAdded + 1
        End If
        End If
    Next
    End If
    ' Close per-view processing conditional
End If
Next

idwDoc.Save
logFile.WriteLine Now & " - Completed auto-ballooning. Balloons added: " & CStr(viewsAdded)
logFile.Close

MsgBox "Done. Balloons added: " & CStr(viewsAdded) & ". See log: " & logPath, vbInformation, "Auto Balloon Views"
WScript.Quit 0