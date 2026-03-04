' Scan_And_Sync_Master.vbs
' Scans views, identifies a Master View, analyzes its complete style configuration,
' and offers to dynamically sync other views to match.

Option Explicit

Dim invApp, doc
On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")
Set doc = invApp.ActiveDocument

If doc.DocumentType <> 12292 Then
    MsgBox "Active document is not a drawing!", vbCritical
    WScript.Quit
End If

Dim sheet
Set sheet = doc.ActiveSheet

' 1. LIST VIEWS
Dim viewList, i
viewList = "VIEWS ON ACTIVE SHEET:" & vbCrLf & vbCrLf
For i = 1 To sheet.DrawingViews.Count
    viewList = viewList & i & ": " & sheet.DrawingViews.Item(i).Name & vbCrLf
Next
viewList = viewList & vbCrLf & "Enter the NUMBER of the MASTER view:"

Dim choice
choice = InputBox(viewList, "Select Master View")
If choice = "" Or Not IsNumeric(choice) Then WScript.Quit

Dim masterIndex
masterIndex = CInt(choice)
If masterIndex < 1 Or masterIndex > sheet.DrawingViews.Count Then
    MsgBox "Invalid selection.", vbExclamation
    WScript.Quit
End If

Dim masterView
Set masterView = sheet.DrawingViews.Item(masterIndex)

' 2. ANALYZE MASTER VIEW STYLE
WScript.Echo "Analyzing Master View: " & masterView.Name & "..."

' A. Get View Style Enum
Dim masterViewStyle
masterViewStyle = masterView.ViewStyle
Dim viewStyleName
Select Case masterViewStyle
    Case 32257: viewStyleName = "From Base / Standard"
    Case 32258: viewStyleName = "Hidden Line"
    Case 32259: viewStyleName = "Hidden Line Removed"
    Case 32260: viewStyleName = "Shaded"
    Case 32261: viewStyleName = "Shaded Hidden Line"
    Case Else: viewStyleName = "Unknown (" & masterViewStyle & ")"
End Select

' B. Detect Layers used for Visible and Hidden Lines
Dim masterVisLayer, masterHidLayer
Set masterVisLayer = Nothing
Set masterHidLayer = Nothing

' Scan curves to find distinct layers for Solid vs Dashed geometry
Dim curves
Set curves = masterView.DrawingCurves
Dim c, s, l
For Each c In curves
    If c.Segments.Count > 0 Then
        Set l = c.Segments.Item(1).Layer
        ' Heuristic: Cyan usually Hidden, Black usually Visible in this context
        ' Or check line type if possible (Layer.LineType)
        ' 32258 = kHiddenEdge
        
        If c.EdgeType = 32258 Then
            If masterHidLayer Is Nothing Then Set masterHidLayer = l
        ElseIf c.EdgeType = 32257 Then
            If masterVisLayer Is Nothing Then Set masterVisLayer = l
        End If
        
        ' Fallback if types match but layers differ (e.g. 82695 issue)
        ' We assume Cyan is the target "Hidden" look the user likes
        If l.Color.Red = 0 And l.Color.Green = 255 And l.Color.Blue = 255 Then
            If masterHidLayer Is Nothing Then Set masterHidLayer = l
        ElseIf l.Color.Red = 0 And l.Color.Green = 0 And l.Color.Blue = 0 Then
            If masterVisLayer Is Nothing Then Set masterVisLayer = l
        End If
    End If
    If Not masterVisLayer Is Nothing And Not masterHidLayer Is Nothing Then Exit For
Next

Dim report
report = "MASTER VIEW ANALYSIS:" & vbCrLf
report = report & "Name: " & masterView.Name & vbCrLf
report = report & "View Style: " & viewStyleName & vbCrLf
If Not masterVisLayer Is Nothing Then
    report = report & "Visible Layer: " & masterVisLayer.Name & vbCrLf
Else
    report = report & "Visible Layer: (Could not identify)" & vbCrLf
End If
If Not masterHidLayer Is Nothing Then
    report = report & "Hidden Layer: " & masterHidLayer.Name & vbCrLf
Else
    report = report & "Hidden Layer: (Could not identify)" & vbCrLf
End If

report = report & vbCrLf & "Do you want to apply this styling to ALL other views on the sheet?"

Dim reply
reply = MsgBox(report, vbYesNo + vbQuestion, "Confirm Sync")

If reply = vbNo Then WScript.Quit

' 3. APPLY TO OTHER VIEWS
WScript.Echo "Applying styles..."

Dim v, count
count = 0 
For Each v In sheet.DrawingViews
    If v.Name <> masterView.Name Then
        WScript.Echo "Updating " & v.Name & "..."
        
        ' A. Match View Style Enum
        If v.ViewStyle <> masterViewStyle Then
            v.ViewStyle = masterViewStyle
        End If
        
        ' B. Match Layers (Dynamic Logic)
        Dim vCurves
        Set vCurves = v.DrawingCurves
        Dim updatedSegs
        updatedSegs = 0
        
        For Each c In vCurves
            For Each s In c.Segments
                ' If curve is DASHDED or HIDDEN type -> Apply Master Hidden Layer
                ' If curve is SOLID or VISIBLE type -> Apply Master Visible Layer
                ' We use text matching on CURRENT layer name to preserve intent
                
                Dim curName
                curName = UCase(s.Layer.Name)
                
                If Not masterHidLayer Is Nothing Then
                    If InStr(curName, "HIDDEN") > 0 Or InStr(curName, "DASHED") > 0 Then
                        s.Layer = masterHidLayer
                        updatedSegs = updatedSegs + 1
                    End If
                End If
                
                If Not masterVisLayer Is Nothing Then
                    If InStr(curName, "VISIBLE") > 0 Or InStr(curName, "CONT") > 0 Or InStr(curName, "DEFAULT") > 0 Then
                         s.Layer = masterVisLayer
                         updatedSegs = updatedSegs + 1
                    End If
                End If
                
                ' Handle the "All Cyan" mess if present:
                ' We can't distinguish purely by layer name if everything is already mapped wrong.
                ' But user said they reverted. So we assume names are distinctive (e.g. Visible vs Hidden).
            Next
        Next
        count = count + 1
    End If
Next

doc.Update
MsgBox "Sync Complete based on Master View: " & masterView.Name, vbInformation
