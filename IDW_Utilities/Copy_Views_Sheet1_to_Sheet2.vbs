' ==================================================================================
' COPY VIEWS - SHEET 1 TO SHEET 2
' ==================================================================================
' Copies all views from Sheet 1 to Sheet 2 in the active IDW drawing
'
' Features:
' - Scans Sheet 1 for all views (base, sections, details, projected)
' - Copies each view to Sheet 2 with smart positioning
' - Preserves view scale and orientation
' - Works with any drawing (active IDW required)
'
' Usage:
' 1. Open IDW drawing in Inventor
' 2. Run this script
' 3. Views will be copied from Sheet 1 to Sheet 2
'
' ==================================================================================

Option Explicit

' Global variables
Dim invApp, invDoc
Dim fso, logFile
Dim logPath

' Main execution
Main

Sub Main()
    On Error Resume Next

    ' Setup logging
    Set fso = CreateObject("Scripting.FileSystemObject")
    logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\Copy_Views_Sheet1_to_Sheet2.log"
    Set logFile = fso.CreateTextFile(logPath, True)

    LogMessage "=== COPY VIEWS - SHEET 1 TO SHEET 2 ==="
    LogMessage "Starting view copy operation..."

    ' Connect to Inventor
    Set invApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        LogMessage "ERROR: Could not connect to Inventor. Please make sure Inventor is running."
        MsgBox "ERROR: Could not connect to Inventor." & vbCrLf & _
               "Please make sure Inventor is running with an IDW file open.", vbCritical, "Inventor Not Found"
        WScript.Quit
    End If
    LogMessage "SUCCESS: Connected to Inventor"

    ' Get active document
    Set invDoc = invApp.ActiveDocument
    If Err.Number <> 0 Or invDoc Is Nothing Then
        LogMessage "ERROR: No active document found."
        MsgBox "ERROR: No active document found." & vbCrLf & _
               "Please open an IDW drawing file.", vbCritical, "No Document"
        WScript.Quit
    End If

    ' Check if it's a drawing document
    If invDoc.DocumentType <> 12294 Then ' kDrawingDocumentObject
        LogMessage "ERROR: Active document is not a drawing (IDW/DWG)."
        MsgBox "ERROR: Active document is not a drawing file." & vbCrLf & _
               "Please open an IDW or DWG drawing.", vbCritical, "Wrong Document Type"
        WScript.Quit
    End If
    LogMessage "SUCCESS: Active document is a drawing: " & invDoc.DisplayName

    ' Check if Sheet 1 exists
    If invDoc.Sheets.Count < 1 Then
        LogMessage "ERROR: Drawing has no sheets!"
        MsgBox "ERROR: This drawing has no sheets!", vbCritical, "No Sheets"
        WScript.Quit
    End If

    ' Check if Sheet 2 exists, if not create it
    Dim sourceSheet, targetSheet
    Set sourceSheet = invDoc.Sheets.Item(1)
    LogMessage "Sheet 1 found: " & sourceSheet.Name & " (" & sourceSheet.DrawingViews.Count & " views)"

    If invDoc.Sheets.Count < 2 Then
        LogMessage "Sheet 2 does not exist. Creating new sheet..."
        Set targetSheet = invDoc.Sheets.Add()
        targetSheet.Name = "Sheet:2"
        LogMessage "Created Sheet 2: " & targetSheet.Name
    Else
        Set targetSheet = invDoc.Sheets.Item(2)
        LogMessage "Sheet 2 found: " & targetSheet.Name & " (" & targetSheet.DrawingViews.Count & " existing views)"
    End If

    ' Check if Sheet 1 has any views
    If sourceSheet.DrawingViews.Count = 0 Then
        LogMessage "ERROR: Sheet 1 has no views to copy!"
        MsgBox "Sheet 1 has no views to copy!" & vbCrLf & _
               "Please add views to Sheet 1 first.", vbExclamation, "No Views"
        WScript.Quit
    End If

    ' Copy views from Sheet 1 to Sheet 2
    LogMessage ""
    LogMessage "Starting view copy operation..."
    LogMessage "=================================="

    Dim viewsCopied, viewsFailed
    viewsCopied = 0
    viewsFailed = 0

    Dim i, sourceView, newView, position
    Dim x, y, spacing

    ' Position settings for placing views on Sheet 2
    x = 3          ' Start 3cm from left edge
    y = 25         ' Start 25cm from bottom (leaving room for title block)
    spacing = 15   ' 15cm spacing between views

    For i = 1 To sourceSheet.DrawingViews.Count
        Set sourceView = sourceSheet.DrawingViews.Item(i)

        LogMessage ""
        LogMessage "View [" & i & "/" & sourceSheet.DrawingViews.Count & "]: " & sourceView.Name
        LogMessage "  Type: " & GetViewTypeName(sourceView)
        LogMessage "  Scale: " & sourceView.ScaleString
        LogMessage "  Position: (" & FormatNumber(sourceView.Position.X, 2) & ", " & FormatNumber(sourceView.Position.Y, 2) & ")"

        ' Calculate position for new view
        Dim positionObj
        Set positionObj = invApp.TransientGeometry.CreatePoint2d(x, y)

        ' Copy the view using CopyTo method
        On Error Resume Next
        Set newView = sourceView.CopyTo(targetSheet)
        If Err.Number <> 0 Or newView Is Nothing Then
            LogMessage "  ERROR: Failed to copy view - " & Err.Description
            viewsFailed = viewsFailed + 1
            Err.Clear
        Else
            ' Position the copied view
            newView.Position = positionObj
            LogMessage "  SUCCESS: Copied and positioned at (" & x & ", " & y & ")"

            viewsCopied = viewsCopied + 1

            ' Update position for next view (place below)
            y = y - spacing
            If y < 5 Then
                ' If we've gone too low, move to right column
                y = 25
                x = x + 20
            End If
        End If
        On Error GoTo 0
    Next

    ' Save the drawing
    LogMessage ""
    LogMessage "=================================="
    LogMessage "Saving drawing..."
    invDoc.Save2 True
    LogMessage "SUCCESS: Drawing saved"

    ' Final report
    LogMessage ""
    LogMessage "=== COPY COMPLETE ==="
    LogMessage "Total views processed: " & sourceSheet.DrawingViews.Count
    LogMessage "Views copied: " & viewsCopied
    LogMessage "Views failed: " & viewsFailed

    logFile.Close

    ' Show message to user
    Dim msg
    msg = "View Copy Operation Complete!" & vbCrLf & vbCrLf
    msg = msg & "Source: Sheet 1 (" & sourceSheet.Name & ")" & vbCrLf
    msg = msg & "Target: Sheet 2 (" & targetSheet.Name & ")" & vbCrLf & vbCrLf
    msg = msg & "Views copied: " & viewsCopied & vbCrLf
    If viewsFailed > 0 Then
        msg = msg & "Views failed: " & viewsFailed & vbCrLf & vbCrLf
        msg = msg & "Check log file for details:" & vbCrLf & logPath
    End If

    MsgBox msg, vbInformation, "Copy Complete"

End Sub

' Helper function to get view type name
Function GetViewTypeName(view)
    Dim typeName
    On Error Resume Next

    ' Check if it's a base view (no parent)
    If view.ParentView Is Nothing Then
        typeName = "Base View"
    Else
        ' Check the type of parent relationship
        Dim parentView
        Set parentView = view.ParentView

        ' Check if it's a section view
        If view.Type = 103620641 Then ' kSectionDrawingViewType
            typeName = "Section View"
        ' Check if it's a detail view
        ElseIf view.Type = 103620643 Then ' kDetailDrawingViewType
            typeName = "Detail View"
        ' Check if it's a projected view
        ElseIf view.Type = 103620639 Then ' kProjectedDrawingViewType
            typeName = "Projected View"
        ' Check if it's an auxiliary view
        ElseIf view.Type = 103620645 Then ' kAuxiliaryDrawingViewType
            typeName = "Auxiliary View"
        ' Check if it's a break view
        ElseIf view.Type = 103620647 Then ' kBreakDrawingViewType
            typeName = "Break View"
        Else
            typeName = "Derived View"
        End If
    End If

    GetViewTypeName = typeName
End Function

Sub LogMessage(msg)
    Dim timestamp
    timestamp = Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & " " & _
                Right("0" & Hour(Now), 2) & ":" & Right("0" & Minute(Now), 2) & ":" & Right("0" & Second(Now), 2)

    If Not logFile Is Nothing Then
        logFile.WriteLine timestamp & " - " & msg
    End If
End Sub
