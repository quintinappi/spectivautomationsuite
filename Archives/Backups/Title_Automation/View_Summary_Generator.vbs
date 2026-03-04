' ===============================================================================
' BASE VIEW SUMMARY - CLEAN LIST FORMAT
' ===============================================================================
' Lists sheets, base views, and current titles in clean summary format
' ===============================================================================

Option Explicit

Dim invApp
Set invApp = GetObject(, "Inventor.Application")

Sub ListBaseViewSummary()
    WScript.Echo "BASE VIEW SUMMARY"
    WScript.Echo "================="

    On Error Resume Next

    Dim drawingDoc
    Set drawingDoc = invApp.ActiveDocument

    If Err.Number <> 0 Then
        WScript.Echo "ERROR: No active document found"
        Exit Sub
    End If

    If drawingDoc.DocumentType <> 12292 Then
        WScript.Echo "ERROR: Active document is not a drawing"
        Exit Sub
    End If

    WScript.Echo "Drawing: " & drawingDoc.DisplayName
    WScript.Echo ""

    Dim sheets
    Set sheets = drawingDoc.Sheets

    Dim i
    For i = 1 To sheets.Count
        Dim sheet
        Set sheet = sheets.Item(i)

        WScript.Echo "SHEET " & i & ": " & sheet.Name
        WScript.Echo "----------------------------------------"

        Call ListSheetBaseViews(sheet)
        WScript.Echo ""
    Next
End Sub

Sub ListSheetBaseViews(sheet)
    Dim drawingViews
    Set drawingViews = sheet.DrawingViews

    Dim baseViewCount
    baseViewCount = 0

    Dim i
    For i = 1 To drawingViews.Count
        Dim view
        Set view = drawingViews.Item(i)

        If IsBaseView(view) Then
            baseViewCount = baseViewCount + 1
            WScript.Echo "  " & baseViewCount & ". " & view.Name

            ' Get current title/label
            On Error Resume Next
            Dim viewLabel
            Set viewLabel = view.Label
            If Err.Number = 0 And Not viewLabel Is Nothing Then
                WScript.Echo "     Current Title: """ & viewLabel.Text & """"
            Else
                WScript.Echo "     Current Title: [No Label]"
            End If
            Err.Clear
        End If
    Next

    If baseViewCount = 0 Then
        WScript.Echo "  No base views found"
    End If
End Sub

Function IsBaseView(view)
    IsBaseView = False

    On Error Resume Next
    Dim parentView
    Set parentView = Nothing

    Set parentView = view.ParentView

    If Err.Number <> 0 Then
        IsBaseView = True
        Err.Clear
        Exit Function
    ElseIf parentView Is Nothing Then
        IsBaseView = True
        Exit Function
    End If

    IsBaseView = False
End Function

Call ListBaseViewSummary()