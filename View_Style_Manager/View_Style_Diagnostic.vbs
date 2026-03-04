' View_Style_Diagnostic.vbs
' Quick diagnostic to see what style properties are available on views

Option Explicit

Dim invApp
On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")

If Err.Number <> 0 Or invApp Is Nothing Then
    WScript.Echo "ERROR: Inventor is not running!"
    WScript.Quit 1
End If

Dim idwDoc
Set idwDoc = invApp.ActiveDocument

If idwDoc Is Nothing Or idwDoc.DocumentType <> 12292 Then
    WScript.Echo "ERROR: No IDW file is open!"
    WScript.Quit 1
End If

WScript.Echo "=== DIAGNOSTIC: VIEW STYLE DETECTION ==="
WScript.Echo "Drawing: " & idwDoc.DisplayName
WScript.Echo ""

' Get active standard
On Error Resume Next
Dim activeStd
Set activeStd = idwDoc.StylesManager.ActiveStandardStyle
If Not activeStd Is Nothing Then
    WScript.Echo "Document Active Standard: " & activeStd.Name
Else
    WScript.Echo "Document Active Standard: (none detected)"
End If
WScript.Echo ""

' Check each view
Dim sheet
Set sheet = idwDoc.ActiveSheet

WScript.Echo "Sheet: " & sheet.Name
WScript.Echo "Number of views: " & sheet.DrawingViews.Count
WScript.Echo ""

Dim i
For i = 1 To sheet.DrawingViews.Count
    Dim view
    Set view = sheet.DrawingViews.Item(i)
    
    WScript.Echo "--- View " & i & ": " & view.Name & " ---"
    WScript.Echo "  ViewType: " & view.ViewType
    
    ' Try different properties
    On Error Resume Next
    Err.Clear
    
    ' Try view.Style
    Dim viewStyle
    Set viewStyle = Nothing
    Set viewStyle = view.Style
    If Err.Number = 0 And Not viewStyle Is Nothing Then
        WScript.Echo "  view.Style.Name = '" & viewStyle.Name & "'"
    ElseIf Err.Number <> 0 Then
        WScript.Echo "  view.Style = ERROR: " & Err.Description
    Else
        WScript.Echo "  view.Style = NULL"
    End If
    Err.Clear
    
    ' Try view.StyleName
    Dim styleName
    styleName = ""
    styleName = view.StyleName
    If Err.Number = 0 And styleName <> "" Then
        WScript.Echo "  view.StyleName = '" & styleName & "'"
    ElseIf Err.Number <> 0 Then
        WScript.Echo "  view.StyleName = ERROR: " & Err.Description
    Else
        WScript.Echo "  view.StyleName = (empty)"
    End If
    Err.Clear
    
    ' Try view.Standard (for standard views)
    Dim std
    Set std = Nothing
    Set std = view.Standard
    If Err.Number = 0 And Not std Is Nothing Then
        WScript.Echo "  view.Standard.Name = '" & std.Name & "'"
    ElseIf Err.Number <> 0 Then
        WScript.Echo "  view.Standard = ERROR: " & Err.Description
    Else
        WScript.Echo "  view.Standard = NULL"
    End If
    Err.Clear
    
    WScript.Echo ""
Next

WScript.Echo "=== END DIAGNOSTIC ==="

