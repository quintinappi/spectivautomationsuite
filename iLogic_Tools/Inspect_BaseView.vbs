Option Explicit
On Error Resume Next
Dim invApp
Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Or invApp Is Nothing Then
    WScript.Echo "ERROR: Inventor not running"
    WScript.Quit 1
End If
Err.Clear

If invApp.ActiveDocument Is Nothing Or invApp.ActiveDocument.DocumentType <> 12292 Then
    WScript.Echo "ERROR: Active document is not a drawing"
    WScript.Quit 1
End If

Dim drawDoc, sheet, views, v
Set drawDoc = invApp.ActiveDocument
Set sheet = drawDoc.Sheets.Item(1)
WScript.Echo "Sheet: " & sheet.Name & " (" & sheet.Width & " x " & sheet.Height & ")"
WScript.Echo "DrawingViews.Count = " & sheet.DrawingViews.Count

If sheet.DrawingViews.Count = 0 Then
    WScript.Echo "No views to inspect"
    WScript.Quit 0
End If

' Prefer ISO1 if present
Dim baseView
Set baseView = Nothing
For Each v In sheet.DrawingViews
    On Error Resume Next
    If LCase(Trim(v.Name)) = "iso1" Then
        Set baseView = v
        Exit For
    End If
Next

If baseView Is Nothing Then
    Set baseView = sheet.DrawingViews.Item(1)
    WScript.Echo "ISO1 not found; using first view: " & baseView.Name
Else
    WScript.Echo "Found base view: " & baseView.Name
End If

' Report known properties
On Error Resume Next
WScript.Echo "TypeName: " & TypeName(baseView)
WScript.Echo "Name: " & baseView.Name
WScript.Echo "Scale: " & baseView.Scale
WScript.Echo "Orientation: " & baseView.Orientation
WScript.Echo "ViewType: " & baseView.ViewType

' Try ReferencedDocumentDescriptor
On Error Resume Next
Dim rdd
Set rdd = baseView.ReferencedDocumentDescriptor
If Err.Number = 0 And Not rdd Is Nothing Then
    WScript.Echo "ReferencedDocumentDescriptor available"
    On Error Resume Next
    WScript.Echo "  ReferencedFile: " & rdd.ReferencedFileDescriptor.DisplayName
    WScript.Echo "  FullPath: " & rdd.ReferencedFileDescriptor.FullFileName
    Err.Clear
Else
    WScript.Echo "ReferencedDocumentDescriptor not available or error"
    Err.Clear
End If

' Check if the view is a base view by attempting to read BaseView-specific property
On Error Resume Next
Dim hasParent
hasParent = False
If baseView.ProjectionOf Is Nothing Then
    ' ProjectionOf may be not set for base views
Else
    hasParent = True
End If
If Err.Number <> 0 Then
    Err.Clear
End If
WScript.Echo "Has ProjectionOf (parent)?: " & CStr(hasParent)

' Dump some available member names (best-effort)
On Error Resume Next
WScript.Echo "Members (selected):"
On Error Resume Next
WScript.Echo "  BoundingBox: " & TypeName(baseView.BoundingBox)
If Err.Number <> 0 Then Err.Clear
On Error Resume Next
WScript.Echo "  ModelDocument: " & TypeName(baseView.ModelDocument)
If Err.Number <> 0 Then Err.Clear
On Error Resume Next
WScript.Echo "  IsBaseView (exists): " & CStr(baseView.IsBaseView)
If Err.Number <> 0 Then Err.Clear

WScript.Echo "Inspection complete."