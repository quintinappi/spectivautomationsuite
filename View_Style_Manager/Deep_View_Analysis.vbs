' Deep_View_Analysis.vbs
' Comprehensive analysis of view properties and how to change them

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

WScript.Echo "=== DEEP VIEW ANALYSIS ==="
WScript.Echo "Drawing: " & idwDoc.DisplayName
WScript.Echo ""

Dim sheet
Set sheet = idwDoc.ActiveSheet

WScript.Echo "Analyzing first view in detail..."
WScript.Echo ""

If sheet.DrawingViews.Count = 0 Then
    WScript.Echo "No views on this sheet!"
    WScript.Quit 1
End If

Dim view
Set view = sheet.DrawingViews.Item(1)

WScript.Echo "View Name: " & view.Name
WScript.Echo "View Type: " & view.ViewType
WScript.Echo ""

' Try to enumerate ALL properties of the view object
WScript.Echo "Attempting to find style-related properties..."
WScript.Echo ""

' Check if it's a DrawingView with model reference
On Error Resume Next
Dim hasModel
hasModel = False

Dim refDoc
Set refDoc = Nothing
Set refDoc = view.ReferencedDocumentDescriptor
If Not refDoc Is Nothing Then
    hasModel = True
    WScript.Echo "View has model reference: YES"
    WScript.Echo "Referenced document: " & refDoc.DisplayName
Else
    WScript.Echo "View has model reference: NO (drafting view)"
End If
Err.Clear

WScript.Echo ""

' Try accessing view representation
On Error Resume Next
Dim viewRep
Set viewRep = Nothing
Set viewRep = view.ViewRepresentation
If Err.Number = 0 And Not viewRep Is Nothing Then
    WScript.Echo "ViewRepresentation: EXISTS"
    
    ' Try to get style from representation
    On Error Resume Next
    Dim repStyle
    Set repStyle = viewRep.Style
    If Err.Number = 0 And Not repStyle Is Nothing Then
        WScript.Echo "  ViewRepresentation.Style.Name = " & repStyle.Name
    Else
        WScript.Echo "  ViewRepresentation.Style = " & Err.Description
    End If
    Err.Clear
Else
    WScript.Echo "ViewRepresentation: " & Err.Description
End If
Err.Clear

WScript.Echo ""

' Try HiddenLineStyle
On Error Resume Next
Dim hlStyle
Set hlStyle = view.HiddenLineStyle
If Err.Number = 0 And Not hlStyle Is Nothing Then
    WScript.Echo "HiddenLineStyle: " & hlStyle.Name
Else
    WScript.Echo "HiddenLineStyle: " & Err.Description
End If
Err.Clear

' Try ShadedStyle  
On Error Resume Next
Dim shadedStyle
Set shadedStyle = view.ShadedStyle
If Err.Number = 0 And Not shadedStyle Is Nothing Then
    WScript.Echo "ShadedStyle: " & shadedStyle.Name
Else
    WScript.Echo "ShadedStyle: " & Err.Description
End If
Err.Clear

' Try DisplayStyle
On Error Resume Next
Dim dispStyle
dispStyle = view.DisplayStyle
If Err.Number = 0 Then
    WScript.Echo "DisplayStyle: " & dispStyle
Else
    WScript.Echo "DisplayStyle: " & Err.Description
End If
Err.Clear

WScript.Echo ""
WScript.Echo "Checking for Update method..."

' Try to call Update on the view
On Error Resume Next
view.Update
If Err.Number = 0 Then
    WScript.Echo "view.Update(): SUCCESS"
Else
    WScript.Echo "view.Update(): " & Err.Description
End If
Err.Clear

WScript.Echo ""
WScript.Echo "=== CHECKING STYLE OVERRIDE ==="

' The key might be in the Style and Standard Editor
' Views might need to have their style "overridden" from the document standard

' Try to access style override settings
On Error Resume Next
Dim styleOverride
styleOverride = view.StyleSourceOverride
If Err.Number = 0 Then
    WScript.Echo "StyleSourceOverride: " & styleOverride
Else
    WScript.Echo "StyleSourceOverride: " & Err.Description
End If
Err.Clear

WScript.Echo ""
WScript.Echo "=== ANALYSIS COMPLETE ==="
