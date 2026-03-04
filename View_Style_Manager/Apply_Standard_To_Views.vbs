' Apply_Standard_To_Views.vbs
' Apply the document's active standard to all views by resetting their ViewStyle

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

WScript.Echo "=== APPLY STANDARD TO ALL VIEWS ==="
WScript.Echo "Drawing: " & idwDoc.DisplayName
WScript.Echo ""

' Get active standard
Dim stylesManager
Set stylesManager = idwDoc.StylesManager

Dim activeStandard
Set activeStandard = stylesManager.ActiveStandardStyle
WScript.Echo "Document Active Standard: " & activeStandard.Name
WScript.Echo ""

' ViewStyle enum values (from Inventor API):
' kFromBaseDrawingViewStyle = 32257
' kHiddenLineDrawingViewStyle = 32258  
' kHiddenLineRemovedDrawingViewStyle = 32259
' kShadedDrawingViewStyle = 32260
' kShadedHiddenLineDrawingViewStyle = 32261

Const kFromBaseDrawingViewStyle = 32257

WScript.Echo "Processing all views..."
WScript.Echo ""

Dim changedCount
changedCount = 0

Dim sheetNum
For sheetNum = 1 To idwDoc.Sheets.Count
    Dim sheet
    Set sheet = idwDoc.Sheets.Item(sheetNum)
    
    WScript.Echo "Sheet: " & sheet.Name
    
    Dim viewNum
    For viewNum = 1 To sheet.DrawingViews.Count
        Dim view
        Set view = sheet.DrawingViews.Item(viewNum)
        
        On Error Resume Next
        Err.Clear
        
        ' Get current ViewStyle
        Dim currentStyle
        currentStyle = view.ViewStyle
        
        WScript.Echo "  View: " & view.Name & " - Current ViewStyle: " & currentStyle
        
        ' Try to set it to kFromBaseDrawingViewStyle (inherit from standard)
        view.ViewStyle = kFromBaseDrawingViewStyle
        
        If Err.Number = 0 Then
            WScript.Echo "    -> Set to kFromBaseDrawingViewStyle (32257)"
            changedCount = changedCount + 1
        Else
            WScript.Echo "    -> ERROR: " & Err.Description
        End If
        Err.Clear
    Next
    WScript.Echo ""
Next

WScript.Echo "Changed " & changedCount & " views"
WScript.Echo ""

' Update the document
WScript.Echo "Updating document..."
On Error Resume Next
idwDoc.Update
If Err.Number = 0 Then
    WScript.Echo "Document updated successfully"
Else
    WScript.Echo "Update error: " & Err.Description
End If
Err.Clear

' Save
WScript.Echo "Saving document..."
idwDoc.Save2
If Err.Number = 0 Then
    WScript.Echo "Document saved!"
Else
    WScript.Echo "Save error: " & Err.Description
End If

WScript.Echo ""
WScript.Echo "=== COMPLETE ==="
WScript.Echo ""
WScript.Echo "All views have been set to use the document standard."
WScript.Echo "Check your views in Inventor to see if the style has been applied."
