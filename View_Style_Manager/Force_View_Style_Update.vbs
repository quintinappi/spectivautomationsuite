' Force_View_Style_Update.vbs
' Try different methods to force views to use the document standard

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

WScript.Echo "=== FORCE VIEW STYLE UPDATE ==="
WScript.Echo "Drawing: " & idwDoc.DisplayName
WScript.Echo ""

' Get the active standard
Dim stylesManager
Set stylesManager = idwDoc.StylesManager

Dim activeStandard
Set activeStandard = stylesManager.ActiveStandardStyle
WScript.Echo "Document Active Standard: " & activeStandard.Name
WScript.Echo ""

' Process all sheets
Dim sheetNum
For sheetNum = 1 To idwDoc.Sheets.Count
    Dim sheet
    Set sheet = idwDoc.Sheets.Item(sheetNum)
    
    WScript.Echo "Sheet: " & sheet.Name & " (" & sheet.DrawingViews.Count & " views)"
    
    Dim viewNum
    For viewNum = 1 To sheet.DrawingViews.Count
        Dim view
        Set view = sheet.DrawingViews.Item(viewNum)
        
        WScript.Echo "  View: " & view.Name
        
        ' Method 1: Try to delete and recreate the view (too destructive)
        ' Method 2: Try to access the view's style settings through the sheet
        ' Method 3: Try to force update through document
        
        ' Let's try accessing the view's curves and their styles
        On Error Resume Next
        
        ' Try to get edge display style
        Dim curves
        Set curves = view.DrawingCurves
        If Err.Number = 0 And Not curves Is Nothing Then
            WScript.Echo "    DrawingCurves count: " & curves.Count
            
            ' Check if we can access curve styles
            If curves.Count > 0 Then
                Dim curve
                Set curve = curves.Item(1)
                
                On Error Resume Next
                Dim curveStyle
                Set curveStyle = curve.Style
                If Err.Number = 0 And Not curveStyle Is Nothing Then
                    WScript.Echo "    First curve style: " & curveStyle.Name
                Else
                    WScript.Echo "    Curve.Style: " & Err.Description
                End If
                Err.Clear
            End If
        Else
            WScript.Echo "    DrawingCurves: " & Err.Description
        End If
        Err.Clear
        
        ' Try to access view's style source type
        On Error Resume Next
        Dim styleSource
        styleSource = view.StyleSourceType
        If Err.Number = 0 Then
            WScript.Echo "    StyleSourceType: " & styleSource
            ' 1 = kFromStandard, 2 = kOverrideStyle
            
            If styleSource = 2 Then
                WScript.Echo "    -> View has OVERRIDDEN style (not using document standard)"
                WScript.Echo "    -> Attempting to reset to document standard..."
                
                ' Try to set it back to use standard
                view.StyleSourceType = 1 ' kFromStandard
                If Err.Number = 0 Then
                    WScript.Echo "    -> SUCCESS! Reset to use document standard"
                Else
                    WScript.Echo "    -> ERROR: " & Err.Description
                End If
                Err.Clear
            Else
                WScript.Echo "    -> View is using document standard"
            End If
        Else
            WScript.Echo "    StyleSourceType: " & Err.Description
        End If
        Err.Clear
        
        WScript.Echo ""
    Next
Next

WScript.Echo "Saving document..."
On Error Resume Next
idwDoc.Save2
If Err.Number = 0 Then
    WScript.Echo "Document saved!"
Else
    WScript.Echo "Save error: " & Err.Description
End If

WScript.Echo ""
WScript.Echo "=== COMPLETE ==="
