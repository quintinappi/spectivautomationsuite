' Reset_To_Default.vbs
' Attempts to clear manual layer overrides on segments so the View decides Hidden vs Visible.

Option Explicit

Dim invApp, doc
On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")
Set doc = invApp.ActiveDocument

WScript.Echo "=== RESETTING LAYERS TO DEFAULT ==="

Dim sheet
Set sheet = doc.Sheets.Item(2)
Dim targetNames, viewName
targetNames = Array("ELEVATION", "C", "D", "PLAN")

For Each viewName In targetNames
    Dim targetView
    Set targetView = Nothing
    Dim v
    For Each v In sheet.DrawingViews
        If InStr(UCase(v.Name), viewName) > 0 Then
            Set targetView = v
            Exit For
        End If
    Next
    
    If Not targetView Is Nothing Then
        WScript.Echo "Resetting: " & targetView.Name
        
        Dim curves
        Set curves = targetView.DrawingCurves
        Dim c, s
        Dim count
        count = 0
        
        For Each c In curves
            For Each s In c.Segments
                ' Attempt to clear override by setting to Nothing
                ' Note: In VBScript, 'Nothing' assignment to COM property can be tricky.
                ' We use 'Set s.Layer = Nothing' structure.
                
                On Error Resume Next
                Set s.Layer = Nothing
                If Err.Number <> 0 Then
                    ' Sometimes API requires a specific method or fails.
                    ' If failure, we assume it's locked or not supported.
                    Err.Clear
                Else
                    count = count + 1
                End If
            Next
        Next
        WScript.Echo "  -> Cleared overrides on " & count & " segments."
    End If
Next

doc.Update
WScript.Echo "Done!"
