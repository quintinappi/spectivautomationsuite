On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")
Set sheet = invApp.ActiveDocument.ActiveSheet

WScript.Echo "Style Default Diagnostic"
WScript.Echo "========================"

For Each v In sheet.DrawingViews
    WScript.Echo "View: " & v.Name
    
    Err.Clear
    isDef = v.IsStyleDefault
    If Err.Number = 0 Then
        WScript.Echo "  IsStyleDefault: " & isDef
    Else
        WScript.Echo "  IsStyleDefault: Error - " & Err.Description
    End If
    
    Err.Clear
    Set s = v.Style
    If Err.Number = 0 And Not s Is Nothing Then
        WScript.Echo "  Style.Name: " & s.Name
    Else
        ' Try common sub-styles
        Err.Clear
        Set s = v.ActiveStandard
        If Err.Number = 0 And Not s Is Nothing Then
            WScript.Echo "  ActiveStandard.Name: " & s.Name
        End If
    End If
    
    WScript.Echo "------------------------"
Next
