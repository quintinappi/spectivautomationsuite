On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")
Set doc = invApp.ActiveDocument
Set sheet = doc.ActiveSheet

WScript.Echo "Deep Style Diagnostic"
WScript.Echo "===================="

For Each v In sheet.DrawingViews
    WScript.Echo "View Name: " & v.Name
    
    ' Check property "Style" type
    Err.Clear
    Set s = v.Style
    If Err.Number = 0 Then
        If Not s Is Nothing Then
            WScript.Echo "  Style Type: " & TypeName(s)
            WScript.Echo "  Style Name: " & s.Name
        Else
            WScript.Echo "  Style is Nothing"
        End If
    Else
        WScript.Echo "  Error accessing Style property: " & Err.Description
    End If
    
    ' Check ViewStyle enum
    WScript.Echo "  ViewStyle Enum: " & v.ViewStyle
    
    ' Check for internal standard name
    Err.Clear
    Set std = v.Standard
    If Err.Number = 0 And Not std Is Nothing Then
        WScript.Echo "  Standard Name: " & std.Name
    End If

    WScript.Echo "--------------------"
Next
