On Error Resume Next

Dim invApp
Set invApp = GetObject(, "Inventor.Application")

Dim doc
Set doc = invApp.ActiveDocument

WScript.Echo "=== CHECKING CUSTOM PROPERTIES ==="
WScript.Echo "Part: " & doc.DisplayName
WScript.Echo ""

Dim customPropSet
Set customPropSet = doc.PropertySets.Item("Inventor User Defined Properties")

If Err.Number <> 0 Then
    WScript.Echo "ERROR: Could not get custom property set"
    WScript.Quit 1
End If

WScript.Echo "Total custom properties: " & customPropSet.Count
WScript.Echo ""

Dim prop, i
For i = 1 To customPropSet.Count
    Set prop = customPropSet.Item(i)
    
    WScript.Echo "Property " & i & ": " & prop.Name
    WScript.Echo "  Display Name: " & prop.DisplayName
    WScript.Echo "  Value: " & prop.Value
    WScript.Echo "  Expression: " & prop.Expression
    WScript.Echo "  Type: " & prop.PropType
    WScript.Echo ""
Next

WScript.Echo "=== TESTING PLATE WIDTH UPDATE ==="

Dim widthProp
On Error Resume Next
Set widthProp = customPropSet.Item("PLATE WIDTH")

If Err.Number <> 0 Then
    WScript.Echo "PLATE WIDTH doesn't exist - will try to add it"
    Err.Clear
    
    customPropSet.Add "=<SHEET METAL WIDTH>", "PLATE WIDTH"
    
    If Err.Number <> 0 Then
        WScript.Echo "ERROR adding: " & Err.Description
        WScript.Echo "Error number: " & Err.Number
    Else
        WScript.Echo "SUCCESS - PLATE WIDTH added"
    End If
Else
    WScript.Echo "PLATE WIDTH exists"
    WScript.Echo "Current value: " & widthProp.Value
    WScript.Echo "Current expression: " & widthProp.Expression
    WScript.Echo ""
    
    WScript.Echo "Trying to update..."
    widthProp.Value = "=<SHEET METAL WIDTH>"
    
    If Err.Number <> 0 Then
        WScript.Echo "ERROR updating: " & Err.Description
        WScript.Echo "Error number: " & Err.Number
    Else
        WScript.Echo "SUCCESS - PLATE WIDTH updated"
        doc.Update
        WScript.Echo "New value: " & widthProp.Value
        WScript.Echo "New expression: " & widthProp.Expression
    End If
End If
