' Simple test - just create the parameter
Dim app, doc, cd, params, up, newParam

Set app = GetObject(, "Inventor.Application")
Set doc = app.ActiveDocument
Set cd = doc.ComponentDefinition
Set params = cd.Parameters
Set up = params.UserParameters

WScript.Echo "Part: " & doc.DisplayName

' Delete if exists
On Error Resume Next
Dim existing
Set existing = up.Item("Length2")
If Err.Number = 0 Then
    WScript.Echo "Deleting existing Length2..."
    existing.Delete
End If
Err.Clear

' Try method 1: Add with type
WScript.Echo "Method 1: userParams.Add(name, type)..."
Set newParam = up.Add("Length2", 1)
If Err.Number <> 0 Then
    WScript.Echo "  FAILED: " & Err.Description
    Err.Clear
Else
    WScript.Echo "  SUCCESS"
    newParam.Value = 94.6
    newParam.Units = "mm"
    newParam.Expression = "d2"
    newParam.ExportParameter = True
    WScript.Echo "  Length2 created = " & newParam.Value & " " & newParam.Units
    WScript.Quit 0
End If

' Try method 2: AddByValue
WScript.Echo "Method 2: AddByValue..."
On Error Resume Next
Set newParam = up.AddByValue("Length2", 94.6, 11269)
If Err.Number <> 0 Then
    WScript.Echo "  FAILED: " & Err.Description
Else
    WScript.Echo "  SUCCESS: " & newParam.Value
End If
