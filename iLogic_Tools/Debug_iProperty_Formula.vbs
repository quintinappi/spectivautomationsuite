' Debug script to check what's happening with LENGTH iProperty on a part
' Run with one of the plate parts open in Inventor

Option Explicit

Dim invApp

Sub Main()
    On Error Resume Next
    
    Set invApp = GetObject(, "Inventor.Application")
    If invApp Is Nothing Then
        WScript.Echo "Inventor not running"
        Exit Sub
    End If
    
    If invApp.ActiveDocument Is Nothing Then
        WScript.Echo "No active document"
        Exit Sub
    End If
    
    Dim doc
    Set doc = invApp.ActiveDocument
    
    WScript.Echo "=== DOCUMENT INFO ==="
    WScript.Echo "Name: " & doc.DisplayName
    WScript.Echo "Type: " & doc.DocumentType
    WScript.Echo "SubType: " & doc.SubType
    WScript.Echo ""
    
    ' Get custom properties
    Dim propSets
    Set propSets = doc.PropertySets
    
    Dim customProps
    Set customProps = propSets.Item("Inventor User Defined Properties")
    
    WScript.Echo "=== CUSTOM iPROPERTIES ==="
    Dim prop
    For Each prop In customProps
        WScript.Echo "Name: " & prop.Name
        WScript.Echo "  Value: " & prop.Value
        
        ' Check if it has Expression property
        Dim expr
        expr = prop.Expression
        If Err.Number = 0 Then
            WScript.Echo "  Expression: " & expr
        Else
            WScript.Echo "  Expression: (not available)"
            Err.Clear
        End If
        
        ' Check ValueType
        Dim valType
        valType = prop.Type
        If Err.Number = 0 Then
            WScript.Echo "  Type: " & valType
        Else
            Err.Clear
        End If
        WScript.Echo ""
    Next
    
    WScript.Echo ""
    WScript.Echo "=== TRYING TO SET LENGTH FORMULA ==="
    
    ' Get or create LENGTH property
    Dim lengthProp
    Set lengthProp = customProps.Item("LENGTH")
    
    If Err.Number <> 0 Then
        WScript.Echo "LENGTH property not found"
        Err.Clear
        Exit Sub
    End If
    
    WScript.Echo "Current LENGTH value: " & lengthProp.Value
    
    ' Try different formula syntaxes
    WScript.Echo ""
    WScript.Echo "Trying: Expression = '<sheet metal length>'"
    lengthProp.Expression = "<sheet metal length>"
    If Err.Number <> 0 Then
        WScript.Echo "  FAILED: " & Err.Description
        Err.Clear
    Else
        WScript.Echo "  SUCCESS! New value: " & lengthProp.Value
    End If
    
    WScript.Echo ""
    WScript.Echo "Trying: Value = '=<sheet metal length>'"
    lengthProp.Value = "=<sheet metal length>"
    If Err.Number <> 0 Then
        WScript.Echo "  FAILED: " & Err.Description
        Err.Clear
    Else
        WScript.Echo "  SUCCESS! New value: " & lengthProp.Value
    End If
    
    ' Show what the value looks like now
    WScript.Echo ""
    WScript.Echo "Final LENGTH property:"
    WScript.Echo "  Value: " & lengthProp.Value
    expr = lengthProp.Expression
    If Err.Number = 0 Then
        WScript.Echo "  Expression: " & expr
    Else
        WScript.Echo "  Expression: (not available)"
        Err.Clear
    End If
    
End Sub

Main
