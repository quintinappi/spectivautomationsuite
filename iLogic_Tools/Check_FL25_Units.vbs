' Check Parameter Units - Debug FL25 specifically
Dim app, doc, occs, i, occ, subdoc, fname

Set app = GetObject(, "Inventor.Application")
Set doc = app.ActiveDocument
Set occs = doc.ComponentDefinition.Occurrences

WScript.Echo "Searching for FL25..."

For i = 1 To occs.Count
    On Error Resume Next
    Set occ = occs.Item(i)
    If Err.Number = 0 And Not occ.Suppressed Then
        Set subdoc = occ.Definition.Document
        fname = Mid(subdoc.FullFileName, InStrRev(subdoc.FullFileName, "\") + 1)
        
        If InStr(1, UCase(fname), "FL25", 1) > 0 Then
            WScript.Echo "Found: " & fname
            WScript.Echo ""
            
            Dim cd, modelParams, j, param
            Set cd = subdoc.ComponentDefinition
            Set modelParams = cd.Parameters.ModelParameters
            
            WScript.Echo "All Model Parameters with Units Info:"
            For j = 1 To modelParams.Count
                Err.Clear
                Set param = modelParams.Item(j)
                If Err.Number = 0 Then
                    WScript.Echo "  " & param.Name & ":"
                    WScript.Echo "    Value: " & param.Value
                    WScript.Echo "    ModelValue: " & param.ModelValue
                    WScript.Echo "    Units: " & param.Units
                    WScript.Echo "    UnitsType: " & param.UnitsType
                    
                    ' Try to convert to mm explicitly
                    Err.Clear
                    Dim valueInMM
                    If param.Units = "mm" Then
                        valueInMM = param.ModelValue
                    ElseIf param.Units = "cm" Then
                        valueInMM = param.ModelValue * 10
                    Else
                        valueInMM = param.ModelValue
                    End If
                    WScript.Echo "    Converted to mm: " & valueInMM
                    WScript.Echo ""
                End If
            Next
            
            Exit For
        End If
    End If
Next
