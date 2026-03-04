' Simple FL25 Parameter Check
Dim app, doc, occs, i, occ, subdoc
Set app = GetObject(, "Inventor.Application")
Set doc = app.ActiveDocument
Set occs = doc.ComponentDefinition.Occurrences

For i = 1 To occs.Count
    Set occ = occs.Item(i)
    If Not occ.Suppressed Then
        Set subdoc = occ.Definition.Document
        If InStr(1, subdoc.FullFileName, "FL25", 1) > 0 Then
            WScript.Echo "FL25 found"
            Dim cd, mp, p
            Set cd = subdoc.ComponentDefinition
            Set mp = cd.Parameters.ModelParameters
            Set p = mp.Item("d2")
            
            WScript.Echo "d2 parameter:"
            WScript.Echo "  Name: " & p.Name
            WScript.Echo "  Value: " & p.Value
            WScript.Echo "  ModelValue: " & p.ModelValue  
            WScript.Echo "  Units: '" & p.Units & "'"
            WScript.Echo "  UnitsType: " & p.UnitsType
            
            ' ModelValue is ALWAYS in base units (cm for Inventor)
            ' Convert to mm: cm * 10 = mm
            WScript.Echo ""
            WScript.Echo "Length in mm = " & (p.ModelValue * 10) & " mm"
            Exit For
        End If
    End If
Next
