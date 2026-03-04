' Check Specific FL Part - NSCR05-780-FL25
Dim app, doc, occs, i, occ, subdoc, fname

Set app = GetObject(, "Inventor.Application")
Set doc = app.ActiveDocument
Set occs = doc.ComponentDefinition.Occurrences

WScript.Echo "Searching " & occs.Count & " occurrences for FL25..."

For i = 1 To occs.Count
    Set occ = occs.Item(i)
    If Not occ.Suppressed Then
        Set subdoc = occ.Definition.Document
        fname = Mid(subdoc.FullFileName, InStrRev(subdoc.FullFileName, "\") + 1)
        
        If InStr(1, UCase(fname), "FL25", 1) > 0 Then
            WScript.Echo "FOUND: " & fname
            WScript.Echo "Full path: " & subdoc.FullFileName
            
            ' Check description
            On Error Resume Next
            Dim ps, dp, desc
            Set ps = subdoc.PropertySets.Item("Design Tracking Properties")
            If Err.Number = 0 Then
                Set dp = ps.Item("Description")
                If Err.Number = 0 Then
                    desc = dp.Value
                    WScript.Echo "Description: " & desc
                End If
            End If
            
            ' Check for PL or S355JR
            Dim isPlate
            isPlate = (InStr(1, UCase(desc), "PL", 1) > 0 Or InStr(1, UCase(desc), "S355JR", 1) > 0)
            WScript.Echo "Considered plate part: " & isPlate
            
            ' Check Length parameter
            Err.Clear
            Dim cd, params, lenParam
            Set cd = subdoc.ComponentDefinition
            Set params = cd.Parameters.UserParameters
            Set lenParam = params.Item("Length")
            
            If Err.Number = 0 Then
                WScript.Echo "HAS Length parameter: " & lenParam.Value
            Else
                WScript.Echo "NO Length parameter"
            End If
        End If
    End If
Next

WScript.Echo "Search complete"
