' Verbose Scanner - Shows everything
Dim app, doc, occs, i, occ, subdoc, fname, desc, isPlate, hasLen
Dim nonPlateCount, nonPlateNoLen, seen

Set app = GetObject(, "Inventor.Application")
Set doc = app.ActiveDocument
Set occs = doc.ComponentDefinition.Occurrences
Set seen = CreateObject("Scripting.Dictionary")

nonPlateCount = 0
nonPlateNoLen = 0

WScript.Echo "Scanning " & occs.Count & " occurrences..."
WScript.Echo ""

For i = 1 To occs.Count
    On Error Resume Next
    Set occ = occs.Item(i)
    If Err.Number = 0 And Not occ.Suppressed Then
        Set subdoc = occ.Definition.Document
        fname = Mid(subdoc.FullFileName, InStrRev(subdoc.FullFileName, "\") + 1)
        
        If LCase(Right(fname, 4)) = ".ipt" Then
            If Not seen.Exists(subdoc.FullFileName) Then
                seen.Add subdoc.FullFileName, True
                
                ' Get description
                Err.Clear
                Dim ps, dp
                Set ps = subdoc.PropertySets.Item("Design Tracking Properties")
                If Err.Number = 0 Then
                    Set dp = ps.Item("Description")
                    If Err.Number = 0 Then
                        desc = Trim(dp.Value)
                    Else
                        desc = ""
                    End If
                Else
                    desc = ""
                End If
                
                ' Check if plate
                isPlate = (InStr(1, UCase(desc), "PL", 1) > 0 Or InStr(1, UCase(desc), "S355JR", 1) > 0)
                
                ' Check Length param
                Err.Clear
                Dim cd, params, lenParam
                Set cd = subdoc.ComponentDefinition
                Set params = cd.Parameters.UserParameters
                Set lenParam = params.Item("Length")
                hasLen = (Err.Number = 0)
                
                If Not isPlate Then
                    nonPlateCount = nonPlateCount + 1
                    If Not hasLen Then
                        nonPlateNoLen = nonPlateNoLen + 1
                        WScript.Echo nonPlateNoLen & ". " & fname
                        WScript.Echo "   Desc: " & desc
                    End If
                End If
            End If
        End If
    End If
Next

WScript.Echo ""
WScript.Echo "=========================================="
WScript.Echo "Non-plate parts: " & nonPlateCount
WScript.Echo "Non-plate WITHOUT Length: " & nonPlateNoLen
WScript.Echo "=========================================="
