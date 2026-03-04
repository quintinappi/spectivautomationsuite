' Scanner with ALL Parameters - Shows model values for non-plate parts without Length
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
                
                ' Check Length param in USER parameters
                Err.Clear
                Dim cd, userParams, lenParam
                Set cd = subdoc.ComponentDefinition
                Set userParams = cd.Parameters.UserParameters
                Set lenParam = userParams.Item("Length")
                hasLen = (Err.Number = 0)
                
                If Not isPlate Then
                    nonPlateCount = nonPlateCount + 1
                    If Not hasLen Then
                        nonPlateNoLen = nonPlateNoLen + 1
                        WScript.Echo nonPlateNoLen & ". " & fname
                        WScript.Echo "   Description: " & desc
                        WScript.Echo "   ALL Parameters (Model + User):"
                        
                        ' List ALL parameters (ModelParameters + UserParameters)
                        Err.Clear
                        Dim allParams, j, param
                        Set allParams = cd.Parameters
                        
                        If Not allParams Is Nothing Then
                            ' Show model parameters
                            Dim modelParams
                            Set modelParams = allParams.ModelParameters
                            If Not modelParams Is Nothing And modelParams.Count > 0 Then
                                WScript.Echo "      Model Parameters (" & modelParams.Count & "):"
                                For j = 1 To modelParams.Count
                                    Err.Clear
                                    Set param = modelParams.Item(j)
                                    If Err.Number = 0 Then
                                        Dim paramName, paramValue, paramModelValue, paramUnits
                                        paramName = param.Name
                                        paramValue = param.Value
                                        paramUnits = param.Units
                                        
                                        ' Get model value
                                        Err.Clear
                                        paramModelValue = param.ModelValue
                                        If Err.Number <> 0 Then
                                            paramModelValue = paramValue
                                        End If
                                        
                                        WScript.Echo "         " & paramName & " = " & paramValue & " " & paramUnits & " (Model: " & paramModelValue & ")"
                                    End If
                                Next
                            Else
                                WScript.Echo "      Model Parameters: (none)"
                            End If
                            
                            ' Show user parameters
                            If Not userParams Is Nothing And userParams.Count > 0 Then
                                WScript.Echo "      User Parameters (" & userParams.Count & "):"
                                For j = 1 To userParams.Count
                                    Err.Clear
                                    Set param = userParams.Item(j)
                                    If Err.Number = 0 Then
                                        paramName = param.Name
                                        paramValue = param.Value
                                        paramUnits = param.Units
                                        
                                        Err.Clear
                                        paramModelValue = param.ModelValue
                                        If Err.Number <> 0 Then
                                            paramModelValue = paramValue
                                        End If
                                        
                                        WScript.Echo "         " & paramName & " = " & paramValue & " " & paramUnits & " (Model: " & paramModelValue & ")"
                                    End If
                                Next
                            Else
                                WScript.Echo "      User Parameters: (none)"
                            End If
                        Else
                            WScript.Echo "      (could not access parameters)"
                        End If
                        
                        WScript.Echo ""
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
