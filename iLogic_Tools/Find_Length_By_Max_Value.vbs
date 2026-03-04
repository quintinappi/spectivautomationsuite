' Find Length by Maximum Parameter Value - For non-plate parts without Length
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
WScript.Echo "Finding length by maximum parameter value..."
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
                        
                        ' Find the parameter with maximum value
                        Dim maxValue, maxParamName, maxParamType
                        maxValue = 0
                        maxParamName = "(none)"
                        maxParamType = ""
                        
                        Err.Clear
                        Dim allParams, modelParams, j, param
                        Set allParams = cd.Parameters
                        
                        If Not allParams Is Nothing Then
                            ' Check model parameters
                            Set modelParams = allParams.ModelParameters
                            If Not modelParams Is Nothing Then
                                For j = 1 To modelParams.Count
                                    Err.Clear
                                    Set param = modelParams.Item(j)
                                    If Err.Number = 0 Then
                                        Dim paramModelValue
                                        paramModelValue = param.ModelValue
                                        
                                        ' Only consider mm/length parameters (ignore angles, unitless, etc)
                                        Dim paramUnits
                                        paramUnits = LCase(Trim(param.Units))
                                        
                                        If paramUnits = "mm" Or paramUnits = "" Then
                                            If paramModelValue > maxValue Then
                                                maxValue = paramModelValue
                                                maxParamName = param.Name
                                                maxParamType = "Model"
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                            
                            ' Check user parameters
                            If Not userParams Is Nothing Then
                                For j = 1 To userParams.Count
                                    Err.Clear
                                    Set param = userParams.Item(j)
                                    If Err.Number = 0 Then
                                        paramModelValue = param.ModelValue
                                        paramUnits = LCase(Trim(param.Units))
                                        
                                        If paramUnits = "mm" Or paramUnits = "" Then
                                            If paramModelValue > maxValue Then
                                                maxValue = paramModelValue
                                                maxParamName = param.Name
                                                maxParamType = "User"
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        
                        ' Convert from cm to mm (Inventor base units are cm)
                        Dim maxValueInMM
                        maxValueInMM = maxValue * 10
                        
                        WScript.Echo nonPlateNoLen & ". " & fname
                        WScript.Echo "   Description: " & desc
                        WScript.Echo "   Likely Length Parameter: " & maxParamName & " (" & maxParamType & ")"
                        WScript.Echo "   Model Value (cm): " & maxValue & " cm"
                        WScript.Echo "   Length (mm): " & maxValueInMM & " mm"
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
