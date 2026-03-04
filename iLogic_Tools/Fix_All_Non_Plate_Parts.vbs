' Fix_All_Non_Plate_Parts.vbs - DETAILING WORKFLOW STEP 4: Add Length2 property
' DETAILING WORKFLOW - STEP 4: Add Length2 property for non-plate parts
' Fix All Non-Plate Parts - Add Length2 parameter to all parts missing Length
' Processes entire assembly automatically

Option Explicit

' Inventor API Constants
Const kAssemblyDocumentObject = 12291
Const kPartDocumentObject = 12290
Const kNumberParameterType = 1
Const kMillimeterLengthUnits = 11269

Dim m_InventorApp
Dim m_Log

Sub Main()
    On Error Resume Next
    
    m_Log = ""
    LogMessage "=========================================="
    LogMessage "FIX ALL NON-PLATE PARTS - ADD LENGTH2"
    LogMessage "=========================================="
    LogMessage ""
    
    ' Get Inventor
    Set m_InventorApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        LogMessage "ERROR: Inventor not running"
        MsgBox "Inventor is not running. Please start Inventor first.", vbCritical
        WScript.Quit 1
    End If
    
    ' Get active document
    Dim asmDoc
    Set asmDoc = m_InventorApp.ActiveDocument
    If asmDoc Is Nothing Then
        LogMessage "ERROR: No active document"
        MsgBox "No active document. Please open an assembly.", vbCritical
        WScript.Quit 1
    End If
    
    If asmDoc.DocumentType <> kAssemblyDocumentObject Then
        LogMessage "ERROR: Not an assembly"
        MsgBox "Please open an ASSEMBLY (.iam), not a part.", vbExclamation
        WScript.Quit 1
    End If
    
    LogMessage "Assembly: " & asmDoc.DisplayName
    LogMessage ""
    
    ' Step 1: Scan assembly for non-plate parts without Length
    LogMessage "STEP 1: Scanning assembly for non-plate parts without Length parameter..."
    LogMessage ""
    
    Dim nonPlateParts
    Set nonPlateParts = CreateObject("Scripting.Dictionary")
    
    ScanAssemblyForNonPlates asmDoc, nonPlateParts
    
    If nonPlateParts.Count = 0 Then
        LogMessage "No non-plate parts found without Length parameter"
        MsgBox "No parts found that need fixing.", vbInformation
        WScript.Quit 0
    End If
    
    LogMessage ""
    LogMessage "Found " & nonPlateParts.Count & " parts to fix"
    LogMessage ""
    
    ' Step 2: Process each part
    LogMessage "STEP 2: Processing parts..."
    LogMessage ""
    
    Dim successCount, skipCount, failCount
    successCount = 0
    skipCount = 0
    failCount = 0
    
    Dim partPath
    For Each partPath In nonPlateParts.Keys
        LogMessage "----------------------------------------"
        LogMessage "Processing: " & partPath
        
        If FixPartLength2(partPath) Then
            successCount = successCount + 1
            LogMessage "SUCCESS"
        Else
            failCount = failCount + 1
            LogMessage "FAILED"
        End If
        
        LogMessage ""
        WScript.Sleep 500
    Next
    
    ' Step 3: Summary
    LogMessage "=========================================="
    LogMessage "PROCESSING COMPLETE"
    LogMessage "=========================================="
    LogMessage "Successfully fixed: " & successCount
    LogMessage "Failed: " & failCount
    LogMessage "Total: " & nonPlateParts.Count
    LogMessage ""
    LogMessage "NOTE: Export checkbox must be manually enabled for each part"
    LogMessage "      (API does not support this for regular part parameters)"
    
    MsgBox "Processing complete!" & vbCrLf & vbCrLf & _
           "Successfully fixed: " & successCount & vbCrLf & _
           "Failed: " & failCount & vbCrLf & _
           "Total: " & nonPlateParts.Count & vbCrLf & vbCrLf & _
           "NOTE: You must manually enable the Export checkbox" & vbCrLf & _
           "for Length2 in each part's Parameters dialog.", vbInformation, "Complete"
    
    SaveLog

End Sub

Sub ScanAssemblyForNonPlates(asmDoc, partsList)
    ' Recursively scan assembly for non-plate parts without Length parameter
    Dim occ
    For Each occ In asmDoc.ComponentDefinition.Occurrences
        On Error Resume Next
        
        Dim refDoc
        Set refDoc = occ.Definition.Document
        
        If Not refDoc Is Nothing Then
            ' Check if it's a sub-assembly
            If refDoc.DocumentType = kAssemblyDocumentObject Then
                ScanAssemblyForNonPlates refDoc, partsList
            ElseIf refDoc.DocumentType = kPartDocumentObject Then
                ' Get description from iProperty
                Dim description
                description = GetDescriptionFromIProperty(refDoc)
                
                ' Skip plates (contains "PL " - with space after)
                If InStr(UCase(description), "PL ") = 0 Then
                    ' Check if it has Length or Length2 parameter
                    If Not HasLengthParameter(refDoc) Then
                        Dim fullPath
                        fullPath = refDoc.FullFileName
                        
                        If Not partsList.Exists(fullPath) Then
                            partsList.Add fullPath, description
                            LogMessage "  Found: " & refDoc.DisplayName & " (" & description & ")"
                        End If
                    End If
                End If
            End If
        End If
        
        Err.Clear
    Next
End Sub

Function GetDescriptionFromIProperty(doc)
    On Error Resume Next
    GetDescriptionFromIProperty = ""
    
    Dim propSet
    Set propSet = doc.PropertySets.Item("Design Tracking Properties")
    If Err.Number = 0 Then
        Dim prop
        Set prop = propSet.Item("Description")
        If Err.Number = 0 Then
            GetDescriptionFromIProperty = prop.Value
        End If
    End If
    Err.Clear
End Function

Function HasLengthParameter(doc)
    On Error Resume Next
    HasLengthParameter = False
    
    Dim userParams
    Set userParams = doc.ComponentDefinition.Parameters.UserParameters
    
    If userParams.Count = 0 Then
        Exit Function
    End If
    
    ' Loop through all parameters and check names explicitly
    Dim i, param
    For i = 1 To userParams.Count
        Err.Clear
        Set param = userParams.Item(i)
        If Err.Number = 0 Then
            If UCase(param.Name) = "LENGTH" Or UCase(param.Name) = "LENGTH2" Then
                HasLengthParameter = True
                Exit Function
            End If
        End If
    Next
    
    Err.Clear
End Function

Function FixPartLength2(partPath)
    ' Fix individual part by adding Length2 parameter
    FixPartLength2 = False
    
    On Error Resume Next
    
    ' Open the part
    LogMessage "  Opening part..."
    Dim partDoc
    Set partDoc = m_InventorApp.Documents.Open(partPath, False)
    
    If Err.Number <> 0 Or partDoc Is Nothing Then
        LogMessage "  ERROR: Failed to open - " & Err.Description
        Exit Function
    End If
    
    ' Get parameters
    Dim compDef, params, modelParams, userParams
    Set compDef = partDoc.ComponentDefinition
    Set params = compDef.Parameters
    Set modelParams = params.ModelParameters
    Set userParams = params.UserParameters
    
    ' Find the largest parameter (length dimension)
    LogMessage "  Finding length parameter..."
    Dim maxValue, maxParamName, maxParam
    maxValue = 0
    maxParamName = ""
    Set maxParam = Nothing
    
    Dim i, param
    For i = 1 To modelParams.Count
        Err.Clear
        Set param = modelParams.Item(i)
        If Err.Number = 0 Then
            Dim paramUnits
            paramUnits = LCase(Trim(param.Units))
            If paramUnits = "mm" Or paramUnits = "" Then
                If param.ModelValue > maxValue Then
                    maxValue = param.ModelValue
                    maxParamName = param.Name
                    Set maxParam = param
                End If
            End If
        End If
    Next
    
    If maxParam Is Nothing Then
        LogMessage "  ERROR: Could not find length parameter"
        partDoc.Close False
        Exit Function
    End If
    
    LogMessage "  Found: " & maxParamName & " = " & (maxParam.ModelValue * 10) & "mm"
    
    ' Check if Length2 already exists
    Err.Clear
    Dim length2Param
    Set length2Param = userParams.Item("Length2")
    If Err.Number = 0 Then
        LogMessage "  Length2 already exists - updating it..."
        length2Param.Delete
        Err.Clear
    End If
    
    ' Create Length2 parameter
    LogMessage "  Creating Length2 parameter..."
    Err.Clear
    Set length2Param = userParams.AddByValue("Length2", maxParam.Value, kMillimeterLengthUnits)
    
    If Err.Number <> 0 Then
        LogMessage "  ERROR creating Length2: " & Err.Description
        On Error Resume Next
        partDoc.Close False
        Exit Function
    End If
    
    LogMessage "  Created Length2 = " & length2Param.Value & " " & length2Param.Units
    
    ' Link Length2 to model parameter
    LogMessage "  Linking Length2 = " & maxParamName
    Err.Clear
    length2Param.Expression = maxParamName
    
    If Err.Number <> 0 Then
        LogMessage "  ERROR setting expression: " & Err.Description
        partDoc.Close False
        Exit Function
    End If
    
    ' Link model parameter to Length2
    LogMessage "  Linking " & maxParamName & " = Length2"
    Err.Clear
    maxParam.Expression = "Length2"
    
    If Err.Number <> 0 Then
        LogMessage "  ERROR setting " & maxParamName & " expression: " & Err.Description
        partDoc.Close False
        Exit Function
    End If
    
    ' Save and close
    LogMessage "  Saving part..."
    Err.Clear
    partDoc.Save
    
    If Err.Number <> 0 Then
        LogMessage "  ERROR: Failed to save - " & Err.Description
        On Error Resume Next
        partDoc.Close False
        Exit Function
    End If
    
    On Error Resume Next
    partDoc.Close True
    LogMessage "  Completed successfully"
    
    FixPartLength2 = True
End Function

Sub LogMessage(msg)
    m_Log = m_Log & msg & vbCrLf
    WScript.Echo msg
End Sub

Sub SaveLog()
    On Error Resume Next
    Dim fso, logFile, logFolder
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    logFolder = fso.GetSpecialFolder(5) & "\Inventor_Logs" ' My Documents
    If Not fso.FolderExists(logFolder) Then
        fso.CreateFolder logFolder
    End If
    
    Dim logPath
    logPath = logFolder & "\Fix_All_Non_Plate_Parts_" & Replace(Replace(Replace(Now, "/", "-"), ":", "-"), " ", "_") & ".log"
    
    Set logFile = fso.CreateTextFile(logPath, True)
    logFile.WriteLine m_Log
    logFile.Close
    
    WScript.Echo "Log saved to: " & logPath
End Sub

' Run main
Main()
