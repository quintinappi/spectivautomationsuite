' =========================================================
' THICKNESS PARAMETER EXPORTER - FIXED VERSION
' =========================================================
' DETAILING WORKFLOW STEP 5b: Enable export for plate thickness properties
' 
' This script:
' 1. Scans the active assembly for PLATE parts
' 2. Opens each plate part
' 3. Finds the "Thickness" parameter (in ModelParameters OR UserParameters)
' 4. Enables export for that parameter (checks the Export Param checkbox)
' 5. Saves the part
'
' Author: Quintin de Bruin © 2026
' =========================================================

Option Explicit

' Inventor API Constants
Const kAssemblyDocumentObject = 12291
Const kPartDocumentObject = 12290
Const kMillimeterLengthUnits = 11269

' Global variables
Dim m_InventorApp
Dim m_Log
Dim m_LogPath

Sub Main()
    On Error Resume Next

    ' Initialize logging
    m_Log = ""
    
    LogMessage "=== THICKNESS PARAMETER EXPORTER STARTED ==="
    LogMessage "Date/Time: " & Now

    ' Get Inventor application
    LogMessage "Attempting to get Inventor application..."

    Set m_InventorApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        LogMessage "No existing Inventor instance found"
        Err.Clear
        Set m_InventorApp = CreateObject("Inventor.Application")
        If Err.Number <> 0 Then
            LogMessage "ERROR: Failed to connect to Inventor - " & Err.Description
            MsgBox "Failed to connect to Inventor. Please make sure Inventor is running.", vbCritical, "Connection Failed"
            SaveLog
            Exit Sub
        End If
        m_InventorApp.Visible = True
    Else
        LogMessage "Connected to existing Inventor instance"
    End If

    If m_InventorApp Is Nothing Then
        LogMessage "ERROR: Inventor application object is Nothing"
        MsgBox "Inventor is not running. Please start Inventor first.", vbCritical, "Error"
        SaveLog
        Exit Sub
    End If

    LogMessage "Inventor application found successfully"

    ' Check if we have an active document
    If m_InventorApp.ActiveDocument Is Nothing Then
        LogMessage "ERROR: No active document found"
        MsgBox "No active document! Please open an assembly in Inventor.", vbCritical, "Error"
        SaveLog
        Exit Sub
    End If

    LogMessage "Active document: " & m_InventorApp.ActiveDocument.FullFileName
    LogMessage "Document type: " & m_InventorApp.ActiveDocument.DocumentType

    If m_InventorApp.ActiveDocument.DocumentType <> kAssemblyDocumentObject Then
        LogMessage "ERROR: Document is not an assembly"
        MsgBox "Please open an ASSEMBLY document (.iam file), not a part." & vbCrLf & vbCrLf & _
               "Current document: " & m_InventorApp.ActiveDocument.DisplayName, vbExclamation, "Assembly Required"
        SaveLog
        Exit Sub
    End If

    Dim asmDoc
    Set asmDoc = m_InventorApp.ActiveDocument
    LogMessage "Processing assembly: " & asmDoc.FullFileName

    ' Step 1: Scan assembly for PLATE parts
    LogMessage ""
    LogMessage "=== STEP 1: Scanning assembly for plate parts ==="
    LogMessage "Looking for parts with PL, PLATE, VRN, or S355JR in description or part number..."
    
    Dim plateParts
    Set plateParts = ScanAssemblyForPlateParts(asmDoc)

    If plateParts.Count = 0 Then
        LogMessage "No plate parts found in BOM"
        MsgBox "No parts found that contain 'PL', 'PLATE', 'VRN', or 'S355JR' in their part number or description." & vbCrLf & vbCrLf & _
               "No plate parts to process.", vbInformation, "No Plate Parts Found"
        SaveLog
        Exit Sub
    End If

    LogMessage "Found " & plateParts.Count & " unique plate part(s)"

    Dim processedCount
    processedCount = 0
    Dim skippedCount
    skippedCount = 0
    Dim failedCount
    failedCount = 0

    ' Step 2: Process each plate part
    LogMessage ""
    LogMessage "=== STEP 2: Processing plate parts ==="
    
    Dim partPath
    For Each partPath In plateParts.Keys
        LogMessage ""
        LogMessage "Processing: " & partPath

        Dim partDoc
        Set partDoc = Nothing
        On Error Resume Next
        Set partDoc = m_InventorApp.Documents.Open(partPath, False)
        On Error GoTo 0

        If partDoc Is Nothing Then
            LogMessage "  ERROR: Failed to open part"
            failedCount = failedCount + 1
        Else
            LogMessage "  Opened part successfully"

            ' Enable Thickness parameter export
            Dim result
            result = EnableThicknessParameterExport(partDoc)
            
            If result = "SUCCESS" Then
                LogMessage "  SUCCESS: Enabled Thickness parameter export"
                processedCount = processedCount + 1

                ' Save the part
                On Error Resume Next
                partDoc.Save
                If Err.Number <> 0 Then
                    LogMessage "  WARNING: Failed to save part - " & Err.Description
                    Err.Clear
                Else
                    LogMessage "  Part saved"
                End If
                On Error GoTo 0

                ' Close the part
                On Error Resume Next
                partDoc.Close True
                On Error GoTo 0
                
            ElseIf result = "ALREADY_EXPORTED" Then
                LogMessage "  SKIPPED: Thickness parameter already exported"
                skippedCount = skippedCount + 1
                On Error Resume Next
                partDoc.Close False
                On Error GoTo 0
                
            ElseIf result = "NOT_FOUND" Then
                LogMessage "  SKIPPED: No Thickness parameter found (checked ModelParameters and UserParameters)"
                skippedCount = skippedCount + 1
                On Error Resume Next
                partDoc.Close False
                On Error GoTo 0
                
            Else
                LogMessage "  ERROR: Failed to enable export - " & result
                failedCount = failedCount + 1
                On Error Resume Next
                partDoc.Close False
                On Error GoTo 0
            End If
        End If

        WScript.Sleep 200
    Next

    ' Step 3: Summary
    LogMessage ""
    LogMessage "=== PROCESSING COMPLETE ==="
    LogMessage "Parts processed successfully: " & processedCount
    LogMessage "Parts skipped (no Thickness param or already exported): " & skippedCount
    LogMessage "Parts failed: " & failedCount
    LogMessage "Total: " & (processedCount + skippedCount + failedCount)

    MsgBox "Processing complete!" & vbCrLf & vbCrLf & _
           "Successfully enabled export on: " & processedCount & " parts" & vbCrLf & _
           "Skipped: " & skippedCount & " parts" & vbCrLf & _
           "Failed: " & failedCount & " parts" & vbCrLf & vbCrLf & _
           "Check the log file for details:", vbInformation, "Conversion Summary"

    SaveLog

End Sub

Function ScanAssemblyForPlateParts(asmDoc)
    ' Returns a Dictionary of full file paths for plate parts
    ' Checks for PL, PLATE, VRN, or S355JR in part number or description
    Dim result
    Set result = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    
    ' Use BOM view to get unique parts
    Dim bom
    Set bom = asmDoc.ComponentDefinition.BOM
    bom.StructuredViewEnabled = True
    bom.StructuredViewFirstLevelOnly = False
    
    Dim bomView
    Set bomView = bom.BOMViews.Item("Structured")
    
    LogMessage "  BOM Rows: " & bomView.BOMRows.Count
    
    Dim i
    For i = 1 To bomView.BOMRows.Count
        Err.Clear
        
        Dim bomRow
        Set bomRow = bomView.BOMRows.Item(i)
        
        If Err.Number = 0 And Not bomRow Is Nothing Then
            Dim compDef
            Set compDef = Nothing
            On Error Resume Next
            Set compDef = bomRow.ComponentDefinitions.Item(1)
            On Error GoTo 0
            
            If Not compDef Is Nothing Then
                Dim partDoc
                Set partDoc = Nothing
                On Error Resume Next
                Set partDoc = compDef.Document
                On Error GoTo 0
                
                If Not partDoc Is Nothing Then
                    Dim fullPath
                    fullPath = partDoc.FullFileName
                    
                    ' Only process part files
                    If LCase(Right(fullPath, 4)) = ".ipt" Then
                        ' Get part number and description
                        Dim partNumber, description
                        partNumber = ""
                        description = ""
                        
                        On Error Resume Next
                        partNumber = partDoc.PropertySets("Design Tracking Properties").Item("Part Number").Value
                        description = partDoc.PropertySets("Design Tracking Properties").Item("Description").Value
                        On Error GoTo 0
                        
                        ' Check if it's a plate part
                        Dim isPlate
                        isPlate = IsPlatePart(partNumber, description)
                        
                        If isPlate And Not result.Exists(fullPath) Then
                            LogMessage "  Found plate: " & GetFileName(fullPath) & " (" & description & ")"
                            result.Add fullPath, partDoc
                        End If
                    End If
                End If
            End If
        End If
    Next

    Set ScanAssemblyForPlateParts = result
End Function

Function IsPlatePart(partNumber, description)
    ' Check if part is a plate based on part number or description
    Dim checkString
    checkString = UCase(partNumber & " " & description)
    
    IsPlatePart = False
    
    ' Check for common plate indicators
    If InStr(checkString, "PL ") > 0 Then IsPlatePart = True
    If InStr(checkString, "PLATE") > 0 Then IsPlatePart = True
    If InStr(checkString, "S355JR") > 0 Then IsPlatePart = True
    If InStr(checkString, "VRN") > 0 Then IsPlatePart = True
    
End Function

Function GetFileName(fullPath)
    ' Extract filename from full path
    Dim pos
    pos = InStrRev(fullPath, "\")
    If pos > 0 Then
        GetFileName = Mid(fullPath, pos + 1)
    Else
        GetFileName = fullPath
    End If
End Function

Function EnableThicknessParameterExport(partDoc)
    ' Enables export for the "Thickness" parameter
    ' Searches in BOTH ModelParameters AND UserParameters
    ' Returns: "SUCCESS", "ALREADY_EXPORTED", "NOT_FOUND", or error message
    
    On Error Resume Next

    Dim compDef
    Set compDef = partDoc.ComponentDefinition

    If compDef Is Nothing Then
        EnableThicknessParameterExport = "No component definition found"
        Exit Function
    End If

    ' Get Parameters collection
    Dim params
    Set params = compDef.Parameters
    
    If params Is Nothing Then
        EnableThicknessParameterExport = "No Parameters collection"
        Exit Function
    End If

    Dim thicknessParam
    Set thicknessParam = Nothing
    Dim paramSource
    paramSource = ""

    ' === SEARCH 1: Look in ModelParameters (Sheet Metal Thickness is usually here) ===
    LogMessage "  Searching ModelParameters for Thickness..."
    
    Dim modelParams
    Set modelParams = params.ModelParameters
    
    If Not modelParams Is Nothing Then
        Dim i
        For i = 1 To modelParams.Count
            If UCase(modelParams.Item(i).Name) = "THICKNESS" Then
                Set thicknessParam = modelParams.Item(i)
                paramSource = "ModelParameters"
                Exit For
            End If
        Next
    End If
    
    ' === SEARCH 2: If not found, look in UserParameters ===
    If thicknessParam Is Nothing Then
        LogMessage "  Searching UserParameters for Thickness..."
        
        Dim userParams
        Set userParams = params.UserParameters
        
        If Not userParams Is Nothing Then
            Dim j
            For j = 1 To userParams.Count
                If UCase(userParams.Item(j).Name) = "THICKNESS" Then
                    Set thicknessParam = userParams.Item(j)
                    paramSource = "UserParameters"
                    Exit For
                End If
            Next
        End If
    End If
    
    ' === SEARCH 3: Check if Thickness exists at all (for debugging) ===
    If thicknessParam Is Nothing Then
        LogMessage "  Thickness parameter not found in ModelParameters or UserParameters"
        
        ' List all parameters for debugging
        LogMessage "  Available ModelParameters:"
        If Not modelParams Is Nothing Then
            For i = 1 To modelParams.Count
                LogMessage "    - " & modelParams.Item(i).Name
            Next
        End If
        
        EnableThicknessParameterExport = "NOT_FOUND"
        Exit Function
    End If

    LogMessage "  Found Thickness parameter in " & paramSource
    LogMessage "  Current Value: " & thicknessParam.Value & " mm"
    
    ' === ENABLE EXPORT ===
    ' Try ExposedAsProperty first (this is the correct property for most cases)
    Dim currentExposed
    On Error Resume Next
    currentExposed = thicknessParam.ExposedAsProperty
    
    If Err.Number = 0 Then
        LogMessage "  Current ExposedAsProperty: " & currentExposed
        
        If currentExposed Then
            EnableThicknessParameterExport = "ALREADY_EXPORTED"
            Exit Function
        End If
        
        thicknessParam.ExposedAsProperty = True
        
        If Err.Number = 0 Then
            LogMessage "  ExposedAsProperty set to True"
            EnableThicknessParameterExport = "SUCCESS"
            Exit Function
        Else
            LogMessage "  ERROR setting ExposedAsProperty: " & Err.Description
            Err.Clear
        End If
    Else
        LogMessage "  ExposedAsProperty not available: " & Err.Description
        Err.Clear
    End If
    
    ' Fallback: Try ExportParameter
    On Error Resume Next
    Dim currentExport
    currentExport = thicknessParam.ExportParameter
    
    If Err.Number = 0 Then
        LogMessage "  Current ExportParameter: " & currentExport
        
        If currentExport Then
            EnableThicknessParameterExport = "ALREADY_EXPORTED"
            Exit Function
        End If
        
        thicknessParam.ExportParameter = True
        
        If Err.Number = 0 Then
            LogMessage "  ExportParameter set to True"
            EnableThicknessParameterExport = "SUCCESS"
            Exit Function
        Else
            LogMessage "  ERROR setting ExportParameter: " & Err.Description
            Err.Clear
        End If
    Else
        LogMessage "  ExportParameter not available: " & Err.Description
        Err.Clear
    End If
    
    ' Fallback: Try ExportedToSheet
    On Error Resume Next
    Dim currentSheetExport
    currentSheetExport = thicknessParam.ExportedToSheet
    
    If Err.Number = 0 Then
        LogMessage "  Current ExportedToSheet: " & currentSheetExport
        
        If currentSheetExport Then
            EnableThicknessParameterExport = "ALREADY_EXPORTED"
            Exit Function
        End If
        
        thicknessParam.ExportedToSheet = True
        
        If Err.Number = 0 Then
            LogMessage "  ExportedToSheet set to True"
            EnableThicknessParameterExport = "SUCCESS"
            Exit Function
        Else
            LogMessage "  ERROR setting ExportedToSheet: " & Err.Description
            Err.Clear
        End If
    Else
        LogMessage "  ExportedToSheet not available: " & Err.Description
        Err.Clear
    End If
    
    EnableThicknessParameterExport = "FAILED: Could not enable export using any method"
    
End Function

Sub LogMessage(msg)
    m_Log = m_Log & Now & " - " & msg & vbCrLf
    WScript.Echo msg
End Sub

Sub SaveLog()
    On Error Resume Next
    Dim fso
    Dim logFile
    Dim logFolder
    Dim wshShell

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set wshShell = CreateObject("WScript.Shell")

    ' Create log path in Documents\Inventor_Logs
    logFolder = wshShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Documents\Inventor_Logs"

    If Not fso.FolderExists(logFolder) Then
        fso.CreateFolder logFolder
    End If

    m_LogPath = logFolder & "\Thickness_Parameter_Export_" & Replace(Replace(Replace(Now, "/", "-"), ":", "-"), " ", "_") & ".log"

    Set logFile = fso.CreateTextFile(m_LogPath, True)
    logFile.WriteLine m_Log
    logFile.Close

    WScript.Echo ""
    WScript.Echo "Log saved to: " & m_LogPath

End Sub

' Start the script
Main()
