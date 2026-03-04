' Length2_Parameter_Exporter.vbs - DETAILING WORKFLOW STEP 5c: Enable export for Length2 user parameter
' DETAILING WORKFLOW - STEP 5c: Enable export for non-plate Length2 properties
' Length2 Parameter Exporter - Standalone VBScript
' Enables export for Length2 user parameter on NON-plate parts
' Author: Quintin de Bruin © 2026

Option Explicit

' Inventor API Constants
Const kAssemblyDocumentObject = 12291
Const kPartDocumentObject = 12290

' Global variables
Dim m_InventorApp
Dim m_Log
Dim m_LogPath

Sub Main()
    On Error Resume Next

    ' Initialize logging
    m_Log = ""
    
    LogMessage "=== LENGTH2 PARAMETER EXPORTER STARTED ==="

    ' Get Inventor application
    LogMessage "Attempting to get Inventor application..."
    On Error Resume Next

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
    On Error GoTo 0

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

    LogMessage "Active document found: " & m_InventorApp.ActiveDocument.FullFileName
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

    ' Step 1: Scan assembly for NON-plate parts
    LogMessage "STEP 1: Scanning assembly for NON-plate parts (excluding 'PL' and 'S355JR')"
    Dim nonPlateParts
    Set nonPlateParts = ScanAssemblyForNonPlateParts(asmDoc)

    If nonPlateParts.Count = 0 Then
        LogMessage "No non-plate parts found in assembly"
        MsgBox "No parts found that exclude 'PL' and 'S355JR' in their part number." & vbCrLf & vbCrLf & _
               "All parts appear to be plates.", vbInformation, "No Non-Plate Parts Found"
        SaveLog
        Exit Sub
    End If

    LogMessage "Found " & nonPlateParts.Count & " non-plate parts"

    Dim processedCount
    processedCount = 0
    Dim skippedCount
    skippedCount = 0
    Dim failedCount
    failedCount = 0

    ' Step 2: Process each part
    LogMessage "STEP 2: Processing non-plate parts"
    LogMessage ""

    Dim partPath
    For Each partPath In nonPlateParts.Keys
        LogMessage "Processing part: " & partPath

        On Error Resume Next

        ' Open the part document
        Dim partDoc
        Set partDoc = m_InventorApp.Documents.Open(partPath, False)

        If Err.Number <> 0 Or partDoc Is Nothing Then
            LogMessage "ERROR: Failed to open part - " & Err.Description
            Err.Clear
            failedCount = failedCount + 1
        Else
            LogMessage "Opened part successfully"

            ' Enable Length2 parameter export
            If EnableLength2ParameterExport(partDoc) Then
                LogMessage "Successfully enabled Length2 parameter export"
                processedCount = processedCount + 1

                ' Save the part
                partDoc.Save
                LogMessage "Saved part"
            Else
                LogMessage "WARNING: Length2 parameter not found or failed to enable export"
                skippedCount = skippedCount + 1
            End If
        End If

        LogMessage ""
        On Error GoTo 0
    Next

    ' Summary
    LogMessage "=== PROCESSING COMPLETE ==="
    LogMessage "Parts processed successfully: " & processedCount
    LogMessage "Parts skipped (no Length2 param): " & skippedCount
    LogMessage "Parts failed: " & failedCount
    LogMessage "Total processed: " & (processedCount + skippedCount + failedCount)

    SaveLog

    MsgBox "Length2 Parameter Export Complete!" & vbCrLf & vbCrLf & _
           "Parts processed: " & processedCount & vbCrLf & _
           "Parts skipped (no Length2): " & skippedCount & vbCrLf & _
           "Parts failed: " & failedCount & vbCrLf & vbCrLf & _
           "Log saved to: " & m_LogPath, vbInformation, "Complete"

End Sub

' Run Main
Main

Function ScanAssemblyForNonPlateParts(asmDoc)
    ' Scans the assembly and returns a Dictionary of NON-plate part paths
    ' Non-plate = parts that do NOT contain "PL" or "S355JR" in part number
    Dim result
    Set result = CreateObject("Scripting.Dictionary")

    On Error Resume Next

    Dim oCompDef
    Set oCompDef = asmDoc.ComponentDefinition

    Dim oComp
    For Each oComp In oCompDef.Occurrences
        LogMessage "Checking component: " & oComp.Name

        ' Get the referenced document
        Dim refDoc
        Set refDoc = oComp.Definition.Document

        If Not refDoc Is Nothing Then
            ' Check if it's a part (not sub-assembly)
            If refDoc.DocumentType = kPartDocumentObject Then
                ' Get part number from iProperties
                Dim partNumber
                partNumber = ""
                On Error Resume Next
                partNumber = refDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value
                Err.Clear
                On Error GoTo 0

                ' If part number is empty, try filename
                If partNumber = "" Then
                    partNumber = oComp.Name
                End If

                LogMessage "  Part Number: " & partNumber

                ' Check if it does NOT contain "PL" or "S355JR"
                Dim fullPath
                fullPath = refDoc.FullFileName

                If InStr(UCase(partNumber), "PL") = 0 And InStr(UCase(partNumber), "S355JR") = 0 Then
                    If Not result.Exists(fullPath) Then
                        LogMessage "  -> Non-plate part identified: " & fullPath
                        result.Add fullPath, True
                    Else
                        LogMessage "  -> Non-plate part (already listed)"
                    End If
                Else
                    LogMessage "  -> Plate part (skipped)"
                End If
            End If
        End If
    Next

    Set ScanAssemblyForNonPlateParts = result
End Function

Function EnableLength2ParameterExport(partDoc)
    ' Enables export for the "Length2" user parameter
    ' Returns True if successful, False otherwise
    On Error Resume Next

    Dim compDef
    Set compDef = partDoc.ComponentDefinition

    If compDef Is Nothing Then
        LogMessage "No component definition found"
        EnableLength2ParameterExport = False
        Exit Function
    End If

    Dim userParams
    Set userParams = compDef.Parameters.UserParameters

    If userParams Is Nothing Then
        LogMessage "No user parameters found"
        EnableLength2ParameterExport = False
        Exit Function
    End If

    LogMessage "Looking for Length2 in UserParameters..."

    Dim length2Param
    Set length2Param = Nothing

    Dim i
    For i = 1 To userParams.Count
        If UCase(userParams.Item(i).Name) = "LENGTH2" Then
            Set length2Param = userParams.Item(i)
            Exit For
        End If
    Next

    If length2Param Is Nothing Then
        LogMessage "Length2 not found in UserParameters"
        EnableLength2ParameterExport = False
        Exit Function
    End If

    LogMessage "Found Length2 parameter"
    LogMessage "  Type: " & TypeName(length2Param)
    LogMessage "  Value: " & length2Param.Value

    ' Try ExportedToSheet first (User Parameters support this)
    Dim currentExport
    currentExport = length2Param.ExportedToSheet
    If Err.Number = 0 Then
        LogMessage "  ExportedToSheet: " & currentExport
        If Not currentExport Then
            LogMessage "  Setting ExportedToSheet = True..."
            length2Param.ExportedToSheet = True
            If Err.Number <> 0 Then
                LogMessage "  ERROR setting ExportedToSheet: " & Err.Description
                Err.Clear
                ' Try ExposedAsProperty as fallback
            Else
                LogMessage "  SUCCESS: ExportedToSheet enabled"
                EnableLength2ParameterExport = True
                Exit Function
            End If
        Else
            LogMessage "  Already exported"
            EnableLength2ParameterExport = True
            Exit Function
        End If
    Else
        LogMessage "  ExportedToSheet not supported: " & Err.Description
        Err.Clear
    End If

    ' Try ExposedAsProperty as alternative
    Dim currentExposed
    currentExposed = length2Param.ExposedAsProperty
    If Err.Number = 0 Then
        LogMessage "  ExposedAsProperty: " & currentExposed
        If Not currentExposed Then
            LogMessage "  Setting ExposedAsProperty = True..."
            length2Param.ExposedAsProperty = True
            If Err.Number <> 0 Then
                LogMessage "  ERROR: " & Err.Description
                Err.Clear
                EnableLength2ParameterExport = False
                Exit Function
            End If
            LogMessage "  SUCCESS: ExposedAsProperty enabled"
        Else
            LogMessage "  Already exposed (export enabled)"
        End If
        EnableLength2ParameterExport = True
        Exit Function
    Else
        LogMessage "  ExposedAsProperty not supported: " & Err.Description
        Err.Clear
    End If

    LogMessage "Could not enable Length2 export"
    EnableLength2ParameterExport = False

End Function

Sub LogMessage(msg)
    m_Log = m_Log & Now & " - " & msg & vbCrLf
    WScript.Echo msg
End Sub

Sub SaveLog()
    Dim fso
    Dim logFile
    Dim logFolder
    Dim wshShell

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set wshShell = CreateObject("WScript.Shell")

    ' Create log path in Documents\Inventor_Logs
    logFolder = wshShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Documents\Inventor_Logs"

    On Error Resume Next
    If Not fso.FolderExists(logFolder) Then
        fso.CreateFolder logFolder
    End If

    ' Generate log filename with timestamp
    Dim timestamp
    timestamp = Replace(Replace(Replace(Now, "/", "-"), ":", "-"), " ", "_")
    m_LogPath = logFolder & "\Length2_Parameter_Export_" & timestamp & ".log"

    ' Write log file
    Set logFile = fso.CreateTextFile(m_LogPath, True)
    logFile.Write m_Log
    logFile.Close

    WScript.Echo "Log saved to: " & m_LogPath
End Sub
