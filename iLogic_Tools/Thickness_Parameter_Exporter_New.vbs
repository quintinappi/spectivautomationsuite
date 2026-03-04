' Thickness_Parameter_Exporter.vbs - DETAILING WORKFLOW STEP 5b: Enable export for plate thickness properties
' DETAILING WORKFLOW - STEP 5b: Enable export for plate thickness properties
' Thickness Parameter Exporter - Standalone VBScript
' Enables export for Thickness parameter on PLATE parts
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
    
    LogMessage "=== THICKNESS PARAMETER EXPORTER STARTED ==="

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

    ' Step 1: Scan assembly for PLATE parts
    LogMessage "STEP 1: Scanning assembly for PLATE parts (containing 'PL' or 'S355JR')"
    Dim plateParts
    Set plateParts = ScanAssemblyForPlateParts(asmDoc)

    If plateParts.Count = 0 Then
        LogMessage "No plate parts found in BOM"
        MsgBox "No parts found that contain 'PL' or 'S355JR' in their part number." & vbCrLf & vbCrLf & _
               "No plate parts to process.", vbInformation, "No Plate Parts Found"
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

    ' Step 2: Process each plate part
    LogMessage "STEP 2: Processing plate parts"
    Dim partPath
    For Each partPath In plateParts
        LogMessage ""
        LogMessage "Processing part: " & partPath

        Dim partDoc
        Set partDoc = Nothing
        On Error Resume Next
        Set partDoc = m_InventorApp.Documents.Open(partPath, False)
        On Error GoTo 0

        If partDoc Is Nothing Then
            LogMessage "ERROR: Failed to open part - " & partPath
            failedCount = failedCount + 1
        Else
            LogMessage "Opened part successfully"

            ' Enable Thickness parameter export
            If EnableThicknessParameterExport(partDoc) Then
                LogMessage "Successfully enabled Thickness parameter export"
                processedCount = processedCount + 1

                ' Save the part
                On Error Resume Next
                partDoc.Save
                If Err.Number <> 0 Then
                    LogMessage "WARNING: Failed to save part - " & Err.Description
                    Err.Clear
                Else
                    LogMessage "Saved part"
                End If
                On Error GoTo 0

                ' Close the part
                On Error Resume Next
                partDoc.Close True
                On Error GoTo 0
            Else
                LogMessage "WARNING: Thickness parameter not found or already exported"
                skippedCount = skippedCount + 1
                On Error Resume Next
                partDoc.Close False
                On Error GoTo 0
            End If
        End If

        WScript.Sleep 500
    Next

    ' Step 3: Summary
    LogMessage ""
    LogMessage "=== PROCESSING COMPLETE ==="
    LogMessage "Parts processed successfully: " & processedCount
    LogMessage "Parts skipped (no Thickness param): " & skippedCount
    LogMessage "Parts failed: " & failedCount
    LogMessage "Total processed: " & (processedCount + skippedCount)

    MsgBox "Processing complete!" & vbCrLf & vbCrLf & _
           "Successfully enabled export on: " & processedCount & " parts" & vbCrLf & _
           "Skipped (no Thickness param): " & skippedCount & " parts" & vbCrLf & _
           "Failed: " & failedCount & " parts" & vbCrLf & vbCrLf & _
           "Check the log file for details.", vbInformation, "Conversion Summary"

    SaveLog

End Sub

Function ScanAssemblyForPlateParts(asmDoc)
    ' Returns a collection of full file paths for parts that DO contain "PL" or "S355JR"
    Dim result
    Set result = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    Dim ocomp
    Dim partNumber
    Dim fullPath

    For Each ocomp In asmDoc.ComponentDefinition.Occurrences
        LogMessage "Checking component: " & ocomp.Name

        ' Get the referenced document
        Dim refDoc
        Set refDoc = ocomp.Definition.Document

        If Not refDoc Is Nothing Then
            ' Get part number from iProperties
            On Error Resume Next
            partNumber = ""
            If refDoc.PropertySets.Count > 0 Then
                ' Try to get from Summary Info
                partNumber = refDoc.PropertySets("Design Tracking Properties").Item("Part Number").Value
                If Err.Number <> 0 Then
                    Err.Clear
                    partNumber = ""
                End If
            End If
            On Error GoTo 0

            ' If part number is empty, try filename
            If partNumber = "" Then
                partNumber = ocomp.Name
            End If

            LogMessage "  Part Number: " & partNumber

            ' Check if it DOES contain "PL" or "S355JR"
            If InStr(UCase(partNumber), "PL") > 0 Or InStr(UCase(partNumber), "S355JR") > 0 Then
                fullPath = refDoc.FullFileName
                LogMessage "  -> Plate part identified: " & fullPath
                result.Add fullPath, True
            Else
                LogMessage "  -> Non-plate part (skipped)"
            End If
        End If
    Next

    Set ScanAssemblyForPlateParts = result
End Function

Function EnableThicknessParameterExport(partDoc)
    ' Enables export for the "Thickness" user parameter
    ' Returns True if successful, False otherwise
    On Error Resume Next

    Dim userParams
    Set userParams = partDoc.ComponentDefinition.Parameters.UserParameters

    If userParams Is Nothing Then
        LogMessage "No user parameters found"
        EnableLengthParameterExport = False
        Exit Function
    End If

    Dim param
    Dim found
    found = False

    ' Look for "Thickness" parameter
    For Each param In userParams
        LogMessage "Checking parameter: " & param.Name
        If UCase(param.Name) = "THICKNESS" Then
            LogMessage "Found Thickness parameter"
            If Not param.ExportedToSheet Then
                LogMessage "Enabling export for Thickness parameter"
                param.ExportedToSheet = True
                found = True
                If Err.Number <> 0 Then
                    LogMessage "ERROR: Failed to set export - " & Err.Description
                    Err.Clear
                    found = False
                End If
            Else
                LogMessage "Length parameter export already enabled"
                found = True
            End If
            Exit For
        End If
    Next

    EnableLengthParameterExport = found

End Function

Sub LogMessage(msg)
    m_Log = m_Log & Now & " - " & msg & vbCrLf
    WScript.Echo msg
End Sub

Sub SaveLog()
    Dim fso
    Dim logFile
    Dim logFolder

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Create log path in Documents
    logFolder = fso.GetSpecialFolder(4) & "\Inventor_Logs"

    If Not fso.FolderExists(logFolder) Then
        fso.CreateFolder logFolder
    End If

    m_LogPath = logFolder & "\Thickness_Parameter_Export_" & Replace(Replace(Replace(Now, "/", "-"), ":", "-"), " ", "_") & ".log"

    Set logFile = fso.CreateTextFile(m_LogPath, True)
    logFile.WriteLine m_Log
    logFile.Close

    WScript.Echo "Log saved to: " & m_LogPath

End Sub

' Start the script
Main()
