' Force_BOM_Precision_Refresh.vbs - Forces BOM to refresh parameter display precision
' Opens each plate part, toggles precision, triggers update events, saves
' Author: Quintin de Bruin © 2026

Option Explicit

' Inventor API Constants
Const kAssemblyDocumentObject = 12291
Const kPartDocumentObject = 12290
Const kMillimeterLengthUnits = 11269
Const kCentimeterLengthUnits = 11266

' Global variables
Dim m_InventorApp
Dim m_Log

Sub Main()
    On Error Resume Next

    m_Log = ""

    LogMessage "=== FORCE BOM PRECISION REFRESH ==="
    LogMessage ""

    ' Get Inventor application
    Set m_InventorApp = GetObject(, "Inventor.Application")
    
    If Err.Number <> 0 Then
        LogMessage "ERROR: Could not connect to Inventor"
        MsgBox "Could not connect to Inventor", vbCritical
        WScript.Quit 1
    End If
    
    LogMessage "Connected to Inventor"

    ' Check for active document
    If m_InventorApp.ActiveDocument Is Nothing Then
        LogMessage "ERROR: No active document"
        MsgBox "No active document! Please open an assembly.", vbCritical
        WScript.Quit 1
    End If

    If m_InventorApp.ActiveDocument.DocumentType <> kAssemblyDocumentObject Then
        LogMessage "ERROR: Please open an assembly"
        MsgBox "Please open an assembly document.", vbExclamation
        WScript.Quit 1
    End If

    Dim asmDoc
    Set asmDoc = m_InventorApp.ActiveDocument
    LogMessage "Assembly: " & asmDoc.DisplayName
    LogMessage ""

    ' Scan for plate parts
    LogMessage "Scanning for plate parts..."
    Dim plateParts
    Set plateParts = CreateObject("Scripting.Dictionary")
    
    Dim occ
    For Each occ In asmDoc.ComponentDefinition.Occurrences
        On Error Resume Next
        Dim refDoc
        Set refDoc = occ.Definition.Document
        
        If Not refDoc Is Nothing Then
            If refDoc.DocumentType = kPartDocumentObject Then
                Dim desc
                desc = ""
                desc = refDoc.PropertySets.Item("Design Tracking Properties").Item("Description").Value
                Err.Clear
                
                ' Check Description for PL, VRN, or S355JR
                If InStr(UCase(desc), "PL") > 0 Or InStr(UCase(desc), "VRN") > 0 Or InStr(UCase(desc), "S355JR") > 0 Then
                    Dim partNum
                    partNum = ""
                    partNum = refDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value
                    Err.Clear
                    
                    Dim fullPath
                    fullPath = refDoc.FullFileName
                    If Not plateParts.Exists(fullPath) Then
                        plateParts.Add fullPath, partNum
                        LogMessage "  Found: " & partNum & " (Desc: " & desc & ")"
                    End If
                End If
            End If
        End If
        Err.Clear
        On Error GoTo 0
    Next
    
    LogMessage ""
    LogMessage "Found " & plateParts.Count & " plate parts"
    LogMessage ""
    
    If plateParts.Count = 0 Then
        MsgBox "No plate parts found.", vbInformation
        WScript.Quit 0
    End If
    
    ' Process each plate part with aggressive refresh
    LogMessage "Processing each part with precision refresh..."
    LogMessage ""
    
    Dim processedCount
    processedCount = 0
    
    Dim partPath
    For Each partPath In plateParts.Keys
        Dim partName
        partName = plateParts.Item(partPath)
        
        LogMessage "Processing: " & partName
        
        On Error Resume Next
        
        ' Open the part visibly (important for UI update)
        Dim partDoc
        Set partDoc = m_InventorApp.Documents.Open(partPath, True)
        
        If Err.Number <> 0 Or partDoc Is Nothing Then
            LogMessage "  ERROR: Could not open - " & Err.Description
            Err.Clear
        Else
            ' Make it the active document briefly
            partDoc.Activate
            
            ' Get parameters
            Dim params
            Set params = partDoc.ComponentDefinition.Parameters
            
            If Not params Is Nothing Then
                ' Read current precision
                Dim currentPrecision
                currentPrecision = params.LinearDimensionPrecision
                LogMessage "  Current LinearDimensionPrecision: " & currentPrecision
                
                ' CRITICAL: Toggle precision to force dirty flag
                ' Change to 3, update, then back to 0, update
                LogMessage "  Toggling precision: 0 -> 3 -> 0"
                params.LinearDimensionPrecision = 3
                partDoc.Update
                
                params.LinearDimensionPrecision = 0
                partDoc.Update
                
                ' Also toggle UnitsOfMeasure LengthDisplayPrecision
                Dim uom
                Set uom = partDoc.UnitsOfMeasure
                If Not uom Is Nothing Then
                    LogMessage "  Toggling LengthDisplayPrecision: 0 -> 3 -> 0"
                    uom.LengthDisplayPrecision = 3
                    partDoc.Update
                    
                    uom.LengthDisplayPrecision = 0
                    partDoc.Update
                    
                    ' Also toggle length units
                    LogMessage "  Toggling LengthUnits: mm -> cm -> mm"
                    Dim origUnits
                    origUnits = uom.LengthUnits
                    uom.LengthUnits = kCentimeterLengthUnits
                    partDoc.Update
                    
                    uom.LengthUnits = kMillimeterLengthUnits
                    partDoc.Update
                End If
                
                If Err.Number <> 0 Then
                    LogMessage "  WARNING: " & Err.Description
                    Err.Clear
                End If
                
                ' Save the part (silent)
                partDoc.Save
                LogMessage "  Saved"
                processedCount = processedCount + 1
            End If
        End If
        
        Err.Clear
        On Error GoTo 0
        
        ' === CHECKPOINT: Continue or Stop ===
        ' Show dialog asking user if they want to continue to next part
        Dim continueMsg
        continueMsg = "Part completed: " & partName & vbCrLf & vbCrLf
        continueMsg = continueMsg & "Processed " & processedCount & " of " & plateParts.Count & " parts" & vbCrLf & vbCrLf
        continueMsg = continueMsg & "Continue to next part?" & vbCrLf & vbCrLf
        continueMsg = continueMsg & "[YES] = Continue to next part" & vbCrLf
        continueMsg = continueMsg & "[NO] = STOP the script"
        
        Dim userResponse
        userResponse = MsgBox(continueMsg, vbYesNo + vbQuestion, "Continue to Next Part?")
        
        If userResponse = vbNo Then
            LogMessage "User chose to STOP after part: " & partName
            LogMessage ""
            Exit For
        End If
        
        LogMessage ""
    Next
    
    ' Now force assembly update
    LogMessage "Updating assembly..."
    asmDoc.Activate
    asmDoc.Update
    
    ' Try to access and refresh BOM
    LogMessage "Refreshing BOM..."
    On Error Resume Next
    Dim bom
    Set bom = asmDoc.ComponentDefinition.BOM
    
    ' Toggle BOM views
    bom.StructuredViewEnabled = False
    asmDoc.Update
    bom.StructuredViewEnabled = True
    asmDoc.Update
    
    ' Try BOMViews refresh
    Dim bomView
    Set bomView = bom.BOMViews.Item("Structured")
    If Not bomView Is Nothing Then
        LogMessage "BOM view accessed"
    End If
    Err.Clear
    On Error GoTo 0
    
    ' Final assembly update
    asmDoc.Update
    
    LogMessage ""
    LogMessage "=== COMPLETE ==="
    LogMessage "Parts processed: " & processedCount & " of " & plateParts.Count
    LogMessage ""
    LogMessage "Please check the BOM in Inventor."
    LogMessage "If values still show decimals, try:"
    LogMessage "  1. Right-click BOM > Refresh"
    LogMessage "  2. Close and reopen the assembly"
    
    SaveLog
    
    MsgBox "Parts processed: " & processedCount & " of " & plateParts.Count & vbCrLf & vbCrLf & _
           "Check BOM for updated precision." & vbCrLf & _
           "If still showing decimals, try refreshing BOM or reopening assembly.", _
           vbInformation, "Complete"
    
End Sub

Sub LogMessage(msg)
    m_Log = m_Log & Now & " - " & msg & vbCrLf
    WScript.Echo msg
End Sub

Sub SaveLog()
    Dim fso, logFile, logFolder, wshShell

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set wshShell = CreateObject("WScript.Shell")

    logFolder = wshShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Documents\Inventor_Logs"

    On Error Resume Next
    If Not fso.FolderExists(logFolder) Then
        fso.CreateFolder logFolder
    End If

    Dim timestamp
    timestamp = Replace(Replace(Replace(Now, "/", "-"), ":", "-"), " ", "_")
    Dim logPath
    logPath = logFolder & "\Force_BOM_Precision_Refresh_" & timestamp & ".log"

    Set logFile = fso.CreateTextFile(logPath, True)
    logFile.Write m_Log
    logFile.Close

    WScript.Echo "Log saved to: " & logPath
End Sub

Main
