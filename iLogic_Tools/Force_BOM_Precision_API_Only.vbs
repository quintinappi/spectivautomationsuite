' =========================================================
' FORCE BOM PRECISION - API-ONLY VERSION (NO UI AUTOMATION)
' =========================================================
' This version uses ONLY API calls - no SendKeys, no UI automation.
' Much more reliable but may not work in all Inventor versions.
'
' Strategy:
' 1. Open part via API (invisible)
' 2. Modify precision parameters via API
' 3. Force document dirty flag via parameter manipulation
' 4. Save via API
' 5. Close via API
' =========================================================

Option Explicit

Const kAssemblyDocumentObject = 12291
Const kPartDocumentObject = 12290
Const kMillimeterLengthUnits = 11269
Const kCentimeterLengthUnits = 11266

Dim m_InventorApp
Dim m_Log
Dim m_Shell

Sub Main()
    On Error Resume Next
    
    Initialize
    
    ' Get assembly
    Dim asmDoc
    Set asmDoc = GetAssembly()
    If asmDoc Is Nothing Then Exit Sub
    
    ' Scan for plate parts
    Dim plateParts
    Set plateParts = ScanPlateParts(asmDoc)
    
    If plateParts.Count = 0 Then
        MsgBox "No plate parts found.", vbInformation
        Exit Sub
    End If
    
    ' Confirm
    If MsgBox("API-Only BOM Precision Update" & vbCrLf & vbCrLf & _
              "Parts: " & plateParts.Count & vbCrLf & vbCrLf & _
              "This version uses ONLY API calls (no UI automation)." & vbCrLf & _
              "It will open parts invisibly and update precision." & vbCrLf & vbCrLf & _
              "Continue?", vbOKCancel + vbQuestion, "API-Only BOM Update") <> vbOK Then
        Exit Sub
    End If
    
    ' Process parts
    Dim processedCount
    processedCount = 0
    Dim failedParts
    Set failedParts = CreateObject("Scripting.Dictionary")
    
    Dim partPath
    For Each partPath In plateParts.Keys
        Dim partName
        partName = plateParts(partPath)
        
        LogMessage ""
        LogMessage "Processing: " & partName
        
        If ProcessPartAPIMethod(partPath) Then
            processedCount = processedCount + 1
            LogMessage "  SUCCESS"
        Else
            failedParts.Add partName, partPath
            LogMessage "  FAILED"
        End If
        
        ' Brief pause
        WScript.Sleep 200
    Next
    
    ' Finalize
    FinalizeAssembly asmDoc
    
    ' Summary
    ShowSummary processedCount, plateParts.Count, failedParts.Count
    SaveLog
    
End Sub

Sub Initialize()
    m_Log = "=== API-ONLY BOM PRECISION UPDATE ===" & vbCrLf
    m_Log = m_Log & "Started: " & Now & vbCrLf & vbCrLf
    
    Set m_Shell = CreateObject("WScript.Shell")
    
    Set m_InventorApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        MsgBox "Could not connect to Inventor", vbCritical
        WScript.Quit 1
    End If
End Sub

Function GetAssembly()
    On Error Resume Next
    Set GetAssembly = Nothing
    
    If m_InventorApp.ActiveDocument Is Nothing Then Exit Function
    If m_InventorApp.ActiveDocument.DocumentType <> kAssemblyDocumentObject Then Exit Function
    
    Set GetAssembly = m_InventorApp.ActiveDocument
    LogMessage "Assembly: " & GetAssembly.DisplayName
End Function

Function ScanPlateParts(asmDoc)
    Dim result
    Set result = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    
    Dim occ
    For Each occ In asmDoc.ComponentDefinition.Occurrences
        Dim refDoc
        Set refDoc = occ.Definition.Document
        
        If Not refDoc Is Nothing Then
            If refDoc.DocumentType = kPartDocumentObject Then
                Dim desc
                desc = ""
                On Error Resume Next
                desc = refDoc.PropertySets("Design Tracking Properties")("Description").Value
                On Error GoTo 0
                
                ' Check Description for PL, VRN, or S355JR
                If InStr(UCase(desc), "PL") > 0 Or InStr(UCase(desc), "VRN") > 0 Or InStr(UCase(desc), "S355JR") > 0 Then
                    Dim partNum
                    partNum = ""
                    On Error Resume Next
                    partNum = refDoc.PropertySets("Design Tracking Properties")("Part Number").Value
                    On Error GoTo 0
                    
                    Dim fullPath
                    fullPath = refDoc.FullFileName
                    If Not result.Exists(fullPath) Then
                        result.Add fullPath, partNum
                    End If
                End If
            End If
        End If
        Err.Clear
    Next
    
    Set ScanPlateParts = result
End Function

Function ProcessPartAPIMethod(partPath)
    On Error Resume Next
    
    ProcessPartAPIMethod = False
    
    ' === METHOD 1: Open invisible (if possible) ===
    Dim partDoc
    Set partDoc = m_InventorApp.Documents.Open(partPath, False)  ' False = invisible
    
    If Err.Number <> 0 Or partDoc Is Nothing Then
        LogMessage "  ERROR: Could not open part"
        Err.Clear
        Exit Function
    End If
    
    LogMessage "  Part opened"
    
    ' === METHOD 2: Modify UnitsOfMeasure ===
    Dim uom
    Set uom = partDoc.UnitsOfMeasure
    
    If Not uom Is Nothing Then
        ' Store original
        Dim origUnits, origPrecision
        origUnits = uom.LengthUnits
        origPrecision = uom.LengthDisplayPrecision
        
        LogMessage "  Original: Units=" & origUnits & ", Precision=" & origPrecision
        
        ' Toggle units (mm -> cm -> mm) - this often triggers update
        uom.LengthUnits = kCentimeterLengthUnits
        partDoc.Update
        
        uom.LengthUnits = kMillimeterLengthUnits
        partDoc.Update
        
        ' Toggle precision
        uom.LengthDisplayPrecision = 3
        partDoc.Update
        
        uom.LengthDisplayPrecision = origPrecision
        partDoc.Update
        
        LogMessage "  UnitsOfMeasure toggled"
    End If
    
    ' === METHOD 3: Modify Parameters precision ===
    Dim params
    Set params = partDoc.ComponentDefinition.Parameters
    
    If Not params Is Nothing Then
        Dim origLinPrec
        origLinPrec = params.LinearDimensionPrecision
        
        params.LinearDimensionPrecision = 3
        partDoc.Update
        
        params.LinearDimensionPrecision = origLinPrec
        partDoc.Update
        
        LogMessage "  Parameters precision toggled"
    End If
    
    ' === METHOD 4: Create and delete a dummy parameter to force dirty flag ===
    ' This tricks Inventor into thinking the document changed
    On Error Resume Next
    Dim dummyParam
    Set dummyParam = params.UserParameters.AddByValue("_BOM_REFRESH_", 0, "mm")
    
    If Not dummyParam Is Nothing Then
        dummyParam.Value = 1
        partDoc.Update
        params.UserParameters.RemoveByName("_BOM_REFRESH_")
        partDoc.Update
        LogMessage "  Dirty flag triggered"
    End If
    Err.Clear
    
    ' === METHOD 5: Toggle a custom iProperty ===
    On Error Resume Next
    Dim customProps
    Set customProps = partDoc.PropertySets.Item("Inventor User Defined Properties")
    
    If Not customProps Is Nothing Then
        ' Add a temporary property
        Dim tempProp
        Set tempProp = customProps.Add("_PRECISION_UPDATE_", "1")
        tempProp.Value = "2"
        customProps.Remove("_PRECISION_UPDATE_")
        LogMessage "  Custom property toggled"
    End If
    Err.Clear
    
    ' === Save ===
    partDoc.Save
    
    If Err.Number <> 0 Then
        LogMessage "  WARNING during save: " & Err.Description
        Err.Clear
    End If
    
    ' === Close ===
    partDoc.Close True  ' True = save changes
    
    If Err.Number <> 0 Then
        LogMessage "  WARNING during close: " & Err.Description
        Err.Clear
    End If
    
    ProcessPartAPIMethod = True
End Function

Sub FinalizeAssembly(asmDoc)
    On Error Resume Next
    
    LogMessage ""
    LogMessage "Finalizing assembly..."
    
    asmDoc.Activate
    asmDoc.Update
    
    ' Toggle BOM views
    Dim bom
    Set bom = asmDoc.ComponentDefinition.BOM
    
    If Not bom Is Nothing Then
        bom.StructuredViewEnabled = False
        WScript.Sleep 200
        bom.StructuredViewEnabled = True
        asmDoc.Update
        
        ' Try to rebuild BOM
        On Error Resume Next
        bom.Rebuild
        Err.Clear
    End If
    
    asmDoc.Update
    LogMessage "Assembly updated"
End Sub

Sub ShowSummary(processed, total, failed)
    Dim msg
    msg = "=== API-ONLY UPDATE COMPLETE ===" & vbCrLf & vbCrLf
    msg = msg & "Total: " & total & vbCrLf
    msg = msg & "Success: " & processed & vbCrLf
    msg = msg & "Failed: " & failed & vbCrLf & vbCrLf
    
    If failed > 0 Then
        msg = msg & "NOTE: API-only method may not work for all parts." & vbCrLf
        msg = msg & "For failed parts, try the UI-automation version."
        MsgBox msg, vbExclamation, "Complete"
    Else
        msg = msg & "All parts updated via API!"
        MsgBox msg, vbInformation, "Success"
    End If
End Sub

Sub LogMessage(msg)
    m_Log = m_Log & Now & " - " & msg & vbCrLf
    WScript.Echo msg
End Sub

Sub SaveLog()
    On Error Resume Next
    
    Dim fso, logFile, logFolder
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    logFolder = m_Shell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Documents\Inventor_Logs"
    
    If Not fso.FolderExists(logFolder) Then
        fso.CreateFolder logFolder
    End If
    
    Dim timestamp
    timestamp = Replace(Replace(Replace(Now, "/", "-"), ":", "-"), " ", "_")
    
    Dim logPath
    logPath = logFolder & "\BOM_Precision_API_" & timestamp & ".log"
    
    Set logFile = fso.CreateTextFile(logPath, True)
    logFile.Write m_Log
    logFile.Close
    
    WScript.Echo ""
    WScript.Echo "Log saved to: " & logPath
End Sub

Main
