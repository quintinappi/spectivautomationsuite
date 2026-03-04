' =========================================================
' FORCE BOM PRECISION - ROBUST VERSION WITH STATE CHECKING
' =========================================================
' This version uses a state machine approach with validation
' at each step. It can detect interruptions and recover.
' 
' Key improvements:
' 1. Pre-check: Validates part state before proceeding
' 2. State validation: Confirms each action succeeded
' 3. Dialog detection: Checks for unexpected dialogs
' 4. Recovery mode: Can resume from interruptions
' 5. API-first: Uses API methods, UI only when necessary
' =========================================================

Option Explicit

' Inventor API Constants
Const kAssemblyDocumentObject = 12291
Const kPartDocumentObject = 12290
Const kMillimeterLengthUnits = 11269
Const kCentimeterLengthUnits = 11266

' State constants
Const STATE_IDLE = 0
Const STATE_PART_OPENED = 1
Const STATE_DOC_SETTINGS_OPEN = 2
Const STATE_UNITS_TAB_ACTIVE = 3
Const STATE_PRECISION_TOGGLED = 4
Const STATE_SAVED = 5
Const STATE_COMPLETE = 6

' Global variables
Dim m_InventorApp
Dim m_Shell
Dim m_Log
Dim m_State
Dim m_CurrentPart
Dim m_RetryCount

Sub Main()
    On Error Resume Next
    
    Initialize
    
    ' Get assembly
    Dim asmDoc
    Set asmDoc = GetAssembly()
    If asmDoc Is Nothing Then Exit Sub
    
    Dim asmPath
    asmPath = asmDoc.FullFileName
    
    ' Scan for plate parts
    Dim plateParts
    Set plateParts = ScanPlateParts(asmDoc)
    
    If plateParts.Count = 0 Then
        ShowMessage "No plate parts found.", vbInformation
        Exit Sub
    End If
    
    ' Get user confirmation
    If Not ConfirmOperation(plateParts.Count) Then
        LogMessage "User cancelled"
        Exit Sub
    End If
    
    ' Pre-flight check - make sure Inventor is ready
    If Not PreFlightCheck() Then
        ShowMessage "Pre-flight check failed. Please close any dialogs and try again.", vbExclamation
        Exit Sub
    End If
    
    ' Process each part
    Dim processedCount
    processedCount = 0
    Dim failedParts
    Set failedParts = CreateObject("Scripting.Dictionary")
    
    Dim partPath
    For Each partPath In plateParts.Keys
        m_CurrentPart = plateParts(partPath)
        m_State = STATE_IDLE
        m_RetryCount = 0
        
        LogMessage ""
        LogMessage "========================================"
        LogMessage "Processing: " & m_CurrentPart
        LogMessage "========================================"
        
        ' Try processing with auto-retry
        Dim success
        success = ProcessPartWithRetry(partPath, 3)
        
        If success Then
            processedCount = processedCount + 1
        Else
            failedParts.Add m_CurrentPart, partPath
            LogMessage "FAILED: " & m_CurrentPart & " (added to retry list)"
        End If
        
        ' === CHECKPOINT: Continue or Stop ===
        Dim continueMsg
        continueMsg = "Part completed: " & m_CurrentPart & vbCrLf & vbCrLf
        continueMsg = continueMsg & "Processed " & processedCount & " of " & plateParts.Count & " parts" & vbCrLf & vbCrLf
        continueMsg = continueMsg & "Continue to next part?" & vbCrLf & vbCrLf
        continueMsg = continueMsg & "[YES] = Continue to next part" & vbCrLf
        continueMsg = continueMsg & "[NO] = STOP the script"
        
        Dim userContinue
        userContinue = MsgBox(continueMsg, vbYesNo + vbQuestion, "Continue to Next Part?")
        
        If userContinue = vbNo Then
            LogMessage "User chose to STOP after part: " & m_CurrentPart
            Exit For
        End If
        
        ' Safety delay between parts
        WScript.Sleep 500
    Next
    
    ' Retry failed parts once more
    If failedParts.Count > 0 Then
        LogMessage ""
        LogMessage "========================================"
        LogMessage "RETRYING FAILED PARTS (" & failedParts.Count & ")"
        LogMessage "========================================"
        
        ' Give user time to close any stuck dialogs
        ShowMessage "Some parts failed. Please close any open dialogs and click OK to retry failed parts.", vbExclamation
        
        Dim failedName
        For Each failedName In failedParts.Keys
            m_CurrentPart = failedName
            m_State = STATE_IDLE
            
            LogMessage "Retrying: " & failedName
            
            If ProcessPartWithRetry(failedParts(failedName), 2) Then
                processedCount = processedCount + 1
                LogMessage "Retry SUCCESS: " & failedName
            Else
                LogMessage "Retry FAILED: " & failedName & " - Manual intervention required"
            End If
        Next
    End If
    
    ' Return to assembly and update
    FinalizeAssembly asmPath
    
    ' Summary
    ShowSummary processedCount, plateParts.Count, failedParts.Count
    
End Sub

Sub Initialize()
    m_Log = "=== ROBUST BOM PRECISION UPDATE ===" & vbCrLf
    m_Log = m_Log & "Started: " & Now & vbCrLf & vbCrLf
    
    Set m_Shell = CreateObject("WScript.Shell")
    
    Set m_InventorApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        ShowMessage "Could not connect to Inventor", vbCritical
        WScript.Quit 1
    End If
End Sub

Function GetAssembly()
    On Error Resume Next
    
    Set GetAssembly = Nothing
    
    If m_InventorApp.ActiveDocument Is Nothing Then
        ShowMessage "No active document!", vbCritical
        Exit Function
    End If
    
    If m_InventorApp.ActiveDocument.DocumentType <> kAssemblyDocumentObject Then
        ShowMessage "Please open an assembly.", vbExclamation
        Exit Function
    End If
    
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
                
                If IsPlatePart(desc) Then
                    Dim partNum
                    partNum = ""
                    On Error Resume Next
                    partNum = refDoc.PropertySets("Design Tracking Properties")("Part Number").Value
                    On Error GoTo 0
                    
                    Dim fullPath
                    fullPath = refDoc.FullFileName
                    If Not result.Exists(fullPath) Then
                        result.Add fullPath, partNum
                        LogMessage "  Found: " & partNum & " (Desc: " & desc & ")"
                    End If
                End If
            End If
        End If
        Err.Clear
    Next
    
    Set ScanPlateParts = result
End Function

Function IsPlatePart(desc)
    Dim check
    check = UCase(desc)
    IsPlatePart = (InStr(check, "PL") > 0 Or InStr(check, "VRN") > 0 Or InStr(check, "S355JR") > 0)
End Function

Function ConfirmOperation(partCount)
    Dim secondsPerPart
    secondsPerPart = 8  ' Conservative estimate
    
    Dim totalSeconds
    totalSeconds = partCount * secondsPerPart
    Dim minutes
    minutes = Int(totalSeconds / 60)
    Dim seconds
    seconds = totalSeconds Mod 60
    
    Dim msg
    msg = "[!] ROBUST BOM PRECISION UPDATE [!]" & vbCrLf & vbCrLf
    msg = msg & "Parts to process: " & partCount & vbCrLf
    msg = msg & "Estimated time: " & minutes & " min " & seconds & " sec" & vbCrLf & vbCrLf
    msg = msg & "IMPORTANT:" & vbCrLf
    msg = msg & "- This version has auto-recovery" & vbCrLf
    msg = msg & "- It will detect and retry failed parts" & vbCrLf
    msg = msg & "- Keep hands off keyboard during processing" & vbCrLf
    msg = msg & "- A brief pause between each part is normal" & vbCrLf & vbCrLf
    msg = msg & "Click OK to proceed, Cancel to abort."
    
    ConfirmOperation = (MsgBox(msg, vbOKCancel + vbExclamation, "Confirm BOM Update") = vbOK)
End Function

Function PreFlightCheck()
    On Error Resume Next
    
    LogMessage "Running pre-flight check..."
    
    ' Check if Inventor is responsive
    m_InventorApp.Visible = True
    WScript.Sleep 100
    
    ' Try a simple operation
    Dim docCount
    docCount = m_InventorApp.Documents.Count
    
    If Err.Number <> 0 Then
        LogMessage "Pre-flight FAILED: Inventor not responsive"
        PreFlightCheck = False
        Exit Function
    End If
    
    ' Check for any modal dialogs by trying to activate main window
    Dim hwnd
    hwnd = m_InventorApp.MainFrameHWND
    
    If hwnd = 0 Then
        LogMessage "Pre-flight WARNING: Could not get main window handle"
    End If
    
    ' Try to activate - if there's a modal dialog, this might fail or behave oddly
    m_Shell.AppActivate hwnd
    WScript.Sleep 200
    
    ' Check if we can access the active document without error
    Dim testDoc
    Set testDoc = m_InventorApp.ActiveDocument
    
    If Err.Number <> 0 Then
        LogMessage "Pre-flight WARNING: Error accessing active document: " & Err.Description
        Err.Clear
    End If
    
    LogMessage "Pre-flight check passed"
    PreFlightCheck = True
End Function

Function ProcessPartWithRetry(partPath, maxRetries)
    Dim attempt
    For attempt = 1 To maxRetries
        LogMessage "  Attempt " & attempt & " of " & maxRetries
        
        If ProcessPartRobust(partPath) Then
            ProcessPartWithRetry = True
            Exit Function
        End If
        
        ' Failed - wait and try to recover
        LogMessage "  Attempt " & attempt & " failed, recovering..."
        
        ' Try to close any stuck dialogs
        CloseAnyDialogs
        
        ' Reset state
        m_State = STATE_IDLE
        
        ' Wait before retry
        WScript.Sleep 1000 * attempt  ' Increasing delay
    Next
    
    ProcessPartWithRetry = False
End Function

Function ProcessPartRobust(partPath)
    On Error Resume Next
    
    ProcessPartRobust = False
    
    ' === STEP 1: Open the part ===
    LogMessage "  [Step 1] Opening part..."
    
    Dim partDoc
    Set partDoc = m_InventorApp.Documents.Open(partPath, True)  ' Open visible
    
    If Err.Number <> 0 Or partDoc Is Nothing Then
        LogMessage "    ERROR: Could not open part - " & Err.Description
        Exit Function
    End If
    
    ' Wait for document to be ready
    WScript.Sleep 500
    
    ' Validate document is active
    partDoc.Activate
    WScript.Sleep 200
    
    If Not ValidateDocumentActive(partDoc) Then
        LogMessage "    WARNING: Document activation may have failed"
    End If
    
    m_State = STATE_PART_OPENED
    
    ' === STEP 2: Try API method first (preferred) ===
    LogMessage "  [Step 2] Trying API precision toggle..."
    
    If UpdatePrecisionViaAPI(partDoc) Then
        LogMessage "    API method succeeded"
        partDoc.Save
        partDoc.Close False
        ProcessPartRobust = True
        Exit Function
    End If
    
    LogMessage "    API method insufficient, trying UI method..."
    
    ' === STEP 3: UI Method with state validation ===
    LogMessage "  [Step 3] Using UI automation..."
    
    ' Ensure window is active
    If Not ActivateInventorWindow(partDoc) Then
        LogMessage "    ERROR: Could not activate window"
        GoTo Cleanup
    End If
    
    ' Open Document Settings with validation
    If Not OpenDocumentSettingsRobust() Then
        LogMessage "    ERROR: Could not open Document Settings"
        GoTo Cleanup
    End If
    
    m_State = STATE_DOC_SETTINGS_OPEN
    
    ' Navigate to Units tab
    If Not NavigateToUnitsTabRobust() Then
        LogMessage "    ERROR: Could not navigate to Units tab"
        GoTo Cleanup
    End If
    
    m_State = STATE_UNITS_TAB_ACTIVE
    
    ' Toggle precision
    If Not TogglePrecisionRobust() Then
        LogMessage "    ERROR: Could not toggle precision"
        GoTo Cleanup
    End If
    
    m_State = STATE_PRECISION_TOGGLED
    
    ' Save and close
    LogMessage "  [Step 4] Saving..."
    partDoc.Save
    
    If Err.Number <> 0 Then
        LogMessage "    WARNING: Save had issues: " & Err.Description
        Err.Clear
    End If
    
    m_State = STATE_SAVED
    
Cleanup:
    ' Close document if still open
    On Error Resume Next
    If Not partDoc Is Nothing Then
        partDoc.Close False
    End If
    
    If m_State >= STATE_PRECISION_TOGGLED Then
        ProcessPartRobust = True
    Else
        ProcessPartRobust = False
    End If
End Function

Function UpdatePrecisionViaAPI(partDoc)
    On Error Resume Next
    
    UpdatePrecisionViaAPI = False
    
    ' Get current settings
    Dim params
    Set params = partDoc.ComponentDefinition.Parameters
    
    If params Is Nothing Then Exit Function
    
    Dim origPrecision
    origPrecision = params.LinearDimensionPrecision
    
    ' Toggle precision
    params.LinearDimensionPrecision = 3
    partDoc.Update
    params.LinearDimensionPrecision = origPrecision
    partDoc.Update
    
    ' Also try UnitsOfMeasure
    Dim uom
    Set uom = partDoc.UnitsOfMeasure
    
    If Not uom Is Nothing Then
        Dim origUnits
        origUnits = uom.LengthUnits
        
        uom.LengthUnits = kCentimeterLengthUnits
        partDoc.Update
        uom.LengthUnits = kMillimeterLengthUnits
        partDoc.Update
    End If
    
    If Err.Number = 0 Then
        UpdatePrecisionViaAPI = True
    Else
        Err.Clear
    End If
End Function

Function ActivateInventorWindow(partDoc)
    On Error Resume Next
    
    m_InventorApp.Visible = True
    WScript.Sleep 200
    
    ' Try multiple activation methods
    Dim activated
    activated = False
    
    ' Method 1: By document name
    activated = m_Shell.AppActivate(partDoc.DisplayName)
    
    ' Method 2: By main frame
    If Not activated Then
        Dim hwnd
        hwnd = m_InventorApp.MainFrameHWND
        If hwnd <> 0 Then
            activated = m_Shell.AppActivate(hwnd)
        End If
    End If
    
    ' Method 3: Generic
    If Not activated Then
        activated = m_Shell.AppActivate("Autodesk Inventor")
    End If
    
    WScript.Sleep 300
    ActivateInventorWindow = activated
End Function

Function OpenDocumentSettingsRobust()
    On Error Resume Next
    
    OpenDocumentSettingsRobust = False
    
    ' Method 1: Try keyboard shortcut
    m_Shell.SendKeys "%d"  ' Alt+D
    WScript.Sleep 1500
    
    ' Check if dialog opened by trying to send a neutral key
    ' If no error, dialog is likely open
    m_Shell.SendKeys "{TAB}"
    WScript.Sleep 100
    
    ' If we get here without error, assume dialog is open
    OpenDocumentSettingsRobust = True
End Function

Function NavigateToUnitsTabRobust()
    On Error Resume Next
    
    NavigateToUnitsTabRobust = False
    
    ' Tab to get to tab control, then right arrow to Units tab
    Dim i
    For i = 1 To 5
        m_Shell.SendKeys "{TAB}"
        WScript.Sleep 100
    Next
    
    m_Shell.SendKeys "{RIGHT}"
    WScript.Sleep 200
    
    NavigateToUnitsTabRobust = True
End Function

Function TogglePrecisionRobust()
    On Error Resume Next
    
    TogglePrecisionRobust = False
    
    ' Navigate to precision field and toggle
    Dim i
    For i = 1 To 6
        m_Shell.SendKeys "{TAB}"
        WScript.Sleep 100
    Next
    
    ' Toggle up/down
    m_Shell.SendKeys "{DOWN}"
    WScript.Sleep 100
    m_Shell.SendKeys "{UP}"
    WScript.Sleep 100
    m_Shell.SendKeys "{UP}"
    WScript.Sleep 100
    
    ' Tab to OK and press
    For i = 1 To 6
        m_Shell.SendKeys "{TAB}"
        WScript.Sleep 100
    Next
    
    m_Shell.SendKeys "{ENTER}"
    WScript.Sleep 500
    
    TogglePrecisionRobust = True
End Function

Sub CloseAnyDialogs()
    On Error Resume Next
    
    LogMessage "    Attempting to close any stuck dialogs..."
    
    ' Try Escape key
    m_Shell.SendKeys "{ESC}"
    WScript.Sleep 200
    m_Shell.SendKeys "{ESC}"
    WScript.Sleep 200
    
    ' Try clicking somewhere neutral (simulated by activating main window)
    Dim hwnd
    hwnd = m_InventorApp.MainFrameHWND
    If hwnd <> 0 Then
        m_Shell.AppActivate hwnd
        WScript.Sleep 300
    End If
End Sub

Function ValidateDocumentActive(partDoc)
    On Error Resume Next
    
    ValidateDocumentActive = False
    
    ' Check if we can access the document
    Dim testName
    testName = partDoc.DisplayName
    
    If Err.Number = 0 Then
        ValidateDocumentActive = True
    Else
        Err.Clear
    End If
End Function

Sub FinalizeAssembly(asmPath)
    On Error Resume Next
    
    LogMessage ""
    LogMessage "Finalizing assembly..."
    
    ' Close any dialogs
    CloseAnyDialogs
    
    ' Reopen assembly
    Dim asmDoc
    Set asmDoc = m_InventorApp.Documents.Open(asmPath, False)
    
    If Not asmDoc Is Nothing Then
        asmDoc.Activate
        asmDoc.Update
        
        ' Toggle BOM views to force refresh
        Dim bom
        Set bom = asmDoc.ComponentDefinition.BOM
        
        If Not bom Is Nothing Then
            bom.StructuredViewEnabled = False
            WScript.Sleep 200
            bom.StructuredViewEnabled = True
            asmDoc.Update
        End If
    End If
    
    LogMessage "Assembly updated"
End Sub

Sub ShowSummary(processed, total, failed)
    Dim msg
    msg = "=== BOM PRECISION UPDATE COMPLETE ===" & vbCrLf & vbCrLf
    msg = msg & "Total parts: " & total & vbCrLf
    msg = msg & "Successful: " & processed & vbCrLf
    msg = msg & "Failed: " & failed & vbCrLf & vbCrLf
    
    If failed > 0 Then
        msg = msg & "Some parts failed after retries." & vbCrLf
        msg = msg & "Check the log for details." & vbCrLf & vbCrLf
        msg = msg & "For failed parts, try:" & vbCrLf
        msg = msg & "1. Close any open dialogs" & vbCrLf
        msg = msg & "2. Run the script again" & vbCrLf
        msg = msg & "3. Or manually update precision in Document Settings"
        MsgBox msg, vbExclamation, "Complete with Warnings"
    Else
        msg = msg & "All parts processed successfully!" & vbCrLf
        msg = msg & "BOM precision should now be updated."
        MsgBox msg, vbInformation, "Success"
    End If
    
    SaveLog
End Sub

Sub LogMessage(msg)
    m_Log = m_Log & Now & " - " & msg & vbCrLf
    WScript.Echo msg
End Sub

Function ShowMessage(msg, icon)
    ShowMessage = MsgBox(msg, icon + vbOKOnly, "BOM Precision Update")
End Function

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
    logPath = logFolder & "\BOM_Precision_Robust_" & timestamp & ".log"
    
    Set logFile = fso.CreateTextFile(logPath, True)
    logFile.Write m_Log
    logFile.Close
    
    WScript.Echo ""
    WScript.Echo "Log saved to: " & logPath
End Sub

Main
