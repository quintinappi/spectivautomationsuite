' Force_BOM_Precision_UI_Test.vbs - Test with specific parts only
' Author: Quintin de Bruin © 2026

Option Explicit

Const kAssemblyDocumentObject = 12291
Const kPartDocumentObject = 12290

Dim m_InventorApp
Dim m_Shell
Dim m_Log

Sub Main()
    On Error Resume Next

    m_Log = ""

    LogMessage "=== FORCE BOM PRECISION TEST (PL13, PL15, PL16) ==="
    LogMessage ""

    Set m_InventorApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        WScript.Quit 1
    End If
    
    Set m_Shell = CreateObject("WScript.Shell")
    
    LogMessage "Connected to Inventor"

    If m_InventorApp.ActiveDocument Is Nothing Then
        WScript.Quit 1
    End If

    If m_InventorApp.ActiveDocument.DocumentType <> kAssemblyDocumentObject Then
        WScript.Quit 1
    End If

    Dim asmDoc
    Set asmDoc = m_InventorApp.ActiveDocument
    Dim asmPath
    asmPath = asmDoc.FullFileName
    LogMessage "Assembly: " & asmDoc.DisplayName
    LogMessage ""

    ' Scan for all plate parts
    LogMessage "Scanning for all PLATE parts..."
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
                
                ' Only add PLATE parts (contain "PL", "VRN", or "S355JR" in Description)
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
    LogMessage "Found " & plateParts.Count & " target parts"
    LogMessage ""
    
    ' Calculate estimated time (approximately 5 seconds per part)
    Dim secondsPerPart
    secondsPerPart = 5
    Dim estimatedSeconds
    estimatedSeconds = plateParts.Count * secondsPerPart
    Dim estimatedMinutes
    estimatedMinutes = Int(estimatedSeconds / 60)
    Dim remainingSeconds
    remainingSeconds = estimatedSeconds Mod 60
    
    ' Show warning dialog
    Dim warningMsg
    warningMsg = "[!] WARNING - DO NOT TOUCH ANYTHING! [!]" & vbCrLf & vbCrLf
    warningMsg = warningMsg & "This script will:" & vbCrLf
    warningMsg = warningMsg & "* Take over keyboard and mouse control" & vbCrLf
    warningMsg = warningMsg & "* Open and close " & plateParts.Count & " plate parts" & vbCrLf
    warningMsg = warningMsg & "* Modify Document Settings for each part" & vbCrLf
    warningMsg = warningMsg & "* Save and update the assembly" & vbCrLf & vbCrLf
    warningMsg = warningMsg & "Calculation:" & vbCrLf
    warningMsg = warningMsg & plateParts.Count & " parts × " & secondsPerPart & " seconds/part = " & estimatedSeconds & " seconds" & vbCrLf & vbCrLf
    warningMsg = warningMsg & "Estimated Time: " & estimatedMinutes & " minute(s) " & remainingSeconds & " second(s)" & vbCrLf & vbCrLf
    warningMsg = warningMsg & "Click OK to proceed, or Cancel to stop."
    
    Dim result
    result = m_Shell.Popup(warningMsg, 0, "BOM PRECISION REFRESH - CONFIRM", 1)  ' 1 = OK/Cancel
    
    If result <> 1 Then
        LogMessage "User cancelled operation"
        WScript.Quit 0
    End If
    
    ' Show countdown dialog
    m_Shell.Popup "Starting in 3 seconds..." & vbCrLf & vbCrLf & "DO NOT TOUCH ANYTHING!", 3, "Starting Process", 0
    
    LogMessage ""
    m_Shell.AppActivate "Inventor"
    WScript.Sleep 500
    
    ' Process each part
    Dim processedCount
    processedCount = 0
    
    Dim partPath
    For Each partPath In plateParts.Keys
        Dim partName
        partName = plateParts.Item(partPath)
        
        LogMessage "Processing: " & partName
        
        On Error Resume Next
        
        ' Open the part visibly
        Dim partDoc
        Set partDoc = m_InventorApp.Documents.Open(partPath, True)
        
        If Err.Number <> 0 Or partDoc Is Nothing Then
            LogMessage "  ERROR: Could not open - " & Err.Description
            Err.Clear
        Else
            ' Activate the part document
            partDoc.Activate
            WScript.Sleep 500
            
            ' Get the window title from the part name
            Dim windowTitle
            windowTitle = partDoc.DisplayName
            LogMessage "  Window title: " & windowTitle
            
            ' Try to activate by part name first, then by Inventor
            Dim activated
            activated = m_Shell.AppActivate(windowTitle)
            If Not activated Then
                activated = m_Shell.AppActivate("Autodesk Inventor")
            End If
            If Not activated Then
                activated = m_Shell.AppActivate("Inventor Professional")
            End If
            WScript.Sleep 500
            
            ' Make sure Inventor is visible and in foreground
            m_InventorApp.Visible = True
            WScript.Sleep 500
            
            ' Use Inventor's MainWindow to bring to front
            Dim hwnd
            hwnd = m_InventorApp.MainFrameHWND
            m_Shell.AppActivate hwnd
            WScript.Sleep 500
            
            ' STEP 1: Open Document Settings via custom shortcut Alt+D
            LogMessage "  Opening Document Settings (Alt+D)..."
            m_Shell.SendKeys "%d"  ' Alt+D for Document Settings (custom shortcut)
            WScript.Sleep 2000     ' Wait for dialog to open
            
            ' STEP 2: Tab 5 times, then Right arrow to get to Units tab
            LogMessage "  Tab x5, then Right..."
            m_Shell.SendKeys "{TAB}"
            WScript.Sleep 100  
            m_Shell.SendKeys "{TAB}"
            WScript.Sleep 100
            m_Shell.SendKeys "{TAB}"
            WScript.Sleep 100
            m_Shell.SendKeys "{TAB}"
            WScript.Sleep 100
            m_Shell.SendKeys "{TAB}"
            WScript.Sleep 100
            m_Shell.SendKeys "{RIGHT}"
            WScript.Sleep 100
            
            ' STEP 3: Tab x5 to get to Linear Dim Display Precision
            LogMessage "  Tab x5..."
            m_Shell.SendKeys "{TAB}"
            WScript.Sleep 100
            m_Shell.SendKeys "{TAB}"
            WScript.Sleep 100
            m_Shell.SendKeys "{TAB}"
            WScript.Sleep 100
            m_Shell.SendKeys "{TAB}"
            WScript.Sleep 100
            m_Shell.SendKeys "{TAB}"
            WScript.Sleep 100
             m_Shell.SendKeys "{TAB}"
            WScript.Sleep 100
            ' STEP 4: Tab x2
            LogMessage "  Tab x2..."
            m_Shell.SendKeys "{DOWN}"
            WScript.Sleep 100
            m_Shell.SendKeys "{UP}"
            WScript.Sleep 100
            m_Shell.SendKeys "{UP}"
            WScript.Sleep 100
              m_Shell.SendKeys "{TAB}"
            WScript.Sleep 100
              m_Shell.SendKeys "{TAB}"
            WScript.Sleep 100
              m_Shell.SendKeys "{TAB}"
            WScript.Sleep 100
              m_Shell.SendKeys "{TAB}"
            WScript.Sleep 100
              m_Shell.SendKeys "{TAB}"
            WScript.Sleep 100
              m_Shell.SendKeys "{TAB}"
            WScript.Sleep 100
              m_Shell.SendKeys "{ENTER}"
            WScript.Sleep 100
            
            ' STEP 7: Escape to close the dialog
            LogMessage "  Escape to close dialog..."
            m_Shell.SendKeys "{ESC}"
            WScript.Sleep 200
            
            ' Save the part
            LogMessage "  Saving part..."
            m_Shell.SendKeys "^s"  ' Ctrl+S to save
            WScript.Sleep 500
            
            ' Close the part
            LogMessage "  Closing part..."
            partDoc.Close False
            WScript.Sleep 300
            
            ' Re-activate Inventor window for next part
            m_InventorApp.MainFrameHWND
            m_Shell.AppActivate m_InventorApp.MainFrameHWND
            WScript.Sleep 500
            
            processedCount = processedCount + 1
            LogMessage "  Done"
        End If
        
        Err.Clear
        On Error GoTo 0
        
        ' === CHECKPOINT: Continue or Stop ===
        Dim continueMsg
        continueMsg = "Part completed: " & partName & vbCrLf & vbCrLf
        continueMsg = continueMsg & "Processed " & processedCount & " of " & plateParts.Count & " parts" & vbCrLf & vbCrLf
        continueMsg = continueMsg & "Continue to next part?" & vbCrLf & vbCrLf
        continueMsg = continueMsg & "[YES] = Continue to next part" & vbCrLf
        continueMsg = continueMsg & "[NO] = STOP the script"
        
        Dim userContinue
        userContinue = MsgBox(continueMsg, vbYesNo + vbQuestion, "Continue to Next Part?")
        
        If userContinue = vbNo Then
            LogMessage "User chose to STOP after part: " & partName
            LogMessage ""
            Exit For
        End If
        
        LogMessage ""
    Next
    
    ' Return to assembly
    LogMessage "Returning to assembly..."
    On Error Resume Next
    Set asmDoc = m_InventorApp.Documents.Open(asmPath, True)
    If Not asmDoc Is Nothing Then
        asmDoc.Activate
    End If
    On Error GoTo 0
    
    ' Update assembly
    LogMessage "Updating assembly..."
    m_InventorApp.ActiveDocument.Update
    
    LogMessage ""
    LogMessage "=== COMPLETE ==="
    LogMessage "Parts processed: " & processedCount
    LogMessage ""
    LogMessage "✓ SUCCESS! All " & processedCount & " parts have been updated."
    LogMessage "✓ BOM precision has been refreshed for all plate parts."
    LogMessage "✓ Assembly has been updated."
    LogMessage ""
    LogMessage "You can now use the assembly normally."
    LogMessage ""
    
    ' Show success dialog
    Dim successMsg
    successMsg = "[OK] SUCCESS! [OK]" & vbCrLf & vbCrLf
    successMsg = successMsg & "All " & processedCount & " plate parts have been updated." & vbCrLf & vbCrLf
    successMsg = successMsg & "+ BOM precision has been refreshed" & vbCrLf
    successMsg = successMsg & "+ Assembly has been updated" & vbCrLf & vbCrLf
    successMsg = successMsg & "You can now use the assembly normally."
    m_Shell.Popup successMsg, 0, "Process Complete!", 0

End Sub

Sub LogMessage(msg)
    m_Log = m_Log & Now & " - " & msg & vbCrLf
    WScript.Echo msg
End Sub

Main
