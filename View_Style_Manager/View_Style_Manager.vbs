' View_Style_Manager.vbs
' ==============================================================================
' VIEW STYLE MANAGER FOR INVENTOR IDW FILES
' ==============================================================================
' This script:
' 1. Scans IDW files to detect what styles are applied to views
' 2. Lists all available styles in the drawing
' 3. Allows changing view styles from one to another
' 4. Handles cases where views were copied from other IDW files with different styles
' ==============================================================================

Option Explicit

Dim g_LogFileNum
Dim g_LogPath
Dim g_fso

' ==============================================================================
' MAIN ENTRY POINT
' ==============================================================================
Call Main()

Sub Main()
    Call StartLogging
    LogMessage "=== VIEW STYLE MANAGER ==="
    LogMessage "Scan and manage view styles in Inventor IDW files"
    
    ' Connect to Inventor
    Dim invApp
    On Error Resume Next
    Set invApp = GetObject(, "Inventor.Application")
    
    If Err.Number <> 0 Or invApp Is Nothing Then
        LogMessage "ERROR: No running Inventor instance found!"
        MsgBox "ERROR: Inventor is not running!" & vbCrLf & vbCrLf & _
               "Please start Inventor and open an IDW file first.", vbCritical, "View Style Manager"
        Exit Sub
    End If
    
    LogMessage "SUCCESS: Connected to Inventor"
    Err.Clear
    
    ' Get active drawing document
    Dim idwDoc
    Set idwDoc = Nothing
    
    If Not invApp.ActiveDocument Is Nothing Then
        If invApp.ActiveDocument.DocumentType = 12292 Then ' kDrawingDocumentObject
            Set idwDoc = invApp.ActiveDocument
        End If
    End If
    
    If idwDoc Is Nothing Then
        ' Search for any open drawing
        Dim doc
        For Each doc In invApp.Documents
            If doc.DocumentType = 12292 Then
                Set idwDoc = doc
                Exit For
            End If
        Next
    End If
    
    If idwDoc Is Nothing Then
        LogMessage "ERROR: No drawing document found"
        MsgBox "ERROR: No IDW/DWG file is open!" & vbCrLf & vbCrLf & _
               "Please open an Inventor drawing file first.", vbCritical, "View Style Manager"
        Exit Sub
    End If
    
    LogMessage "Working with drawing: " & idwDoc.DisplayName
    
    ' Show main menu
    Call ShowMainMenu(invApp, idwDoc)
    
    Call StopLogging
End Sub

' ==============================================================================
' MAIN MENU
' ==============================================================================
Sub ShowMainMenu(invApp, idwDoc)
    Dim choice
    
    Do
        choice = MsgBox("VIEW STYLE MANAGER" & vbCrLf & vbCrLf & _
                       "Current Drawing: " & idwDoc.DisplayName & vbCrLf & vbCrLf & _
                       "What would you like to do?" & vbCrLf & vbCrLf & _
                       "YES = Scan and display view styles" & vbCrLf & _
                       "NO = Change view styles" & vbCrLf & _
                       "CANCEL = Exit", _
                       vbYesNoCancel + vbQuestion, "View Style Manager")
        
        If choice = vbYes Then
            Call ScanViewStyles(idwDoc)
        ElseIf choice = vbNo Then
            Call ChangeViewStyles(idwDoc)
        End If
    Loop While choice <> vbCancel
    
    LogMessage "User exited View Style Manager"
End Sub

' ==============================================================================
' SCAN VIEW STYLES - Detect what styles are applied to views
' ==============================================================================
Sub ScanViewStyles(idwDoc)
    LogMessage "=== SCANNING VIEW STYLES ==="
    
    ' Get all available styles in the document
    Dim availableStyles
    Set availableStyles = CreateObject("Scripting.Dictionary")
    Call GetAvailableStyles(idwDoc, availableStyles)
    
    ' Build report
    Dim report
    report = "AVAILABLE STYLES IN DOCUMENT:" & vbCrLf
    report = report & String(50, "=") & vbCrLf
    
    If availableStyles.Count > 0 Then
        Dim styleKeys
        styleKeys = availableStyles.Keys
        Dim i
        For i = 0 To UBound(styleKeys)
            report = report & (i + 1) & ". " & styleKeys(i) & vbCrLf
        Next
    Else
        report = report & "(No styles found)" & vbCrLf
    End If
    
    report = report & vbCrLf & "VIEWS AND THEIR STYLES:" & vbCrLf
    report = report & String(50, "=") & vbCrLf
    
    ' Scan all sheets
    Dim sheetNum
    For sheetNum = 1 To idwDoc.Sheets.Count
        Dim sheet
        Set sheet = idwDoc.Sheets.Item(sheetNum)
        
        report = report & vbCrLf & "Sheet: " & sheet.Name & vbCrLf
        LogMessage "Scanning sheet: " & sheet.Name
        
        ' Scan all views on this sheet
        Dim viewNum
        For viewNum = 1 To sheet.DrawingViews.Count
            Dim view
            Set view = sheet.DrawingViews.Item(viewNum)
            
            Dim viewInfo
            viewInfo = GetViewStyleInfo(view)
            
            report = report & "  View " & viewNum & ": " & view.Name & vbCrLf
            report = report & "    Type: " & viewInfo("Type") & vbCrLf
            report = report & "    Style: " & viewInfo("Style") & vbCrLf
            
            LogMessage "  View: " & view.Name & " | Type: " & viewInfo("Type") & " | Style: " & viewInfo("Style")
        Next
    Next
    
    ' Save report to file
    Dim reportPath
    reportPath = g_fso.GetParentFolderName(idwDoc.FullFileName) & "\ViewStyleReport_" & _
                 Replace(Replace(Replace(Now, "/", "-"), ":", "-"), " ", "_") & ".txt"
    
    Dim reportFile
    Set reportFile = g_fso.CreateTextFile(reportPath, True)
    reportFile.Write report
    reportFile.Close
    
    LogMessage "Report saved to: " & reportPath
    
    ' Display report
    MsgBox report, vbInformation, "View Style Scan Results"
    
    ' Offer to open report file
    Dim openReport
    openReport = MsgBox("Report saved to:" & vbCrLf & reportPath & vbCrLf & vbCrLf & _
                        "Would you like to open the report file?", _
                        vbYesNo + vbQuestion, "View Style Report")
    
    If openReport = vbYes Then
        Dim shell
        Set shell = CreateObject("WScript.Shell")
        shell.Run """" & reportPath & """"
    End If
End Sub

' ==============================================================================
' GET VIEW STYLE INFORMATION
' ==============================================================================
Function GetViewStyleInfo(view)
    Dim info
    Set info = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    
    ' Get view type
    Dim viewType
    viewType = "Unknown"
    Select Case view.ViewType
        Case 117506049 ' kStandardDrawingViewType
            viewType = "Standard"
        Case 117506050 ' kProjectedDrawingViewType
            viewType = "Projected"
        Case 117506051 ' kAuxiliaryDrawingViewType
            viewType = "Auxiliary"
        Case 117506052 ' kSectionDrawingViewType
            viewType = "Section"
        Case 117506053 ' kDetailDrawingViewType
            viewType = "Detail"
        Case 117506054 ' kDraftingViewType
            viewType = "Drafting"
        Case 117506055 ' kOverlayDrawingViewType
            viewType = "Overlay"
        Case Else
            viewType = "Type " & CStr(view.ViewType)
    End Select
    
    info.Add "Type", viewType
    
    ' Get the ACTUAL style applied to this specific view
    ' This is critical - views copied from other IDWs retain their original style
    Dim styleName
    styleName = "(No style)"
    
    On Error Resume Next
    Err.Clear
    
    ' Try to get the style from the view
    Dim viewStyle
    Set viewStyle = Nothing
    
    ' Method 1: Try view.Style property
    Set viewStyle = view.Style
    If Err.Number = 0 And Not viewStyle Is Nothing Then
        styleName = viewStyle.Name
        LogMessage "      View style (via view.Style): " & styleName
    Else
        Err.Clear
        
        ' Method 2: Try view.StyleName property (some view types)
        On Error Resume Next
        styleName = view.StyleName
        If Err.Number = 0 And styleName <> "" Then
            LogMessage "      View style (via view.StyleName): " & styleName
        Else
            Err.Clear
            
            ' Method 3: Check if it's using a standard
            On Error Resume Next
            Dim standard
            Set standard = Nothing
            
            ' For drawing views, try to get the standard being used
            If view.ViewType = 117506049 Then ' Standard view
                ' Try to access the view's standard through its properties
                On Error Resume Next
                Set standard = view.Standard
                If Err.Number = 0 And Not standard Is Nothing Then
                    styleName = standard.Name
                    LogMessage "      View style (via view.Standard): " & styleName
                Else
                    styleName = "(Unable to read style)"
                    LogMessage "      Could not read view style - may need different API approach"
                End If
            Else
                styleName = "(Non-standard view)"
            End If
        End If
    End If
    
    Err.Clear
    info.Add "Style", styleName
    
    Set GetViewStyleInfo = info
End Function

' ==============================================================================
' GET AVAILABLE STYLES
' ==============================================================================
Sub GetAvailableStyles(idwDoc, styleDict)
    On Error Resume Next
    
    ' Get styles manager
    Dim stylesManager
    Set stylesManager = idwDoc.StylesManager
    
    If stylesManager Is Nothing Then
        LogMessage "WARNING: Could not access StylesManager"
        Exit Sub
    End If
    
    ' Try different style collections
    ' Method 1: Try ActiveStandardStyle and related standards
    Dim activeStandard
    Set activeStandard = Nothing
    On Error Resume Next
    Set activeStandard = stylesManager.ActiveStandardStyle
    
    If Not activeStandard Is Nothing Then
        LogMessage "Active Standard: " & activeStandard.Name
        If Not styleDict.Exists(activeStandard.Name) Then
            styleDict.Add activeStandard.Name, activeStandard
        End If
    End If
    Err.Clear
    
    ' Method 2: Get all standard styles
    On Error Resume Next
    Dim standardStyles
    Set standardStyles = stylesManager.StandardStyles
    
    If Not standardStyles Is Nothing Then
        LogMessage "Found " & standardStyles.Count & " standard styles"
        
        Dim i
        For i = 1 To standardStyles.Count
            Dim style
            Set style = standardStyles.Item(i)
            
            If Not style Is Nothing Then
                Dim styleName
                styleName = style.Name
                
                If Not styleDict.Exists(styleName) Then
                    styleDict.Add styleName, style
                    LogMessage "  Standard Style: " & styleName
                End If
            End If
        Next
    Else
        LogMessage "WARNING: Could not access StandardStyles collection"
    End If
    
    Err.Clear
End Sub

' ==============================================================================
' CHANGE VIEW STYLES
' ==============================================================================
Sub ChangeViewStyles(idwDoc)
    LogMessage "=== CHANGE VIEW STYLES ==="
    
    ' Get available styles
    Dim availableStyles
    Set availableStyles = CreateObject("Scripting.Dictionary")
    Call GetAvailableStyles(idwDoc, availableStyles)
    
    If availableStyles.Count = 0 Then
        MsgBox "No styles found in this drawing!", vbExclamation, "View Style Manager"
        Exit Sub
    End If
    
    ' Build style list for user selection
    Dim styleList
    styleList = "Available Styles:" & vbCrLf & vbCrLf
    
    Dim styleKeys
    styleKeys = availableStyles.Keys
    
    Dim i
    For i = 0 To UBound(styleKeys)
        styleList = styleList & (i + 1) & ". " & styleKeys(i) & vbCrLf
    Next
    
    ' Ask user which style to change FROM
    Dim fromStyleInput
    fromStyleInput = InputBox(styleList & vbCrLf & _
                             "Enter the NUMBER or NAME of the style you want to CHANGE FROM:" & vbCrLf & _
                             "(Leave blank to change ALL views)", _
                             "Select Source Style")
    
    Dim fromStyle
    Set fromStyle = Nothing
    Dim changeAll
    changeAll = False
    
    If Trim(fromStyleInput) = "" Then
        changeAll = True
        LogMessage "User selected: Change ALL views"
    Else
        ' Parse user input
        If IsNumeric(fromStyleInput) Then
            Dim styleIndex
            styleIndex = CInt(fromStyleInput)
            If styleIndex >= 1 And styleIndex <= availableStyles.Count Then
                Set fromStyle = availableStyles.Item(styleKeys(styleIndex - 1))
                LogMessage "User selected FROM style by index: " & styleKeys(styleIndex - 1)
            End If
        Else
            ' Try to match by name
            Dim styleName
            styleName = Trim(fromStyleInput)
            If availableStyles.Exists(styleName) Then
                Set fromStyle = availableStyles.Item(styleName)
                LogMessage "User selected FROM style by name: " & styleName
            End If
        End If
        
        If fromStyle Is Nothing And Not changeAll Then
            MsgBox "Invalid style selection!", vbExclamation, "View Style Manager"
            Exit Sub
        End If
    End If
    
    ' Ask user which style to change TO
    Dim toStyleInput
    toStyleInput = InputBox(styleList & vbCrLf & _
                           "Enter the NUMBER or NAME of the style you want to CHANGE TO:", _
                           "Select Target Style")
    
    If Trim(toStyleInput) = "" Then
        LogMessage "User cancelled style change"
        Exit Sub
    End If
    
    Dim toStyle
    Set toStyle = Nothing
    
    ' Parse user input
    If IsNumeric(toStyleInput) Then
        Dim toStyleIndex
        toStyleIndex = CInt(toStyleInput)
        If toStyleIndex >= 1 And toStyleIndex <= availableStyles.Count Then
            Set toStyle = availableStyles.Item(styleKeys(toStyleIndex - 1))
            LogMessage "User selected TO style by index: " & styleKeys(toStyleIndex - 1)
        End If
    Else
        ' Try to match by name
        Dim toStyleName
        toStyleName = Trim(toStyleInput)
        If availableStyles.Exists(toStyleName) Then
            Set toStyle = availableStyles.Item(toStyleName)
            LogMessage "User selected TO style by name: " & toStyleName
        End If
    End If
    
    If toStyle Is Nothing Then
        MsgBox "Invalid target style selection!", vbExclamation, "View Style Manager"
        Exit Sub
    End If
    
    ' Confirm the change
    Dim confirmMsg
    If changeAll Then
        confirmMsg = "Change ALL views to style: " & toStyle.Name & "?"
    Else
        confirmMsg = "Change views from style: " & fromStyle.Name & vbCrLf & _
                    "To style: " & toStyle.Name & "?"
    End If
    
    Dim confirm
    confirm = MsgBox(confirmMsg, vbYesNo + vbQuestion, "Confirm Style Change")
    
    If confirm = vbNo Then
        LogMessage "User cancelled style change"
        Exit Sub
    End If
    
    ' Apply the style change
    Call ApplyStyleChange(idwDoc, fromStyle, toStyle, changeAll)
End Sub

' ==============================================================================
' APPLY STYLE CHANGE
' ==============================================================================
Sub ApplyStyleChange(idwDoc, fromStyle, toStyle, changeAll)
    LogMessage "=== APPLYING STYLE CHANGE ==="
    
    If changeAll Then
        LogMessage "Mode: Change ALL views to: " & toStyle.Name
    Else
        LogMessage "Mode: Change views from '" & fromStyle.Name & "' to '" & toStyle.Name & "'"
    End If
    
    Dim changedCount
    changedCount = 0
    
    Dim errorCount
    errorCount = 0
    
    Dim skippedCount
    skippedCount = 0
    
    ' Process all sheets
    Dim sheetNum
    For sheetNum = 1 To idwDoc.Sheets.Count
        Dim sheet
        Set sheet = idwDoc.Sheets.Item(sheetNum)
        
        LogMessage "Processing sheet: " & sheet.Name
        
        ' Process all views
        Dim viewNum
        For viewNum = 1 To sheet.DrawingViews.Count
            Dim view
            Set view = sheet.DrawingViews.Item(viewNum)
            
            On Error Resume Next
            
            ' Check if we should change this view
            Dim shouldChange
            shouldChange = False
            
            Dim currentStyleName
            currentStyleName = "(unknown)"
            
            If changeAll Then
                shouldChange = True
                LogMessage "  View: " & view.Name & " - Will change (change all mode)"
            Else
                ' Check if current style matches fromStyle
                On Error Resume Next
                Err.Clear
                
                If Not view.Style Is Nothing Then
                    currentStyleName = view.Style.Name
                    If currentStyleName = fromStyle.Name Then
                        shouldChange = True
                        LogMessage "  View: " & view.Name & " - Current style matches '" & fromStyle.Name & "'"
                    Else
                        LogMessage "  View: " & view.Name & " - Current style is '" & currentStyleName & "' (skipping)"
                    End If
                Else
                    LogMessage "  View: " & view.Name & " - No style detected (skipping)"
                End If
                
                Err.Clear
            End If
            
            If shouldChange Then
                LogMessage "    Attempting to change style to: " & toStyle.Name
                
                ' Try to change the style
                Err.Clear
                view.Style = toStyle
                
                If Err.Number = 0 Then
                    LogMessage "    ✓ SUCCESS: Changed to " & toStyle.Name
                    changedCount = changedCount + 1
                Else
                    LogMessage "    ✗ ERROR: " & Err.Description & " (Code: " & Err.Number & ")"
                    errorCount = errorCount + 1
                    Err.Clear
                End If
            Else
                skippedCount = skippedCount + 1
            End If
        Next
    Next
    
    ' Save the document
    If changedCount > 0 Then
        LogMessage "Saving document with " & changedCount & " changes..."
        
        On Error Resume Next
        idwDoc.Save2
        
        If Err.Number = 0 Then
            LogMessage "Document saved successfully"
        Else
            LogMessage "ERROR saving document: " & Err.Description
            Err.Clear
        End If
    End If
    
    ' Show results
    Dim resultMsg
    resultMsg = "Style Change Complete!" & vbCrLf & vbCrLf & _
               "Views changed: " & changedCount & vbCrLf & _
               "Views skipped: " & skippedCount & vbCrLf & _
               "Errors: " & errorCount
    
    If changedCount > 0 Then
        resultMsg = resultMsg & vbCrLf & vbCrLf & "Document has been saved."
    End If
    
    MsgBox resultMsg, vbInformation, "View Style Manager"
    
    LogMessage "=== STYLE CHANGE COMPLETE ==="
    LogMessage "Changed: " & changedCount & " | Skipped: " & skippedCount & " | Errors: " & errorCount
End Sub

' ==============================================================================
' LOGGING FUNCTIONS
' ==============================================================================
Sub StartLogging()
    Set g_fso = CreateObject("Scripting.FileSystemObject")
    
    Dim scriptDir
    scriptDir = g_fso.GetParentFolderName(WScript.ScriptFullName)
    
    g_LogPath = scriptDir & "\ViewStyleManager_" & _
                Replace(Replace(Replace(Now, "/", "-"), ":", "-"), " ", "_") & ".log"
    
    g_LogFileNum = FreeFile()
    Set g_LogFileNum = g_fso.CreateTextFile(g_LogPath, True)
    
    LogMessage "View Style Manager Started"
    LogMessage "Log file: " & g_LogPath
End Sub

Sub StopLogging()
    LogMessage "View Style Manager Ended"
    If Not g_LogFileNum Is Nothing Then
        g_LogFileNum.Close
    End If
End Sub

Sub LogMessage(msg)
    Dim timestamp
    timestamp = Now
    
    If Not g_LogFileNum Is Nothing Then
        g_LogFileNum.WriteLine timestamp & " - " & msg
    End If
    
    ' Also output to console if running with cscript
    On Error Resume Next
    WScript.Echo timestamp & " - " & msg
    Err.Clear
End Sub

Function FreeFile()
    ' VBScript doesn't have FreeFile, so we just return a placeholder
    ' The actual file handle will be created by CreateTextFile
    FreeFile = 1
End Function
