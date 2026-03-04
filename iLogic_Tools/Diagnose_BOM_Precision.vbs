' =========================================================
' BOM PRECISION DIAGNOSTIC - COMPREHENSIVE ANALYSIS
' =========================================================
' This script analyzes WHY BOM precision isn't updating
' It checks multiple potential causes and suggests solutions
' =========================================================

Option Explicit

Const kAssemblyDocumentObject = 12291
Const kPartDocumentObject = 12290

Dim m_InventorApp
Dim m_Log
Dim m_Shell

Sub Main()
    On Error Resume Next
    
    Initialize
    
    ' Get assembly
    Set m_InventorApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Cannot connect to Inventor"
        WScript.Quit 1
    End If
    
    If m_InventorApp.ActiveDocument Is Nothing Then
        WScript.Echo "ERROR: No active document"
        WScript.Quit 1
    End If
    
    If m_InventorApp.ActiveDocument.DocumentType <> kAssemblyDocumentObject Then
        WScript.Echo "ERROR: Not an assembly"
        WScript.Quit 1
    End If
    
    Dim asmDoc
    Set asmDoc = m_InventorApp.ActiveDocument
    
    m_Log = m_Log & "DIAGNOSING: " & asmDoc.FullFileName & vbCrLf
    m_Log = m_Log & String(60, "=") & vbCrLf & vbCrLf
    
    WScript.Echo "BOM PRECISION DIAGNOSTIC"
    WScript.Echo "======================="
    WScript.Echo ""
    WScript.Echo "Assembly: " & asmDoc.DisplayName
    WScript.Echo ""
    
    ' Run diagnostics
    CheckAssemblyPrecisionSettings asmDoc
    CheckBOMStructure asmDoc
    CheckSampleParts asmDoc
    CheckiLogicRules asmDoc
    CheckDocumentSettings asmDoc
    
    ' Summary
    ShowDiagnosisSummary
    
    ' Save log
    SaveLog
    
    WScript.Echo ""
    WScript.Echo "Diagnostic complete! Check the log file."
    WScript.Echo ""
    
End Sub

Sub Initialize()
    m_Log = "=== BOM PRECISION DIAGNOSTIC REPORT ===" & vbCrLf
    m_Log = m_Log & "Generated: " & Now & vbCrLf & vbCrLf
    Set m_Shell = CreateObject("WScript.Shell")
End Sub

Sub CheckAssemblyPrecisionSettings(asmDoc)
    m_Log = m_Log & "SECTION 1: ASSEMBLY PRECISION SETTINGS" & vbCrLf
    m_Log = m_Log & String(40, "-") & vbCrLf
    
    WScript.Echo "1. Checking assembly precision settings..."
    
    On Error Resume Next
    
    ' Check assembly-level units
    Dim uom
    Set uom = asmDoc.UnitsOfMeasure
    
    If Not uom Is Nothing Then
        m_Log = m_Log & "Assembly UnitsOfMeasure:" & vbCrLf
        m_Log = m_Log & "  LengthUnits: " & uom.LengthUnits & vbCrLf
        m_Log = m_Log & "  LengthDisplayPrecision: " & uom.LengthDisplayPrecision & vbCrLf
        m_Log = m_Log & "  AngularUnits: " & uom.AngularUnits & vbCrLf
        
        WScript.Echo "   Length Display Precision: " & uom.LengthDisplayPrecision
    End If
    
    ' Check if assembly has parameters
    Dim params
    Set params = asmDoc.ComponentDefinition.Parameters
    
    If Not params Is Nothing Then
        m_Log = m_Log & "Assembly Parameters:" & vbCrLf
        m_Log = m_Log & "  LinearDimensionPrecision: " & params.LinearDimensionPrecision & vbCrLf
        m_Log = m_Log & "  DimensionDisplayType: " & params.DimensionDisplayType & vbCrLf
        
        WScript.Echo "   Linear Dim Precision: " & params.LinearDimensionPrecision
    End If
    
    Err.Clear
    m_Log = m_Log & vbCrLf
End Sub

Sub CheckBOMStructure(asmDoc)
    m_Log = m_Log & "SECTION 2: BOM STRUCTURE" & vbCrLf
    m_Log = m_Log & String(40, "-") & vbCrLf
    
    WScript.Echo "2. Checking BOM structure..."
    
    On Error Resume Next
    
    Dim bom
    Set bom = asmDoc.ComponentDefinition.BOM
    
    If bom Is Nothing Then
        m_Log = m_Log & "ERROR: Cannot access BOM" & vbCrLf
        Exit Sub
    End If
    
    m_Log = m_Log & "BOM Properties:" & vbCrLf
    m_Log = m_Log & "  StructuredViewEnabled: " & bom.StructuredViewEnabled & vbCrLf
    m_Log = m_Log & "  StructuredViewFirstLevelOnly: " & bom.StructuredViewFirstLevelOnly & vbCrLf
    
    WScript.Echo "   Structured View: " & bom.StructuredViewEnabled
    
    ' Check BOM views
    Dim viewCount
    viewCount = bom.BOMViews.Count
    m_Log = m_Log & "  BOMViews.Count: " & viewCount & vbCrLf
    
    Dim i
    For i = 1 To viewCount
        Dim view
        Set view = bom.BOMViews.Item(i)
        m_Log = m_Log & "  View " & i & ": " & view.Name & " (Rows: " & view.BOMRows.Count & ")" & vbCrLf
    Next
    
    ' Check for plate parts in BOM
    Dim bomView
    Set bomView = bom.BOMViews.Item("Structured")
    
    Dim plateCount
    plateCount = 0
    
    For i = 1 To bomView.BOMRows.Count
        Dim row
        Set row = bomView.BOMRows.Item(i)
        
        Dim compDef
        Set compDef = row.ComponentDefinitions.Item(1)
        
        Dim partDoc
        Set partDoc = compDef.Document
        
        Dim desc
        desc = ""
        On Error Resume Next
        desc = partDoc.PropertySets("Design Tracking Properties")("Description").Value
        On Error GoTo 0
        
        Dim partNum
        partNum = ""
        On Error Resume Next
        partNum = partDoc.PropertySets("Design Tracking Properties")("Part Number").Value
        On Error GoTo 0
        
        ' Check Description for PL, VRN, or S355JR
        If InStr(UCase(desc), "PL") > 0 Or InStr(UCase(desc), "VRN") > 0 Or InStr(UCase(desc), "S355JR") > 0 Then
            plateCount = plateCount + 1
            
            ' Check precision for this part
            CheckPartPrecision partDoc, partNum, False
        End If
    Next
    
    m_Log = m_Log & "Plate parts in BOM: " & plateCount & vbCrLf
    WScript.Echo "   Plate parts found: " & plateCount
    
    Err.Clear
    m_Log = m_Log & vbCrLf
End Sub

Sub CheckPartPrecision(partDoc, partNum, detailed)
    On Error Resume Next
    
    If detailed Then
        m_Log = m_Log & "  Part: " & partNum & vbCrLf
    End If
    
    ' Check UnitsOfMeasure
    Dim uom
    Set uom = partDoc.UnitsOfMeasure
    
    If Not uom Is Nothing Then
        If detailed Then
            m_Log = m_Log & "    LengthDisplayPrecision: " & uom.LengthDisplayPrecision & vbCrLf
            m_Log = m_Log & "    LengthUnits: " & uom.LengthUnits & vbCrLf
        End If
    End If
    
    ' Check Parameters
    Dim params
    Set params = partDoc.ComponentDefinition.Parameters
    
    If Not params Is Nothing Then
        If detailed Then
            m_Log = m_Log & "    LinearDimensionPrecision: " & params.LinearDimensionPrecision & vbCrLf
        End If
    End If
    
    Err.Clear
End Sub

Sub CheckSampleParts(asmDoc)
    m_Log = m_Log & "SECTION 3: SAMPLE PLATE PARTS (Detailed)" & vbCrLf
    m_Log = m_Log & String(40, "-") & vbCrLf
    
    WScript.Echo "3. Analyzing sample plate parts..."
    
    On Error Resume Next
    
    Dim checkedCount
    checkedCount = 0
    Dim maxCheck
    maxCheck = 5  ' Only check first 5 to avoid too much output
    
    Dim occ
    For Each occ In asmDoc.ComponentDefinition.Occurrences
        If checkedCount >= maxCheck Then Exit For
        
        Dim refDoc
        Set refDoc = occ.Definition.Document
        
        If Not refDoc Is Nothing Then
            If refDoc.DocumentType = kPartDocumentObject Then
                Dim desc2
                desc2 = ""
                On Error Resume Next
                desc2 = refDoc.PropertySets("Design Tracking Properties")("Description").Value
                On Error GoTo 0
                
                Dim partNum
                partNum = ""
                On Error Resume Next
                partNum = refDoc.PropertySets("Design Tracking Properties")("Part Number").Value
                On Error GoTo 0
                
                ' Check Description for PL, VRN, or S355JR
                If InStr(UCase(desc2), "PL") > 0 Or InStr(UCase(desc2), "VRN") > 0 Or InStr(UCase(desc2), "S355JR") > 0 Then
                    CheckPartPrecision refDoc, partNum, True
                    checkedCount = checkedCount + 1
                End If
            End If
        End If
        Err.Clear
    Next
    
    WScript.Echo "   Checked " & checkedCount & " sample parts"
    m_Log = m_Log & vbCrLf
End Sub

Sub CheckiLogicRules(asmDoc)
    m_Log = m_Log & "SECTION 4: ILOGIC RULES" & vbCrLf
    m_Log = m_Log & String(40, "-") & vbCrLf
    
    WScript.Echo "4. Checking for iLogic rules..."
    
    On Error Resume Next
    
    ' Check if iLogic is available
    Dim iLogicAutomation
    Set iLogicAutomation = m_InventorApp.iLogicAutomation
    
    If iLogicAutomation Is Nothing Then
        m_Log = m_Log & "iLogic Automation not available" & vbCrLf
        Exit Sub
    End If
    
    m_Log = m_Log & "iLogic is available" & vbCrLf
    
    ' Check for rules in assembly
    Dim ruleNames
    ruleNames = iLogicAutomation.GetRuleNames(asmDoc)
    
    If IsArray(ruleNames) Then
        m_Log = m_Log & "Assembly iLogic rules: " & UBound(ruleNames) + 1 & vbCrLf
        
        Dim i
        For i = 0 To UBound(ruleNames)
            m_Log = m_Log & "  - " & ruleNames(i) & vbCrLf
        Next
    Else
        m_Log = m_Log & "No iLogic rules in assembly" & vbCrLf
    End If
    
    WScript.Echo "   iLogic rules checked"
    
    Err.Clear
    m_Log = m_Log & vbCrLf
End Sub

Sub CheckDocumentSettings(asmDoc)
    m_Log = m_Log & "SECTION 5: DOCUMENT SETTINGS" & vbCrLf
    m_Log = m_Log & String(40, "-") & vbCrLf
    
    WScript.Echo "5. Checking document settings..."
    
    On Error Resume Next
    
    ' Check for any saved settings
    Dim propSets
    Set propSets = asmDoc.PropertySets
    
    m_Log = m_Log & "Property Sets: " & propSets.Count & vbCrLf
    
    Dim i
    For i = 1 To propSets.Count
        m_Log = m_Log & "  " & i & ". " & propSets.Item(i).Name & vbCrLf
    Next
    
    ' Check Design Tracking Properties
    Dim designProps
    Set designProps = propSets.Item("Design Tracking Properties")
    
    If Not designProps Is Nothing Then
        m_Log = m_Log & "Design Tracking Properties:" & vbCrLf
        
        Dim j
        For j = 1 To designProps.Count
            On Error Resume Next
            Dim prop
            Set prop = designProps.Item(j)
            If Not prop Is Nothing Then
                m_Log = m_Log & "  " & prop.Name & ": " & prop.Value & vbCrLf
            End If
            Err.Clear
        Next
    End If
    
    WScript.Echo "   Document settings checked"
    
    Err.Clear
    m_Log = m_Log & vbCrLf
End Sub

Sub ShowDiagnosisSummary()
    m_Log = m_Log & "SECTION 6: DIAGNOSIS SUMMARY" & vbCrLf
    m_Log = m_Log & String(40, "-") & vbCrLf
    
    WScript.Echo ""
    WScript.Echo "DIAGNOSIS SUMMARY"
    WScript.Echo "================="
    WScript.Echo ""
    
    m_Log = m_Log & "LIKELY CAUSES OF BOM PRECISION ISSUES:" & vbCrLf & vbCrLf
    
    m_Log = m_Log & "1. DOCUMENT DIRTY FLAG NOT TRIGGERED" & vbCrLf
    m_Log = m_Log & "   - Inventor only updates BOM when document is marked 'dirty'" & vbCrLf
    m_Log = m_Log & "   - API changes may not always trigger this flag" & vbCrLf
    m_Log = m_Log & "   - Solution: Use UI automation OR create dummy parameter" & vbCrLf & vbCrLf
    
    m_Log = m_Log & "2. BOM VIEW NOT REFRESHED" & vbCrLf
    m_Log = m_Log & "   - BOM views cache values" & vbCrLf
    m_Log = m_Log & "   - Solution: Toggle StructuredViewEnabled or call Rebuild" & vbCrLf & vbCrLf
    
    m_Log = m_Log & "3. PRECISION SET AT WRONG LEVEL" & vbCrLf
    m_Log = m_Log & "   - Part-level vs Assembly-level precision" & vbCrLf
    m_Log = m_Log & "   - BOM uses part-level settings" & vbCrLf
    m_Log = m_Log & "   - Solution: Update each part individually" & vbCrLf & vbCrLf
    
    m_Log = m_Log & "4. DRAWING/DETAILING SETTINGS" & vbCrLf
    m_Log = m_Log & "   - BOM in drawings may have different settings" & vbCrLf
    m_Log = m_Log & "   - Solution: Check drawing template settings" & vbCrLf & vbCrLf
    
    m_Log = m_Log & "RECOMMENDED SOLUTIONS:" & vbCrLf & vbCrLf
    
    m_Log = m_Log & "Option A: MANUAL (Most Reliable)" & vbCrLf
    m_Log = m_Log & "  1. Open each part" & vbCrLf
    m_Log = m_Log & "  2. Go to Document Settings > Units" & vbCrLf
    m_Log = m_Log & "  3. Toggle any setting (e.g., precision 0->1->0)" & vbCrLf
    m_Log = m_Log & "  4. Click OK (this marks document dirty)" & vbCrLf
    m_Log = m_Log & "  5. Save" & vbCrLf
    m_Log = m_Log & "  6. Close" & vbCrLf & vbCrLf
    
    m_Log = m_Log & "Option B: SCRIPT WITH DIRTY FLAG" & vbCrLf
    m_Log = m_Log & "  - Use Force_BOM_Precision_Robust.vbs" & vbCrLf
    m_Log = m_Log & "  - It creates a dummy parameter to force dirty flag" & vbCrLf
    m_Log = m_Log & "  - More reliable than pure API toggling" & vbCrLf & vbCrLf
    
    m_Log = m_Log & "Option C: ILOGIC RULE" & vbCrLf
    m_Log = m_Log & "  - Create an iLogic rule that runs on part open" & vbCrLf
    m_Log = m_Log & "  - Rule toggles a dummy parameter" & vbCrLf
    m_Log = m_Log & "  - Ensures all parts are 'touched' before BOM" & vbCrLf & vbCrLf
    
    WScript.Echo "Check the full log file for detailed analysis."
    WScript.Echo ""
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
    logPath = logFolder & "\BOM_Precision_Full_Diagnostic_" & timestamp & ".log"
    
    Set logFile = fso.CreateTextFile(logPath, True)
    logFile.Write m_Log
    logFile.Close
    
    WScript.Echo "Full diagnostic log saved to:"
    WScript.Echo logPath
End Sub

Main
