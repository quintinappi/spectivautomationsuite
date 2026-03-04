' BOM_Precision_Diagnostic.vbs - Diagnose BOM display precision issues
' Scans the BOM and each part's precision settings to identify mismatches
' Author: Quintin de Bruin © 2026

Option Explicit

' Inventor API Constants
Const kAssemblyDocumentObject = 12291
Const kPartDocumentObject = 12290

' Global variables
Dim m_InventorApp
Dim m_Log

Sub Main()
    On Error Resume Next

    m_Log = ""

    LogMessage "=== BOM PRECISION DIAGNOSTIC ==="
    LogMessage ""

    ' Get Inventor application
    Set m_InventorApp = GetObject(, "Inventor.Application")
    
    If Err.Number <> 0 Then
        LogMessage "ERROR: Could not connect to Inventor"
        WScript.Quit 1
    End If
    
    LogMessage "Connected to Inventor"

    ' Check for active document
    If m_InventorApp.ActiveDocument Is Nothing Then
        LogMessage "ERROR: No active document"
        WScript.Quit 1
    End If

    If m_InventorApp.ActiveDocument.DocumentType <> kAssemblyDocumentObject Then
        LogMessage "ERROR: Please open an assembly"
        WScript.Quit 1
    End If

    Dim asmDoc
    Set asmDoc = m_InventorApp.ActiveDocument
    LogMessage "Assembly: " & asmDoc.DisplayName
    LogMessage ""

    ' Get the BOM
    Dim bom
    Set bom = asmDoc.ComponentDefinition.BOM
    
    ' Enable structured view if needed
    If Not bom.StructuredViewEnabled Then
        bom.StructuredViewEnabled = True
    End If
    bom.StructuredViewFirstLevelOnly = False
    
    Dim bomView
    Set bomView = bom.BOMViews.Item("Structured")
    
    LogMessage "BOM Rows: " & bomView.BOMRows.Count
    LogMessage ""
    LogMessage "=============================================="
    LogMessage "SCANNING BOM FOR PRECISION ISSUES"
    LogMessage "=============================================="
    LogMessage ""

    Dim i
    For i = 1 To bomView.BOMRows.Count
        Dim bomRow
        Set bomRow = bomView.BOMRows.Item(i)
        
        ' Get component definition
        Dim compDef
        Set compDef = bomRow.ComponentDefinitions.Item(1)
        
        Dim partDoc
        Set partDoc = compDef.Document
        
        Dim desc
        desc = ""
        On Error Resume Next
        desc = partDoc.PropertySets.Item("Design Tracking Properties").Item("Description").Value
        Err.Clear
        On Error GoTo 0
        
        Dim partNum
        partNum = ""
        On Error Resume Next
        partNum = partDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value
        Err.Clear
        On Error GoTo 0
        
        ' Check if plate part (by Description: PL, VRN, or S355JR)
        Dim isPlate
        isPlate = (InStr(UCase(desc), "PL") > 0 Or InStr(UCase(desc), "VRN") > 0 Or InStr(UCase(desc), "S355JR") > 0)
        
        If isPlate Then
            LogMessage "--- PLATE: " & partNum & " (Desc: " & desc & ") ---"
        Else
            LogMessage "--- NON-PLATE: " & partNum & " (Desc: " & desc & ") ---"
        End If
        
        ' Get part's precision settings
        On Error Resume Next
        Dim partCompDef
        Set partCompDef = partDoc.ComponentDefinition
        
        Dim params
        Set params = partCompDef.Parameters
        
        If Not params Is Nothing Then
            LogMessage "  LinearDimensionPrecision: " & params.LinearDimensionPrecision
            LogMessage "  DimensionDisplayType: " & params.DimensionDisplayType
            LogMessage "  DisplayParameterAsExpression: " & params.DisplayParameterAsExpression
        Else
            LogMessage "  ERROR: Could not get Parameters"
        End If
        
        ' Get UnitsOfMeasure settings
        Dim uom
        Set uom = partDoc.UnitsOfMeasure
        If Not uom Is Nothing Then
            LogMessage "  LengthUnits: " & uom.LengthUnits
            LogMessage "  LengthDisplayPrecision: " & uom.LengthDisplayPrecision
        Else
            LogMessage "  ERROR: Could not get UnitsOfMeasure"
        End If
        Err.Clear
        On Error GoTo 0
        
        LogMessage ""
    Next

    ' Now try to force BOM refresh
    LogMessage "=============================================="
    LogMessage "ATTEMPTING BOM REFRESH METHODS"
    LogMessage "=============================================="
    LogMessage ""
    
    ' Method 1: Rebuild BOM
    LogMessage "Method 1: Rebuilding BOM..."
    On Error Resume Next
    bom.Rebuild
    If Err.Number <> 0 Then
        LogMessage "  ERROR: " & Err.Description
        Err.Clear
    Else
        LogMessage "  BOM.Rebuild executed"
    End If
    On Error GoTo 0
    
    ' Method 2: Assembly Update
    LogMessage "Method 2: Assembly Update..."
    asmDoc.Update
    LogMessage "  Assembly.Update executed"
    
    ' Method 3: Disable/Enable BOM view
    LogMessage "Method 3: Toggle BOM StructuredViewEnabled..."
    On Error Resume Next
    bom.StructuredViewEnabled = False
    bom.StructuredViewEnabled = True
    If Err.Number <> 0 Then
        LogMessage "  ERROR: " & Err.Description
        Err.Clear
    Else
        LogMessage "  BOM view toggled"
    End If
    On Error GoTo 0
    
    ' Method 4: Rebuild assembly
    LogMessage "Method 4: Rebuild assembly..."
    On Error Resume Next
    asmDoc.Rebuild
    If Err.Number <> 0 Then
        LogMessage "  ERROR: " & Err.Description
        Err.Clear
    Else
        LogMessage "  Assembly.Rebuild executed"
    End If
    On Error GoTo 0
    
    LogMessage ""
    LogMessage "=============================================="
    LogMessage "DIAGNOSTIC COMPLETE"
    LogMessage "=============================================="
    LogMessage ""
    LogMessage "Check the BOM in Inventor to see if values updated."
    LogMessage "If not, manual intervention may be required."

    SaveLog
    
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
    logPath = logFolder & "\BOM_Precision_Diagnostic_" & timestamp & ".log"

    Set logFile = fso.CreateTextFile(logPath, True)
    logFile.Write m_Log
    logFile.Close

    WScript.Echo "Log saved to: " & logPath
End Sub

Main
