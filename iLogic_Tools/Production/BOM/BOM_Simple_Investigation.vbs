' BOM Simple Investigation
' Basic investigation into BOM display formatting
' Author: Quintin de Bruin © 2026

Option Explicit

Const kAssemblyDocumentObject = 12291

Dim m_InventorApp
Dim m_Report

Sub Main()
    On Error Resume Next

    m_Report = ""

    LogLine "=== BOM SIMPLE INVESTIGATION ==="
    LogLine "Date: " & Date & " " & Time
    LogLine ""

    ' Connect to Inventor
    Set m_InventorApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        LogLine "ERROR: Inventor not running"
        ShowReport
        Exit Sub
    End If

    If m_InventorApp.ActiveDocument Is Nothing Then
        LogLine "ERROR: No active document"
        ShowReport
        Exit Sub
    End If

    If m_InventorApp.ActiveDocument.DocumentType <> kAssemblyDocumentObject Then
        LogLine "ERROR: Not an assembly"
        ShowReport
        Exit Sub
    End If

    Dim asmDoc
    Set asmDoc = m_InventorApp.ActiveDocument
    LogLine "Assembly: " & asmDoc.DisplayName
    LogLine ""

    ' Get BOM
    Dim bom
    Set bom = asmDoc.ComponentDefinition.BOM
    If bom Is Nothing Then
        LogLine "ERROR: Could not access BOM"
        ShowReport
        Exit Sub
    End If

    ' Get BOM view
    If Not bom.StructuredViewEnabled Then
        bom.StructuredViewEnabled = True
    End If

    Dim bomView
    Set bomView = bom.BOMViews.Item("Structured")
    If bomView Is Nothing Then
        LogLine "ERROR: Could not access Structured BOMView"
        ShowReport
        Exit Sub
    End If

    LogLine "BOMView.BOMRows.Count: " & bomView.BOMRows.Count
    LogLine ""

    ' Check first row
    If bomView.BOMRows.Count > 0 Then
        Dim bomRow
        Set bomRow = bomView.BOMRows.Item(1)

        LogLine "First BOM Row:"
        LogLine "  Component: " & bomRow.ComponentDefinitions.Item(1).Definition.Name
        LogLine "  ItemQuantity: " & bomRow.ItemQuantity

        ' Check BOMQuantity
        On Error Resume Next
        Dim bomQty
        Set bomQty = bomRow.BOMQuantity

        If Err.Number = 0 Then
            LogLine "  BOMQuantity exists"

            Dim precision
            precision = bomQty.Precision
            If Err.Number = 0 Then
                LogLine "  BOMQuantity.Precision: " & precision
            Else
                LogLine "  BOMQuantity.Precision: not found"
                Err.Clear
            End If
        Else
            LogLine "  BOMQuantity: not accessible"
            Err.Clear
        End If
    End If

    LogLine ""
    LogLine "KEY QUESTION: Does BOM show decimals?"
    LogLine "If YES: LinearDimensionPrecision is not controlling BOM"
    LogLine "If NO: Scripts work but cache invalidation failed"

    ShowReport

End Sub

Sub LogLine(msg)
    m_Report = m_Report & msg & vbCrLf
    WScript.Echo msg
End Sub

Sub ShowReport()
    Dim fso, file, reportPath
    Set fso = CreateObject("Scripting.FileSystemObject")

    reportPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\BOM_SIMPLE_REPORT.txt"

    Set file = fso.CreateTextFile(reportPath, True)
    file.Write m_Report
    file.Close

    WScript.Echo ""
    WScript.Echo "Report saved to: " & reportPath

    MsgBox "Simple investigation complete! Check BOM_SIMPLE_REPORT.txt", vbInformation, "BOM Investigation"
End Sub

Main</content>
<parameter name="filePath">c:\Users\Quintin\Documents\Spectiv\3. Working\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\iLogic_Tools\BOM_Simple_Investigation.vbs