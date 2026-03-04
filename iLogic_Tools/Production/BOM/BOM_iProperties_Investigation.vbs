' BOM iProperties Investigation
' Check if BOM display uses iProperties instead of parameter precision
' Many BOMs display custom iProperties that may have their own formatting
' Author: Quintin de Bruin © 2026

Option Explicit

Const kAssemblyDocumentObject = 12291

Dim m_InventorApp

Sub Main()
    On Error Resume Next

    WScript.Echo "=== BOM iPROPERTIES INVESTIGATION ==="
    WScript.Echo ""

    ' Connect to Inventor
    Set m_InventorApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Inventor not running"
        Exit Sub
    End If

    If m_InventorApp.ActiveDocument Is Nothing Then
        WScript.Echo "ERROR: No active document"
        Exit Sub
    End If

    If m_InventorApp.ActiveDocument.DocumentType <> kAssemblyDocumentObject Then
        WScript.Echo "ERROR: Not an assembly"
        Exit Sub
    End If

    Dim asmDoc
    Set asmDoc = m_InventorApp.ActiveDocument
    WScript.Echo "Assembly: " & asmDoc.DisplayName
    WScript.Echo ""

    ' Get BOM
    Dim bom
    Set bom = asmDoc.ComponentDefinition.BOM
    If bom Is Nothing Then
        WScript.Echo "ERROR: Could not access BOM"
        Exit Sub
    End If

    ' Get BOM view
    If Not bom.StructuredViewEnabled Then
        bom.StructuredViewEnabled = True
    End If

    Dim bomView
    Set bomView = bom.BOMViews.Item("Structured")
    If bomView Is Nothing Then
        WScript.Echo "ERROR: Could not access Structured BOMView"
        Exit Sub
    End If

    WScript.Echo "Investigating BOM rows for iProperty usage..."
    WScript.Echo ""

    ' Investigate first few rows
    Dim rowCount
    rowCount = bomView.BOMRows.Count
    If rowCount > 5 Then rowCount = 5

    Dim i
    For i = 1 To rowCount
        Dim bomRow
        Set bomRow = bomView.BOMRows.Item(i)

        WScript.Echo "--- BOMRow " & i & " ---"

        ' Get the component
        Dim compDef
        Set compDef = bomRow.ComponentDefinitions.Item(1)

        WScript.Echo "  Component: " & compDef.Name
        WScript.Echo "  Document: " & compDef.Document.DisplayName

        ' Check iProperties
        On Error Resume Next
        Dim designProps
        Set designProps = compDef.Document.PropertySets.Item("Design Tracking Properties")

        If Err.Number = 0 Then
            WScript.Echo "  iProperties found:"

            ' Check common quantity-related properties
            Dim props
            props = Array("Part Number", "Description", "Material", "Stock Number", "Quantity")

            Dim prop
            For Each prop In props
                Dim propVal
                propVal = designProps.Item(prop).Value
                If propVal <> "" Then
                    WScript.Echo "    " & prop & ": " & propVal
                End If
            Next

            ' Check custom properties
            Dim customProps
            Set customProps = compDef.Document.PropertySets.Item("Inventor User Defined Properties")

            If Err.Number = 0 Then
                WScript.Echo "  Custom Properties:"
                Dim j
                For j = 1 To customProps.Count
                    Dim customProp
                    Set customProp = customProps.Item(j)
                    WScript.Echo "    " & customProp.Name & ": " & customProp.Value & " (" & customProp.PropId & ")"
                Next
            Else
                WScript.Echo "  No custom properties"
                Err.Clear
            End If
        Else
            WScript.Echo "  No design tracking properties accessible"
            Err.Clear
        End If

        WScript.Echo ""
    Next

    ' === CHECK BOM COLUMN MAPPINGS ===
    WScript.Echo "=== BOM COLUMN PROPERTY MAPPINGS ==="
    WScript.Echo ""

    On Error Resume Next
    Dim bomColumns
    Set bomColumns = bomView.BOMColumns

    If Err.Number = 0 Then
        WScript.Echo "BOM Columns:"

        For i = 1 To bomColumns.Count
            Dim col
            Set col = bomColumns.Item(i)
            WScript.Echo "  Column " & i & ": '" & col.Title & "' -> Property: " & col.Property & " (ID: " & col.PropertyId & ")"
        Next
    Else
        WScript.Echo "BOMColumns not accessible"
        Err.Clear
    End If

    WScript.Echo ""
    WScript.Echo "=== ANALYSIS ==="
    WScript.Echo ""
    WScript.Echo "If BOM shows decimals, check:"
    WScript.Echo "1. Are there custom iProperties with quantity values?"
    WScript.Echo "2. Do BOM columns map to iProperties instead of parameters?"
    WScript.Echo "3. Are iProperty values stored as strings with formatting?"
    WScript.Echo ""
    WScript.Echo "If BOM columns use Property IDs like 29, 30, etc. (not parameter-based),"
    WScript.Echo "then the issue is iProperty formatting, not parameter precision!"

End Sub

Main</content>
<parameter name="filePath">c:\Users\Quintin\Documents\Spectiv\3. Working\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\iLogic_Tools\BOM_iProperties_Investigation.vbs