' BOM Direct Manipulation - Alternative Approach
' Instead of trying to trigger cache invalidation, directly manipulate BOM display
' Try setting BOMQuantity.Precision or other display properties directly
' Author: Quintin de Bruin © 2026

Option Explicit

Const kAssemblyDocumentObject = 12291

Dim m_InventorApp

Sub Main()
    On Error Resume Next

    WScript.Echo "=== BOM DIRECT MANIPULATION TEST ==="
    WScript.Echo ""

    ' Connect to Inventor
    Set m_InventorApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Inventor not running"
        Exit Sub
    End If

    If m_InventorApp.ActiveDocument Is Nothing Then
        WScript.Echo "ERROR: No active document"
        ShowReport
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

    ' Enable structured view
    If Not bom.StructuredViewEnabled Then
        bom.StructuredViewEnabled = True
        WScript.Echo "Enabled structured BOM view"
    End If

    ' Get BOM view
    Dim bomView
    Set bomView = bom.BOMViews.Item("Structured")
    If bomView Is Nothing Then
        WScript.Echo "ERROR: Could not access Structured BOMView"
        Exit Sub
    End If

    WScript.Echo "BOMView.BOMRows.Count: " & bomView.BOMRows.Count
    WScript.Echo ""

    ' === APPROACH 1: Try to set BOMQuantity.Precision directly ===
    WScript.Echo "APPROACH 1: Setting BOMQuantity.Precision = 0 directly..."

    Dim successCount
    successCount = 0

    Dim i
    For i = 1 To bomView.BOMRows.Count
        Dim bomRow
        Set bomRow = bomView.BOMRows.Item(i)

        On Error Resume Next
        Dim bomQty
        Set bomQty = bomRow.BOMQuantity

        If Err.Number = 0 Then
            ' Try to set precision directly
            bomQty.Precision = 0
            If Err.Number = 0 Then
                WScript.Echo "  Row " & i & ": Set BOMQuantity.Precision = 0"
                successCount = successCount + 1
            Else
                WScript.Echo "  Row " & i & ": Could not set Precision - " & Err.Description
                Err.Clear
            End If
        Else
            WScript.Echo "  Row " & i & ": BOMQuantity not accessible"
            Err.Clear
        End If
    Next

    WScript.Echo "Approach 1 result: " & successCount & " rows updated"
    WScript.Echo ""

    ' === APPROACH 2: Try BOMQuantity.DisplayFormat ===
    WScript.Echo "APPROACH 2: Setting BOMQuantity.DisplayFormat..."

    successCount = 0

    For i = 1 To bomView.BOMRows.Count
        Set bomRow = bomView.BOMRows.Item(i)

        On Error Resume Next
        Set bomQty = bomRow.BOMQuantity

        If Err.Number = 0 Then
            ' Try to set display format
            bomQty.DisplayFormat = "0"  ' No decimals
            If Err.Number = 0 Then
                WScript.Echo "  Row " & i & ": Set BOMQuantity.DisplayFormat = '0'"
                successCount = successCount + 1
            Else
                WScript.Echo "  Row " & i & ": Could not set DisplayFormat - " & Err.Description
                Err.Clear
            End If
        End If
        Err.Clear
    Next

    WScript.Echo "Approach 2 result: " & successCount & " rows updated"
    WScript.Echo ""

    ' === APPROACH 3: Try BOMColumn formatting ===
    WScript.Echo "APPROACH 3: Setting BOMColumn formatting..."

    On Error Resume Next
    Dim bomColumns
    Set bomColumns = bomView.BOMColumns

    If Err.Number = 0 Then
        WScript.Echo "Found " & bomColumns.Count & " BOM columns"

        For i = 1 To bomColumns.Count
            Dim col
            Set col = bomColumns.Item(i)

            ' Look for quantity-related columns
            If InStr(LCase(col.Title), "qty") > 0 Or InStr(LCase(col.Title), "quantity") > 0 Then
                WScript.Echo "  Column '" & col.Title & "' - trying to set precision..."

                ' Try various precision properties
                col.Precision = 0
                If Err.Number = 0 Then
                    WScript.Echo "    Set Precision = 0"
                Else
                    Err.Clear
                End If

                col.DecimalPlaces = 0
                If Err.Number = 0 Then
                    WScript.Echo "    Set DecimalPlaces = 0"
                Else
                    Err.Clear
                End If

                col.Format = "0"
                If Err.Number = 0 Then
                    WScript.Echo "    Set Format = '0'"
                Else
                    Err.Clear
                End If
            End If
        Next
    Else
        WScript.Echo "BOMColumns not accessible"
        Err.Clear
    End If

    WScript.Echo ""

    ' === APPROACH 4: Force BOM refresh ===
    WScript.Echo "APPROACH 4: Forcing BOM refresh..."

    ' Try various refresh methods
    bomView.Renumber
    If Err.Number = 0 Then
        WScript.Echo "  BOMView.Renumber() completed"
    Else
        WScript.Echo "  BOMView.Renumber() failed - " & Err.Description
        Err.Clear
    End If

    asmDoc.Update
    WScript.Echo "  Document updated"

    asmDoc.Rebuild2 True
    WScript.Echo "  Document rebuilt"

    WScript.Echo ""
    WScript.Echo "=== TEST COMPLETE ==="
    WScript.Echo ""
    WScript.Echo "Check the BOM in Inventor now."
    WScript.Echo "Do quantities show 0 decimals?"
    WScript.Echo ""
    WScript.Echo "If YES: Direct BOM manipulation works!"
    WScript.Echo "If NO: The issue is elsewhere (iProperties, custom formatting, etc.)"

End Sub

Main</content>
<parameter name="filePath">c:\Users\Quintin\Documents\Spectiv\3. Working\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\iLogic_Tools\BOM_Direct_Manipulation_Test.vbs