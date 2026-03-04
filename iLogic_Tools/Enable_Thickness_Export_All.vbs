' =========================================================
' ENABLE THICKNESS EXPORT - ALL PLATE PARTS (API METHOD)
' =========================================================
' Uses the ExposedAsProperty property to enable
' the Export Parameter checkbox for Thickness on all plates
' =========================================================

Option Explicit

Dim m_InventorApp

Main()

Sub Main()
    On Error Resume Next
    
    WScript.Echo "=== ENABLE THICKNESS EXPORT - ALL PLATES ==="
    WScript.Echo ""
    
    Set m_InventorApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Could not connect to Inventor."
        WScript.Quit 1
    End If
    
    Dim activeDoc
    Set activeDoc = m_InventorApp.ActiveDocument
    
    If activeDoc Is Nothing Then
        WScript.Echo "ERROR: No active document."
        WScript.Quit 1
    End If
    
    ' Check if it's an assembly
    If activeDoc.DocumentType <> 12291 Then ' kAssemblyDocumentObject
        WScript.Echo "ERROR: Not an assembly"
        WScript.Quit 1
    End If
    
    WScript.Echo "Assembly: " & activeDoc.DisplayName
    WScript.Echo ""
    
    ' Get BOM to find plate parts
    Dim bom
    Set bom = activeDoc.ComponentDefinition.BOM
    bom.StructuredViewEnabled = True
    bom.StructuredViewFirstLevelOnly = False
    
    Dim bomView
    Set bomView = bom.BOMViews.Item("Structured")
    
    WScript.Echo "Scanning BOM for plate parts..."
    WScript.Echo ""
    
    ' Collect all unique plate part documents
    Dim plateParts()
    Dim platePartCount
    platePartCount = 0
    
    Dim i
    For i = 1 To bomView.BOMRows.Count
        Dim bomRow
        Set bomRow = bomView.BOMRows.Item(i)
        
        Dim compDef
        Set compDef = bomRow.ComponentDefinitions.Item(1)
        Dim partDoc
        Set partDoc = compDef.Document
        
        Dim partNum
        partNum = ""
        On Error Resume Next
        partNum = partDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value
        On Error Goto 0
        
        If InStr(1, partNum, "PL", vbTextCompare) > 0 Then
            ' Check if already in list
            Dim j
            Dim isNewPart
            isNewPart = True
            
            For j = 0 To platePartCount - 1
                If plateParts(j).FullDocumentName = partDoc.FullDocumentName Then
                    isNewPart = False
                    Exit For
                End If
            Next
            
            If isNewPart Then
                If platePartCount = 0 Then
                    ReDim plateParts(0)
                Else
                    ReDim Preserve plateParts(platePartCount)
                End If
                Set plateParts(platePartCount) = partDoc
                platePartCount = platePartCount + 1
            End If
        End If
    Next
    
    If platePartCount = 0 Then
        WScript.Echo "No plate parts found."
        WScript.Quit 0
    End If
    
    WScript.Echo "Found " & platePartCount & " unique plate part(s)"
    WScript.Echo ""
    WScript.Echo "Enabling Thickness export for each part..."
    WScript.Echo ""
    
    ' Process each plate part
    Dim processedCount
    processedCount = 0
    
    For i = 0 To platePartCount - 1
        Dim plateDoc
        Set plateDoc = plateParts(i)
        
        Dim plateName
        plateName = ""
        On Error Resume Next
        plateName = plateDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value
        On Error Goto 0
        
        WScript.Echo "  [" & (i + 1) & "/" & platePartCount & "] " & plateName
        
        ' Get Parameters
        Dim params
        Set params = plateDoc.ComponentDefinition.Parameters
        
        ' Find Thickness in ModelParameters
        Dim modelParams
        Set modelParams = params.ModelParameters
        
        Dim thicknessParam
        Set thicknessParam = Nothing
        
        Dim k
        For k = 1 To modelParams.Count
            If modelParams.Item(k).Name = "Thickness" Then
                Set thicknessParam = modelParams.Item(k)
                Exit For
            End If
        Next
        
        If thicknessParam Is Nothing Then
            WScript.Echo "        - WARNING: No Thickness parameter found"
        Else
            ' Enable export
            On Error Resume Next
            thicknessParam.ExposedAsProperty = True
            If Err.Number = 0 Then
                plateDoc.Save
                WScript.Echo "        - ExposedAsProperty enabled, saved"
                processedCount = processedCount + 1
            Else
                WScript.Echo "        - ERROR: " & Err.Description
            End If
            Err.Clear
            On Error Goto 0
        End If
    Next
    
    WScript.Echo ""
    WScript.Echo "================================"
    WScript.Echo "Completed: " & processedCount & " part(s) updated"
    WScript.Echo "================================"
    WScript.Echo ""
    
End Sub
