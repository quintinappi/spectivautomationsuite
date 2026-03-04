' Force BOM Refresh - Nuclear option to force assembly to re-read all part properties
' This mimics what happens when you manually open a part and save it

Option Explicit

' UnitsTypeEnum constants for forcing BOM refresh via units toggle
Const kMillimeterLengthUnits = 11269  ' mm (most common)
Const kCentimeterLengthUnits = 11266  ' cm

Dim m_InventorApp

Sub ForceUnitsRefreshEvent(partDoc)
    ' Forces BOM to refresh display format by triggering UnitsOfMeasure change event
    ' This mimics the manual UI workaround of toggling units without saving
    On Error Resume Next

    WScript.Echo "  Triggering UnitsOfMeasure change event for: " & partDoc.DisplayName

    Dim unitsOfMeasure
    Set unitsOfMeasure = partDoc.UnitsOfMeasure

    If unitsOfMeasure Is Nothing Then
        WScript.Echo "    ERROR: Could not access UnitsOfMeasure object"
        Exit Sub
    End If

    ' Get current length units
    Dim originalLengthUnits
    originalLengthUnits = unitsOfMeasure.LengthUnits

    ' Toggle to different unit and back
    Dim tempUnits
    If originalLengthUnits = kMillimeterLengthUnits Then
        tempUnits = kCentimeterLengthUnits
    Else
        tempUnits = kMillimeterLengthUnits
    End If

    ' Change units (triggers event)
    unitsOfMeasure.LengthUnits = tempUnits
    If Err.Number <> 0 Then
        Err.Clear
    End If

    partDoc.Update

    ' Restore original units
    unitsOfMeasure.LengthUnits = originalLengthUnits
    If Err.Number <> 0 Then
        Err.Clear
    End If

    partDoc.Update
    WScript.Echo "    Units toggle complete"
End Sub

Sub Main()
    On Error Resume Next
    
    WScript.Echo "=== FORCE BOM REFRESH ==="
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
    
    If m_InventorApp.ActiveDocument.DocumentType <> 12291 Then
        WScript.Echo "ERROR: Not an assembly"
        Exit Sub
    End If
    
    Dim asmDoc
    Set asmDoc = m_InventorApp.ActiveDocument
    WScript.Echo "Assembly: " & asmDoc.DisplayName
    WScript.Echo ""
    
    ' Method 1: Force assembly Rebuild (not just Update)
    WScript.Echo "Method 1: Forcing assembly REBUILD..."
    asmDoc.Rebuild
    If Err.Number <> 0 Then
        WScript.Echo "  Rebuild failed: " & Err.Description
        Err.Clear
    Else
        WScript.Echo "  Rebuild complete"
    End If
    
    ' Method 2: Rebuild2 with AcceptErrorsAndContinue
    WScript.Echo "Method 2: Forcing assembly REBUILD2..."
    asmDoc.Rebuild2 True
    If Err.Number <> 0 Then
        WScript.Echo "  Rebuild2 failed: " & Err.Description
        Err.Clear
    Else
        WScript.Echo "  Rebuild2 complete"
    End If
    
    ' Method 3: Touch the BOM directly
    WScript.Echo "Method 3: Refreshing BOM views..."
    Dim compDef
    Set compDef = asmDoc.ComponentDefinition
    
    Dim bom
    Set bom = compDef.BOM
    
    If Not bom Is Nothing Then
        ' Check BOM enabled status
        WScript.Echo "  BOM.StructuredViewEnabled: " & bom.StructuredViewEnabled
        WScript.Echo "  BOM.PartsOnlyViewEnabled: " & bom.PartsOnlyViewEnabled
        
        ' Toggle structured view off and on to force refresh
        If bom.StructuredViewEnabled Then
            WScript.Echo "  Toggling StructuredView..."
            bom.StructuredViewEnabled = False
            bom.StructuredViewEnabled = True
        End If
        
        If bom.PartsOnlyViewEnabled Then
            WScript.Echo "  Toggling PartsOnlyView..."
            bom.PartsOnlyViewEnabled = False
            bom.PartsOnlyViewEnabled = True
        End If
    Else
        WScript.Echo "  Could not access BOM object"
    End If
    
    ' Method 4: Force all occurrences to refresh
    WScript.Echo "Method 4: Touching all occurrences..."
    Dim occurrences
    Set occurrences = compDef.Occurrences
    
    Dim i, occ, doc
    Dim plateCount
    plateCount = 0
    
    For i = 1 To occurrences.Count
        Set occ = occurrences.Item(i)
        If Not occ.Suppressed Then
            Set doc = occ.Definition.Document
            
            If LCase(Right(doc.FullFileName, 4)) = ".ipt" Then
                Dim desc1
                desc1 = ""
                On Error Resume Next
                desc1 = doc.PropertySets.Item("Design Tracking Properties").Item("Description").Value
                Err.Clear
                
                ' Check Description for PL, VRN, or S355JR
                If InStr(1, desc1, "PL", vbTextCompare) > 0 Or InStr(1, desc1, "VRN", vbTextCompare) > 0 Or InStr(1, desc1, "S355JR", vbTextCompare) > 0 Then
                    plateCount = plateCount + 1
                    
                    ' Force the document to mark as needing update
                    doc.Update2 True
                    If Err.Number <> 0 Then
                        Err.Clear
                    End If
                End If
            End If
        End If
    Next
    WScript.Echo "  Touched " & plateCount & " plate occurrences"

    ' Method 5: Trigger UnitsOfMeasure change event on each plate part (THE REAL FIX!)
    WScript.Echo "Method 5: Triggering UnitsOfMeasure events on plate parts..."
    Dim changedCount
    changedCount = 0

    For i = 1 To occurrences.Count
        Set occ = occurrences.Item(i)
        If Not occ.Suppressed Then
            Set doc = occ.Definition.Document

            If LCase(Right(doc.FullFileName, 4)) = ".ipt" Then
                Dim desc2
                desc2 = ""
                On Error Resume Next
                desc2 = doc.PropertySets.Item("Design Tracking Properties").Item("Description").Value
                Err.Clear

                ' Check Description for PL, VRN, or S355JR
                If InStr(1, desc2, "PL", vbTextCompare) > 0 Or InStr(1, desc2, "VRN", vbTextCompare) > 0 Or InStr(1, desc2, "S355JR", vbTextCompare) > 0 Then
                    ' Trigger units change event for this plate
                    Call ForceUnitsRefreshEvent(doc)
                    changedCount = changedCount + 1
                End If
            End If
        End If
    Next
    WScript.Echo "  Triggered UnitsOfMeasure events on " & changedCount & " plate parts"

    ' Method 6: Final assembly update to propagate all changes
    WScript.Echo "Method 6: Final assembly update..."
    asmDoc.Update2 True
    If Err.Number <> 0 Then
        WScript.Echo "  Update2 failed: " & Err.Description
        Err.Clear
    Else
        WScript.Echo "  Update2 complete"
    End If
    
    ' Method 6: Save and reopen assembly
    WScript.Echo ""
    WScript.Echo "Saving assembly..."
    asmDoc.Save
    WScript.Echo "Assembly saved"
    
    WScript.Echo ""
    WScript.Echo "=== REFRESH COMPLETE ==="
    WScript.Echo "Check the BOM now. If still not updated, try:"
    WScript.Echo "  1. Close and reopen the assembly"
    WScript.Echo "  2. In the BOM dialog, right-click > Refresh"
    WScript.Echo "  3. Export BOM to Excel and check values there"
    
End Sub

Main
