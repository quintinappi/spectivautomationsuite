' Force UI Document Settings Interaction
' Opens Document Settings programmatically to trigger re-evaluation
' Author: Quintin de Bruin © 2026

Option Explicit

Const kAssemblyDocumentObject = 12291
Const kMillimeterLengthUnits = 11269
Const kCentimeterLengthUnits = 11266

Dim m_InventorApp

Sub Main()
    On Error Resume Next

    WScript.Echo "=== FORCE UI DOCUMENT SETTINGS INTERACTION ==="
    WScript.Echo ""

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

    Dim occurrences
    Set occurrences = asmDoc.ComponentDefinition.Occurrences

    Dim updatedCount
    updatedCount = 0

    Dim i
    For i = 1 To occurrences.Count
        Dim occ
        Set occ = occurrences.Item(i)

        If Not occ.Suppressed Then
            Dim refDoc
            Set refDoc = occ.Definition.Document

            If LCase(Right(refDoc.FullFileName, 4)) = ".ipt" Then
                Dim desc
                desc = ""
                On Error Resume Next
                desc = refDoc.PropertySets.Item("Design Tracking Properties").Item("Description").Value
                Err.Clear

                ' Check Description for PL, VRN, or S355JR
                If InStr(UCase(desc), "PL") > 0 Or InStr(UCase(desc), "VRN") > 0 Or InStr(UCase(desc), "S355JR") > 0 Then
                    Dim partNumber
                    partNumber = ""
                    On Error Resume Next
                    partNumber = refDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value
                    Err.Clear
                    
                    WScript.Echo "Processing: " & partNumber & " (Desc: " & desc & ")"

                    If ForceCompleteUpdate(refDoc) Then
                        updatedCount = updatedCount + 1
                        WScript.Echo "  SUCCESS"
                    Else
                        WScript.Echo "  FAILED"
                    End If

                    WScript.Echo ""
                End If
            End If
        End If
    Next

    WScript.Echo "=== COMPLETE ==="
    WScript.Echo "Parts updated: " & updatedCount

End Sub

Function ForceCompleteUpdate(partDoc)
    On Error Resume Next

    ForceCompleteUpdate = False

    Dim compDef
    Set compDef = partDoc.ComponentDefinition

    Dim params
    Set params = compDef.Parameters

    Dim unitsOfMeasure
    Set unitsOfMeasure = partDoc.UnitsOfMeasure

    If params Is Nothing Or unitsOfMeasure Is Nothing Then
        WScript.Echo "  ERROR: Could not access parameters or units"
        Exit Function
    End If

    ' Get current states
    Dim originalPrecision
    originalPrecision = params.LinearDimensionPrecision

    Dim originalUnits
    originalUnits = unitsOfMeasure.LengthUnits

    WScript.Echo "  Original precision: " & originalPrecision & ", units: " & originalUnits

    ' COMBO APPROACH: Toggle both units AND precision together
    WScript.Echo "  Toggling units + precision..."

    ' Change units
    If originalUnits = kMillimeterLengthUnits Then
        unitsOfMeasure.LengthUnits = kCentimeterLengthUnits
    Else
        unitsOfMeasure.LengthUnits = kMillimeterLengthUnits
    End If

    ' Change precision
    params.LinearDimensionPrecision = 3

    ' Force update
    partDoc.Update

    ' Change back
    unitsOfMeasure.LengthUnits = originalUnits
    params.LinearDimensionPrecision = originalPrecision

    ' Force update again
    partDoc.Update

    ' Rebuild
    partDoc.Rebuild

    ' Save
    partDoc.Save
    If Err.Number <> 0 Then
        WScript.Echo "  ERROR saving: " & Err.Description
        Err.Clear
        Exit Function
    End If

    ForceCompleteUpdate = True

End Function

Main