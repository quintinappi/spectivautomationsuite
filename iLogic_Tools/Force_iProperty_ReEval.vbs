' Force iProperty Re-evaluation via Precision Toggle
' Mimics manual Document Settings → Units → Precision change/changeback/save
' Author: Quintin de Bruin © 2026

Option Explicit

Const kAssemblyDocumentObject = 12291
Const kZeroDecimalPlaceLinearPrecision = 0

Dim m_InventorApp

Sub Main()
    On Error Resume Next

    WScript.Echo "=== FORCE IPROPERTY RE-EVALUATION ==="
    WScript.Echo "Mimics manual precision toggle in Document Settings"
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

                    If ForcePrecisionToggle(refDoc) Then
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
    WScript.Echo ""
    WScript.Echo "Check BOM - Stock Number should show whole numbers"

End Sub

Function ForcePrecisionToggle(partDoc)
    On Error Resume Next

    ForcePrecisionToggle = False

    Dim compDef
    Set compDef = partDoc.ComponentDefinition

    Dim params
    Set params = compDef.Parameters

    If params Is Nothing Then
        WScript.Echo "  ERROR: Could not access parameters"
        Exit Function
    End If

    WScript.Echo "  Current precision: " & params.LinearDimensionPrecision

    ' Toggle precision to trigger re-evaluation (like manual UI change)
    WScript.Echo "  Toggling precision: 0 -> 3 -> 0"
    
    params.LinearDimensionPrecision = 3
    If Err.Number <> 0 Then
        WScript.Echo "  ERROR changing to 3: " & Err.Description
        Err.Clear
        Exit Function
    End If

    params.LinearDimensionPrecision = kZeroDecimalPlaceLinearPrecision
    If Err.Number <> 0 Then
        WScript.Echo "  ERROR changing back to 0: " & Err.Description
        Err.Clear
        Exit Function
    End If

    ' Save (like clicking OK in Document Settings)
    partDoc.Save
    If Err.Number <> 0 Then
        WScript.Echo "  ERROR saving: " & Err.Description
        Err.Clear
        Exit Function
    End If

    WScript.Echo "  Final precision: " & params.LinearDimensionPrecision
    ForcePrecisionToggle = True

End Function

Main