' Diagnose Plate Values - Check what's actually happening with iProperties
' This reads and reports on all plate parts to find the BOM refresh issue

Option Explicit

Dim m_InventorApp

Sub Main()
    On Error Resume Next
    
    WScript.Echo "=== PLATE VALUES DIAGNOSTIC ==="
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
    
    ' Scan for plate parts
    Dim occurrences
    Set occurrences = asmDoc.ComponentDefinition.Occurrences
    
    Dim processedParts
    Set processedParts = CreateObject("Scripting.Dictionary")
    
    Dim i, occ, doc, partNumber
    For i = 1 To occurrences.Count
        Set occ = occurrences.Item(i)
        If Not occ.Suppressed Then
            Set doc = occ.Definition.Document
            
            If LCase(Right(doc.FullFileName, 4)) = ".ipt" Then
                partNumber = GetPartNumber(doc)
                
                If InStr(1, partNumber, "PL", vbTextCompare) > 0 Then
                    If Not processedParts.Exists(doc.FullFileName) Then
                        processedParts.Add doc.FullFileName, True
                        DiagnosePart doc, partNumber
                    End If
                End If
            End If
        End If
    Next
    
    WScript.Echo ""
    WScript.Echo "=== DIAGNOSTIC COMPLETE ==="
End Sub

Function GetPartNumber(doc)
    On Error Resume Next
    Dim propSet
    Set propSet = doc.PropertySets.Item("Design Tracking Properties")
    GetPartNumber = propSet.Item("Part Number").Value
    If Err.Number <> 0 Then GetPartNumber = doc.DisplayName
    Err.Clear
End Function

Sub DiagnosePart(doc, partNumber)
    On Error Resume Next
    
    WScript.Echo "----------------------------------------"
    WScript.Echo "PART: " & partNumber
    WScript.Echo "File: " & doc.FullFileName
    
    ' Check if sheet metal
    Const kSheetMetalSubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
    Dim isSheetMetal
    isSheetMetal = (doc.SubType = kSheetMetalSubType)
    WScript.Echo "  Sheet Metal: " & isSheetMetal
    
    If Not isSheetMetal Then
        WScript.Echo "  *** NOT SHEET METAL - skipping ***"
        Exit Sub
    End If
    
    ' Get flat pattern dimensions
    Dim compDef
    Set compDef = doc.ComponentDefinition
    
    Dim fpLength, fpWidth
    fpLength = 0
    fpWidth = 0
    
    If compDef.HasFlatPattern Then
        Dim fp
        Set fp = compDef.FlatPattern
        fpLength = fp.Length * 10  ' cm to mm
        fpWidth = fp.Width * 10
        WScript.Echo "  Flat Pattern: " & FormatNumber(fpLength, 1) & " x " & FormatNumber(fpWidth, 1) & " mm"
    Else
        WScript.Echo "  Flat Pattern: NONE"
    End If
    
    ' Get iProperty values
    Dim customProps
    Set customProps = doc.PropertySets.Item("Inventor User Defined Properties")
    
    Dim plateLengthProp, plateWidthProp
    Dim plateLengthValue, plateWidthValue
    Dim plateLengthExpression, plateWidthExpression
    
    ' PLATE LENGTH
    On Error Resume Next
    Set plateLengthProp = customProps.Item("PLATE LENGTH")
    If Err.Number = 0 And Not plateLengthProp Is Nothing Then
        plateLengthValue = plateLengthProp.Value
        plateLengthExpression = plateLengthProp.Expression
        WScript.Echo "  PLATE LENGTH Value: " & plateLengthValue
        WScript.Echo "  PLATE LENGTH Expression: " & plateLengthExpression
        WScript.Echo "  PLATE LENGTH DisplayString: " & plateLengthProp.DisplayString
    Else
        WScript.Echo "  PLATE LENGTH: NOT FOUND"
    End If
    Err.Clear
    
    ' PLATE WIDTH
    On Error Resume Next
    Set plateWidthProp = customProps.Item("PLATE WIDTH")
    If Err.Number = 0 And Not plateWidthProp Is Nothing Then
        plateWidthValue = plateWidthProp.Value
        plateWidthExpression = plateWidthProp.Expression
        WScript.Echo "  PLATE WIDTH Value: " & plateWidthValue
        WScript.Echo "  PLATE WIDTH Expression: " & plateWidthExpression
        WScript.Echo "  PLATE WIDTH DisplayString: " & plateWidthProp.DisplayString
    Else
        WScript.Echo "  PLATE WIDTH: NOT FOUND"
    End If
    Err.Clear
    
    ' Check Document Settings
    Dim params
    Set params = compDef.Parameters
    
    WScript.Echo "  LinearDimensionPrecision: " & params.LinearDimensionPrecision
    WScript.Echo "  DimensionDisplayType: " & params.DimensionDisplayType
    WScript.Echo "  DisplayParameterAsExpression: " & params.DisplayParameterAsExpression
    
    ' Check for mismatch
    If fpLength > 0 Then
        If IsNumeric(plateLengthValue) Then
            If Abs(CDbl(plateLengthValue) - fpLength) > 1 Then
                WScript.Echo "  *** MISMATCH: PLATE LENGTH iProperty doesn't match flat pattern! ***"
            End If
        End If
        If IsNumeric(plateWidthValue) Then
            If Abs(CDbl(plateWidthValue) - fpWidth) > 1 Then
                WScript.Echo "  *** MISMATCH: PLATE WIDTH iProperty doesn't match flat pattern! ***"
            End If
        End If
    End If
    
    WScript.Echo ""
End Sub

Main
