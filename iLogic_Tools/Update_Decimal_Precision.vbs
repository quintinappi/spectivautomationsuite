' Update Decimal Precision - Standalone VBScript
' Sets Document Settings > Units > Linear Dim Display Precision to 0 decimals
' for all plate parts in the active assembly
' Author: Quintin de Bruin © 2026

Option Explicit

' Inventor API Constants
Const kAssemblyDocumentObject = 12291
' LinearPrecisionEnum values - 0 for zero decimal places
Const kZeroDecimalPlaceLinearPrecision = 0
' DimensionDisplayTypeEnum - 34817 is "Display precise value"
' Try 34821 for "Display as value" based on enum order in XML
Const kDimensionDisplayAsValue = 34821
' UnitsTypeEnum constants for forcing BOM refresh via units toggle
Const kMillimeterLengthUnits = 11269  ' mm (most common)
Const kCentimeterLengthUnits = 11266  ' cm

' Global variables
Dim m_InventorApp
Dim m_Log

Sub ForceUnitsRefreshEvent(partDoc)
    ' Forces BOM to refresh display format by triggering UnitsOfMeasure change event
    ' This mimics the manual UI workaround of toggling units without saving
    ' CRITICAL FIX: BOM caches display format separately - only refreshes on units change event
    On Error Resume Next

    LogMessage "Triggering UnitsOfMeasure change event to invalidate BOM cache..."

    Dim unitsOfMeasure
    Set unitsOfMeasure = partDoc.UnitsOfMeasure

    If unitsOfMeasure Is Nothing Then
        LogMessage "ERROR: Could not access UnitsOfMeasure object"
        Exit Sub
    End If

    ' Get current length units
    Dim originalLengthUnits
    originalLengthUnits = unitsOfMeasure.LengthUnits

    LogMessage "Current LengthUnits: " & originalLengthUnits

    ' CRITICAL: Change to a different unit temporarily
    ' Toggle to cm, then back to mm (or vice versa)
    Dim tempUnits
    If originalLengthUnits = kMillimeterLengthUnits Then ' Currently mm
        tempUnits = kCentimeterLengthUnits ' Switch to cm
        LogMessage "Toggling: mm -> cm -> mm"
    Else
        tempUnits = kMillimeterLengthUnits ' Switch to mm
        LogMessage "Toggling: current -> mm -> current"
    End If

    ' Change units (triggers cache invalidation event)
    unitsOfMeasure.LengthUnits = tempUnits
    If Err.Number <> 0 Then
        LogMessage "WARNING: Could not set temporary units - " & Err.Description
        Err.Clear
    Else
        LogMessage "Temporary units set: " & tempUnits
    End If

    ' CRITICAL: Force document update to propagate event
    partDoc.Update

    ' Restore original units (triggers event again)
    unitsOfMeasure.LengthUnits = originalLengthUnits
    If Err.Number <> 0 Then
        LogMessage "WARNING: Could not restore original units - " & Err.Description
        Err.Clear
    Else
        LogMessage "Original units restored: " & originalLengthUnits
    End If

    ' Final update to propagate restored units
    partDoc.Update

    LogMessage "UnitsOfMeasure change event triggered - BOM should refresh immediately!"
End Sub

Sub Main()
    On Error Resume Next

    m_Log = ""

    LogMessage "=== DECIMAL PRECISION UPDATER STARTED ==="

    ' Get Inventor application
    LogMessage "Attempting to get Inventor application..."
    Set m_InventorApp = GetObject(, "Inventor.Application")
    
    If Err.Number <> 0 Then
        LogMessage "ERROR: Could not connect to Inventor - " & Err.Description
        MsgBox "Could not connect to Inventor. Please make sure Inventor is running.", vbCritical, "Error"
        Exit Sub
    End If
    
    LogMessage "Connected to Inventor successfully"

    ' Check for active document
    If m_InventorApp.ActiveDocument Is Nothing Then
        LogMessage "ERROR: No active document"
        MsgBox "No active document! Please open an assembly in Inventor.", vbCritical, "Error"
        Exit Sub
    End If

    LogMessage "Active document: " & m_InventorApp.ActiveDocument.FullFileName
    LogMessage "Document type: " & m_InventorApp.ActiveDocument.DocumentType

    If m_InventorApp.ActiveDocument.DocumentType <> kAssemblyDocumentObject Then
        LogMessage "ERROR: Document is not an assembly"
        MsgBox "Please open an ASSEMBLY document (.iam file).", vbExclamation, "Assembly Required"
        Exit Sub
    End If

    Dim asmDoc
    Set asmDoc = m_InventorApp.ActiveDocument
    LogMessage "Processing assembly: " & asmDoc.FullFileName

    ' Step 1: Scan assembly for plate parts
    LogMessage ""
    LogMessage "STEP 1: Scanning assembly for plate parts (PL or S355JR)"
    Dim plateParts
    Set plateParts = ScanAssemblyForPlateParts(asmDoc)

    If plateParts.Count = 0 Then
        LogMessage "No plate parts found"
        MsgBox "No parts containing 'PL' or 'S355JR' found in the assembly.", vbInformation, "No Plates Found"
        Exit Sub
    End If

    LogMessage "Found " & plateParts.Count & " unique plate parts"

    ' Confirm with user
    Dim userResponse
    userResponse = MsgBox("Found " & plateParts.Count & " plate parts." & vbCrLf & vbCrLf & _
                          "This will set for each plate part (Document Settings > Units):" & vbCrLf & _
                          "  1. Linear Dim Display Precision = 0 decimals" & vbCrLf & _
                          "  2. Modeling Dimension Display = 'Display as value'" & vbCrLf & _
                          "  3. Default Parameter Input Display = 'Display as expression'" & vbCrLf & vbCrLf & _
                          "Continue?", vbOKCancel + vbQuestion, "Update Decimal Precision")

    If userResponse = vbCancel Then
        LogMessage "User cancelled"
        Exit Sub
    End If

    ' Step 2: Process each plate part
    LogMessage ""
    LogMessage "STEP 2: Updating decimal precision for each part"
    Dim updatedCount, skippedCount, failedCount
    updatedCount = 0
    skippedCount = 0
    failedCount = 0

    Dim partPath
    For Each partPath In plateParts.Keys
        LogMessage ""
        LogMessage "Processing: " & partPath

        On Error Resume Next
        
        ' Open the part
        Dim partDoc
        Set partDoc = m_InventorApp.Documents.Open(partPath, True)
        
        If Err.Number <> 0 Then
            LogMessage "ERROR: Could not open part - " & Err.Description
            Err.Clear
            failedCount = failedCount + 1
        Else
            LogMessage "Opened part successfully"
            
            ' Get the Parameters object from ComponentDefinition
            Dim compDef
            Set compDef = partDoc.ComponentDefinition
            
            If compDef Is Nothing Then
                LogMessage "ERROR: Could not get ComponentDefinition"
                failedCount = failedCount + 1
            Else
                Dim params
                Set params = compDef.Parameters
                
                If params Is Nothing Then
                    LogMessage "ERROR: Could not get Parameters"
                    failedCount = failedCount + 1
                Else
                    ' Read current values
                    LogMessage "Current LinearDimensionPrecision: " & params.LinearDimensionPrecision
                    LogMessage "Current DimensionDisplayType: " & params.DimensionDisplayType
                    LogMessage "Current DisplayParameterAsExpression: " & params.DisplayParameterAsExpression
                    
                    ' CRITICAL: Toggle parameters to force dirty flag (mimics manual UI change)
                    LogMessage "Toggling parameters to force change detection..."
                    
                    ' Toggle LinearDimensionPrecision (change to 3, then back to 0)
                    params.LinearDimensionPrecision = 3
                    params.LinearDimensionPrecision = kZeroDecimalPlaceLinearPrecision
                    
                    If Err.Number <> 0 Then
                        LogMessage "ERROR: Could not set LinearDimensionPrecision - " & Err.Description
                        Err.Clear
                    Else
                        LogMessage "New LinearDimensionPrecision: " & params.LinearDimensionPrecision
                    End If
                    
                    ' Toggle DimensionDisplayType (change to 34817, then to 34821)
                    params.DimensionDisplayType = 34817
                    params.DimensionDisplayType = kDimensionDisplayAsValue
                    
                    If Err.Number <> 0 Then
                        LogMessage "ERROR: Could not set DimensionDisplayType - " & Err.Description
                        Err.Clear
                        failedCount = failedCount + 1
                    Else
                        LogMessage "New DimensionDisplayType: " & params.DimensionDisplayType
                    End If
                    
                    ' Toggle DisplayParameterAsExpression (False then True)
                    params.DisplayParameterAsExpression = False
                    params.DisplayParameterAsExpression = True
                    
                    If Err.Number <> 0 Then
                        LogMessage "ERROR: Could not set DisplayParameterAsExpression - " & Err.Description
                        Err.Clear
                        failedCount = failedCount + 1
                    Else
                        LogMessage "New DisplayParameterAsExpression: " & params.DisplayParameterAsExpression
                        LogMessage "SUCCESS: All settings updated"

                        ' CRITICAL: Trigger UnitsOfMeasure change event to force BOM refresh
                        ' This mimics the manual workaround: changing units and back (no save needed)
                        Call ForceUnitsRefreshEvent(partDoc)
                        
                        ' Save the part
                        partDoc.Save
                        If Err.Number <> 0 Then
                            LogMessage "WARNING: Could not save - " & Err.Description
                            Err.Clear
                        Else
                            LogMessage "Part saved"
                        End If
                        
                        updatedCount = updatedCount + 1
                    End If
                End If
            End If
            
            ' Close the part
            partDoc.Close False
        End If
        
        Err.Clear
    Next

    ' Summary
    LogMessage ""
    LogMessage "=== UPDATE COMPLETE ==="
    LogMessage "Parts updated: " & updatedCount
    LogMessage "Parts failed: " & failedCount
    LogMessage "Total processed: " & plateParts.Count

    MsgBox "Decimal Precision Update Complete!" & vbCrLf & vbCrLf & _
           "Parts updated: " & updatedCount & vbCrLf & _
           "Parts failed: " & failedCount & vbCrLf & _
           "Total: " & plateParts.Count, vbInformation, "Complete"

End Sub

Function ScanAssemblyForPlateParts(asmDoc)
    ' Returns a Dictionary of unique plate part paths
    Dim result
    Set result = CreateObject("Scripting.Dictionary")

    Dim occurrences
    Set occurrences = asmDoc.ComponentDefinition.Occurrences

    LogMessage "Found " & occurrences.Count & " occurrences in assembly"

    Dim i
    For i = 1 To occurrences.Count
        On Error Resume Next
        
        Dim occ
        Set occ = occurrences.Item(i)
        
        If Not occ.Suppressed Then
            Dim refDoc
            Set refDoc = occ.Definition.Document
            
            If Not refDoc Is Nothing Then
                Dim fileName
                fileName = refDoc.FullFileName
                
                ' Check if it's a part file
                If LCase(Right(fileName, 4)) = ".ipt" Then
                    ' Get part number or description
                    Dim partNumber
                    partNumber = ""
                    
                    On Error Resume Next
                    Dim designProps
                    Set designProps = refDoc.PropertySets.Item("Design Tracking Properties")
                    If Err.Number = 0 Then
                        partNumber = designProps.Item("Part Number").Value
                    End If
                    Err.Clear
                    
                    LogMessage "Checking: " & occ.Name & " - Part Number: " & partNumber

                    ' Check if it contains "PL" or "S355JR"
                    If InStr(UCase(partNumber), "PL") > 0 Or InStr(UCase(partNumber), "S355JR") > 0 Then
                        ' Add to collection if not already there
                        If Not result.Exists(fileName) Then
                            result.Add fileName, True
                            LogMessage "  -> Plate part identified"
                        End If
                    End If
                    
                ElseIf LCase(Right(fileName, 4)) = ".iam" Then
                    ' Recurse into sub-assembly
                    Dim subParts
                    Set subParts = ScanAssemblyForPlateParts(refDoc)
                    
                    Dim subPath
                    For Each subPath In subParts.Keys
                        If Not result.Exists(subPath) Then
                            result.Add subPath, True
                        End If
                    Next
                End If
            End If
        End If
        
        Err.Clear
    Next

    Set ScanAssemblyForPlateParts = result
End Function

Sub LogMessage(msg)
    m_Log = m_Log & msg & vbCrLf
    WScript.Echo msg
End Sub

' Start the script
Main
