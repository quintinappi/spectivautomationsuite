' Fix BOM Plate Dimensions - Direct BOM Column Injection
' Adds WIDTH column to BOM and injects formulas directly into BOM cells
' for plate parts only.
' Author: Quintin de Bruin © 2026
'
' This version works directly with the BOM PropertyColumns and cells
' rather than setting custom iProperties on parts.

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
    
    LogMessage "=== FIX BOM PLATE DIMENSIONS (DIRECT) ==="
    LogMessage "Date/Time: " & Now
    LogMessage ""

    ' Get Inventor application
    Set m_InventorApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        MsgBox "Inventor is not running. Please start Inventor first.", vbCritical, "Error"
        Exit Sub
    End If

    If m_InventorApp.ActiveDocument Is Nothing Then
        MsgBox "No active document! Please open an assembly in Inventor.", vbCritical, "Error"
        Exit Sub
    End If

    If m_InventorApp.ActiveDocument.DocumentType <> kAssemblyDocumentObject Then
        MsgBox "Please open an ASSEMBLY document (.iam file).", vbExclamation, "Assembly Required"
        Exit Sub
    End If

    Dim asmDoc
    Set asmDoc = m_InventorApp.ActiveDocument
    LogMessage "Assembly: " & asmDoc.DisplayName
    LogMessage ""

    ' Get BOM
    Dim bom
    Set bom = asmDoc.ComponentDefinition.BOM
    
    ' Enable structured view if not already
    If Not bom.StructuredViewEnabled Then
        LogMessage "Enabling Structured BOM view..."
        bom.StructuredViewEnabled = True
    End If
    
    ' Get the structured BOM view
    Dim bomView
    Set bomView = bom.BOMViews.Item("Structured")
    If bomView Is Nothing Then
        MsgBox "Could not access Structured BOM view", vbCritical, "Error"
        Exit Sub
    End If
    
    LogMessage "BOM Rows: " & bomView.BOMRows.Count
    LogMessage ""
    
    ' List existing columns
    LogMessage "=== EXISTING BOM COLUMNS ==="
    Dim col
    Dim colIndex
    colIndex = 1
    Dim lengthColIndex, widthColIndex
    lengthColIndex = -1
    widthColIndex = -1
    
    For Each col In bomView.BOMPropertyColumns
        LogMessage "  " & colIndex & ": " & col.PropertyDefinition.DisplayName
        If UCase(col.PropertyDefinition.DisplayName) = "LENGTH" Then
            lengthColIndex = colIndex
            LogMessage "     ^ Found LENGTH column at index " & colIndex
        End If
        If UCase(col.PropertyDefinition.DisplayName) = "WIDTH" Then
            widthColIndex = colIndex
            LogMessage "     ^ Found WIDTH column at index " & colIndex
        End If
        colIndex = colIndex + 1
    Next
    LogMessage ""
    
    ' Check if we need to add columns
    If lengthColIndex = -1 Then
        LogMessage "LENGTH column not found - attempting to add..."
        AddCustomPropertyColumn bomView, "LENGTH"
        ' Re-scan for column index
        colIndex = 1
        For Each col In bomView.BOMPropertyColumns
            If UCase(col.PropertyDefinition.DisplayName) = "LENGTH" Then
                lengthColIndex = colIndex
                LogMessage "  LENGTH column added at index " & colIndex
                Exit For
            End If
            colIndex = colIndex + 1
        Next
    End If
    
    If widthColIndex = -1 Then
        LogMessage "WIDTH column not found - attempting to add..."
        AddCustomPropertyColumn bomView, "WIDTH"
        ' Re-scan for column index
        colIndex = 1
        For Each col In bomView.BOMPropertyColumns
            If UCase(col.PropertyDefinition.DisplayName) = "WIDTH" Then
                widthColIndex = colIndex
                LogMessage "  WIDTH column added at index " & colIndex
                Exit For
            End If
            colIndex = colIndex + 1
        Next
    End If
    
    LogMessage ""
    LogMessage "=== PROCESSING BOM ROWS ==="
    
    Dim processedCount
    processedCount = 0
    
    ' Process each BOM row
    Dim bomRow
    For Each bomRow In bomView.BOMRows
        ProcessBOMRow bomRow, lengthColIndex, widthColIndex, processedCount
    Next
    
    ' Force BOM refresh
    LogMessage ""
    LogMessage "=== REFRESHING BOM ==="
    asmDoc.Update
    
    LogMessage ""
    LogMessage "=== COMPLETE ==="
    LogMessage "Plate parts updated: " & processedCount
    
    MsgBox "Processing complete!" & vbCrLf & vbCrLf & _
           "Plate parts updated: " & processedCount, vbInformation, "Fix BOM Plate Dimensions"

    SaveLog
End Sub

Sub ProcessBOMRow(bomRow, lengthColIndex, widthColIndex, ByRef processedCount)
    On Error Resume Next
    
    ' Get the component definition
    Dim compDefs
    Set compDefs = bomRow.ComponentDefinitions
    If compDefs Is Nothing Or compDefs.Count = 0 Then Exit Sub
    
    Dim compDef
    Set compDef = compDefs.Item(1)
    If compDef Is Nothing Then Exit Sub
    
    ' Get the document
    Dim partDoc
    Set partDoc = compDef.Document
    If partDoc Is Nothing Then Exit Sub
    
    ' Get description
    Dim description
    description = GetDescription(partDoc)
    
    ' Check if it's a plate
    If Not IsPlate(description) Then
        ' Process child rows if any
        If Not bomRow.ChildRows Is Nothing Then
            Dim childRow1
            For Each childRow1 In bomRow.ChildRows
                ProcessBOMRow childRow1, lengthColIndex, widthColIndex, processedCount
            Next
        End If
        Exit Sub
    End If
    
    ' It's a plate - get the part number for logging
    Dim partNumber
    partNumber = ""
    partNumber = partDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value
    If Err.Number <> 0 Then
        Err.Clear
        partNumber = partDoc.DisplayName
    End If
    
    LogMessage "Processing plate: " & partNumber
    
    ' Check if it's a sheet metal part
    Dim isSheetMetal
    isSheetMetal = (partDoc.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
    
    If Not isSheetMetal Then
        LogMessage "  Skipped: Not sheet metal"
        Exit Sub
    End If
    
    ' Try to set the cell values using BOMRowProperties
    Dim rowProps
    Set rowProps = bomRow.BOMRowProperties
    
    If Not rowProps Is Nothing Then
        LogMessage "  BOMRowProperties count: " & rowProps.Count
        
        Dim prop
        For Each prop In rowProps
            LogMessage "    Property: " & prop.PropertyDefinition.DisplayName & " = " & prop.Value
            
            ' Check if this is LENGTH column
            If UCase(prop.PropertyDefinition.DisplayName) = "LENGTH" Then
                LogMessage "    Attempting to set LENGTH formula..."
                prop.Expression = "<sheet metal length>"
                If Err.Number <> 0 Then
                    LogMessage "      Expression failed: " & Err.Description
                    Err.Clear
                    prop.Value = "=<sheet metal length>"
                    If Err.Number <> 0 Then
                        LogMessage "      Value failed: " & Err.Description
                        Err.Clear
                    Else
                        LogMessage "      Set LENGTH value"
                    End If
                Else
                    LogMessage "      Set LENGTH expression"
                End If
            End If
            
            ' Check if this is WIDTH column
            If UCase(prop.PropertyDefinition.DisplayName) = "WIDTH" Then
                LogMessage "    Attempting to set WIDTH formula..."
                prop.Expression = "<sheet metal width>"
                If Err.Number <> 0 Then
                    LogMessage "      Expression failed: " & Err.Description
                    Err.Clear
                    prop.Value = "=<sheet metal width>"
                    If Err.Number <> 0 Then
                        LogMessage "      Value failed: " & Err.Description
                        Err.Clear
                    Else
                        LogMessage "      Set WIDTH value"
                    End If
                Else
                    LogMessage "      Set WIDTH expression"
                End If
            End If
        Next
    Else
        LogMessage "  BOMRowProperties is Nothing"
    End If
    
    processedCount = processedCount + 1
    
    ' Process child rows if any
    If Not bomRow.ChildRows Is Nothing Then
        Dim childRow2
        For Each childRow2 In bomRow.ChildRows
            ProcessBOMRow childRow2, lengthColIndex, widthColIndex, processedCount
        Next
    End If
End Sub

Sub AddCustomPropertyColumn(bomView, propName)
    On Error Resume Next
    
    ' Try to add a custom property column
    ' The BOMPropertyColumns.Add method requires a PropertyDefinition
    
    LogMessage "  Attempting to add custom column: " & propName
    
    ' First, try to find the property definition in the property set
    Dim propDefs
    Set propDefs = m_InventorApp.ActiveDocument.PropertySets
    
    ' Look for the custom property
    Dim customPropSet
    Set customPropSet = propDefs.Item("Inventor User Defined Properties")
    
    If customPropSet Is Nothing Then
        LogMessage "  ERROR: Cannot access custom property set"
        Exit Sub
    End If
    
    ' Check if property exists
    Dim customProp
    Set customProp = customPropSet.Item(propName)
    If Err.Number <> 0 Then
        LogMessage "  Property " & propName & " not found in assembly custom properties"
        Err.Clear
    Else
        LogMessage "  Found " & propName & " in custom properties"
    End If
    
    ' Try to add via BOMView's CustomPropertyColumns
    ' The API method is typically: BOMView.AddProperty(PropertySetName, PropertyName)
    Dim addResult
    addResult = bomView.AddProperty("User Defined Properties", propName)
    If Err.Number <> 0 Then
        LogMessage "  AddProperty failed: " & Err.Description
        Err.Clear
        
        ' Alternative approach - try AddColumn
        bomView.AddColumn propName
        If Err.Number <> 0 Then
            LogMessage "  AddColumn failed: " & Err.Description
            Err.Clear
        End If
    Else
        LogMessage "  AddProperty succeeded"
    End If
End Sub

Function GetDescription(doc)
    On Error Resume Next
    GetDescription = ""
    
    If doc Is Nothing Then Exit Function
    
    Dim propSets
    Set propSets = doc.PropertySets
    If propSets Is Nothing Then Exit Function
    
    Dim designProps
    Set designProps = propSets.Item("Design Tracking Properties")
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    
    Dim descProp
    Set descProp = designProps.Item("Description")
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    
    GetDescription = Trim(descProp.Value)
End Function

Function IsPlate(description)
    Dim upperDesc
    upperDesc = UCase(description)
    
    If InStr(upperDesc, "PL") > 0 Then
        IsPlate = True
        Exit Function
    End If
    
    If InStr(upperDesc, "VRN") > 0 Then
        IsPlate = True
        Exit Function
    End If
    
    If InStr(upperDesc, "S355JR") > 0 Then
        IsPlate = True
        Exit Function
    End If
    
    IsPlate = False
End Function

Sub LogMessage(msg)
    m_Log = m_Log & Now & " - " & msg & vbCrLf
    WScript.Echo msg
End Sub

Sub SaveLog()
    Dim fso, logFile, logFolder
    Set fso = CreateObject("Scripting.FileSystemObject")

    logFolder = fso.GetParentFolderName(WScript.ScriptFullName) & "\Logs"
    If Not fso.FolderExists(logFolder) Then fso.CreateFolder logFolder

    Dim logPath
    logPath = logFolder & "\Fix_BOM_Plate_Dimensions_Direct_" & Replace(Replace(Replace(Now, "/", "-"), ":", "-"), " ", "_") & ".log"

    Set logFile = fso.CreateTextFile(logPath, True)
    logFile.WriteLine m_Log
    logFile.Close

    WScript.Echo ""
    WScript.Echo "Log saved to: " & logPath
End Sub

Main
