' Add Custom iProperty Columns to BOM
' This script adds LENGTH and WIDTH custom iProperty columns to the assembly BOM
' Author: Quintin de Bruin © 2026

Option Explicit

Dim invApp

Sub Main()
    On Error Resume Next
    
    Set invApp = GetObject(, "Inventor.Application")
    If invApp Is Nothing Then
        WScript.Echo "Inventor not running"
        Exit Sub
    End If
    
    If invApp.ActiveDocument Is Nothing Then
        WScript.Echo "No active document"
        Exit Sub
    End If
    
    Dim asmDoc
    Set asmDoc = invApp.ActiveDocument
    
    If asmDoc.DocumentType <> 12291 Then
        WScript.Echo "Not an assembly"
        Exit Sub
    End If
    
    WScript.Echo "=== ADD BOM COLUMNS ==="
    WScript.Echo "Assembly: " & asmDoc.DisplayName
    WScript.Echo ""
    
    Dim bom
    Set bom = asmDoc.ComponentDefinition.BOM
    
    ' Enable structured view
    If Not bom.StructuredViewEnabled Then
        bom.StructuredViewEnabled = True
    End If
    
    Dim bomView
    Set bomView = bom.BOMViews.Item("Structured")
    
    WScript.Echo "=== CURRENT COLUMNS ==="
    Dim col
    For Each col In bomView.BOMPropertyColumns
        WScript.Echo "  " & col.PropertyDefinition.DisplayName
    Next
    WScript.Echo ""
    
    ' Get the property definitions manager
    Dim propDefSet
    Set propDefSet = invApp.PropertyDefinitions
    
    WScript.Echo "=== PROPERTY DEFINITIONS ==="
    WScript.Echo "PropertyDefinitions count: " & propDefSet.Count
    
    ' List property sets available
    Dim propDef
    Dim i
    i = 0
    For Each propDef In propDefSet
        i = i + 1
        If i <= 20 Then
            WScript.Echo "  " & propDef.Name & " (ID: " & propDef.PropertyDefinitionId & ")"
        End If
    Next
    If i > 20 Then WScript.Echo "  ... and " & (i - 20) & " more"
    WScript.Echo ""
    
    ' Try to find or create LENGTH column
    WScript.Echo "=== ADDING LENGTH COLUMN ==="
    
    ' Try various approaches to add a custom property column
    
    ' Approach 1: Try AddCustomProperty on bomView
    WScript.Echo "Trying bomView.AddCustomProperty..."
    bomView.AddCustomProperty "LENGTH"
    If Err.Number <> 0 Then
        WScript.Echo "  Failed: " & Err.Description
        Err.Clear
    Else
        WScript.Echo "  Success!"
    End If
    
    ' Approach 2: Try using PropertyDefinition
    WScript.Echo "Trying to find LENGTH PropertyDefinition..."
    Dim lengthDef
    For Each propDef In propDefSet
        If UCase(propDef.Name) = "LENGTH" Then
            Set lengthDef = propDef
            WScript.Echo "  Found: " & propDef.Name
            Exit For
        End If
    Next
    
    If Not lengthDef Is Nothing Then
        WScript.Echo "Trying bomView.AddPropertyColumn..."
        bomView.AddPropertyColumn lengthDef
        If Err.Number <> 0 Then
            WScript.Echo "  Failed: " & Err.Description
            Err.Clear
        Else
            WScript.Echo "  Success!"
        End If
    End If
    
    ' Approach 3: Try BOMPropertyColumns.Add
    WScript.Echo "Trying bomView.BOMPropertyColumns.Add..."
    Dim newCol
    Set newCol = bomView.BOMPropertyColumns.Add("LENGTH")
    If Err.Number <> 0 Then
        WScript.Echo "  Failed: " & Err.Description
        Err.Clear
    Else
        WScript.Echo "  Success!"
    End If
    
    ' Approach 4: Use CustomPropertySet
    WScript.Echo ""
    WScript.Echo "=== TRYING CustomPropertyColumns ==="
    
    Dim custCols
    Set custCols = bomView.CustomPropertyColumns
    If Err.Number <> 0 Then
        WScript.Echo "CustomPropertyColumns not accessible: " & Err.Description
        Err.Clear
    Else
        WScript.Echo "CustomPropertyColumns count: " & custCols.Count
        Dim custCol
        For Each custCol In custCols
            WScript.Echo "  " & custCol.PropertyName
        Next
        
        ' Try to add LENGTH
        WScript.Echo "Trying custCols.Add..."
        Set newCol = custCols.Add("LENGTH")
        If Err.Number <> 0 Then
            WScript.Echo "  Failed: " & Err.Description
            Err.Clear
        Else
            WScript.Echo "  Success!"
        End If
    End If
    
    WScript.Echo ""
    WScript.Echo "=== FINAL COLUMNS ==="
    For Each col In bomView.BOMPropertyColumns
        WScript.Echo "  " & col.PropertyDefinition.DisplayName
    Next
    
End Sub

Main
