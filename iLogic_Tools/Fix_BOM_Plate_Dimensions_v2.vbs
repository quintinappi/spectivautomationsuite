' Fix BOM Plate Dimensions - Clean Version
' Sets LENGTH and WIDTH custom iProperties on plate parts with sheet metal formulas
' Run on open assembly - finds all plate parts (PL/VRN/S355JR in description)
' Author: Quintin de Bruin © 2026

Option Explicit

Const kAssemblyDocumentObject = 12291
Const kPartDocumentObject = 12290

Dim m_InventorApp
Dim m_Log
Dim m_ProcessedParts

Sub Main()
    On Error Resume Next
    
    m_Log = ""
    Set m_ProcessedParts = CreateObject("Scripting.Dictionary")
    
    LogMsg "=== FIX BOM PLATE DIMENSIONS ==="
    LogMsg "Date: " & Now
    LogMsg ""

    Set m_InventorApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Or m_InventorApp Is Nothing Then
        MsgBox "Inventor is not running!", vbCritical
        Exit Sub
    End If

    If m_InventorApp.ActiveDocument Is Nothing Then
        MsgBox "No active document!", vbCritical
        Exit Sub
    End If

    If m_InventorApp.ActiveDocument.DocumentType <> kAssemblyDocumentObject Then
        MsgBox "Please open an assembly!", vbCritical
        Exit Sub
    End If

    Dim asmDoc
    Set asmDoc = m_InventorApp.ActiveDocument
    LogMsg "Assembly: " & asmDoc.DisplayName
    LogMsg ""

    ' Scan and process all plate parts
    LogMsg "=== SCANNING FOR PLATE PARTS ==="
    ScanOccurrences asmDoc.ComponentDefinition.Occurrences
    
    ' Force update
    LogMsg ""
    LogMsg "=== UPDATING ASSEMBLY ==="
    asmDoc.Update
    asmDoc.Rebuild
    LogMsg "Assembly updated"
    
    LogMsg ""
    LogMsg "=== COMPLETE ==="
    LogMsg "Plates processed: " & m_ProcessedParts.Count
    
    MsgBox "Complete!" & vbCrLf & vbCrLf & _
           "Plates processed: " & m_ProcessedParts.Count & vbCrLf & vbCrLf & _
           "NEXT STEPS:" & vbCrLf & _
           "1. In BOM, right-click column header" & vbCrLf & _
           "2. Choose 'Add Custom iProperty Columns'" & vbCrLf & _
           "3. Add LENGTH and WIDTH (Type: Text)", vbInformation

    SaveLog
End Sub

Sub ScanOccurrences(occurrences)
    On Error Resume Next
    
    Dim occ
    For Each occ In occurrences
        If Not occ.Suppressed Then
            Dim refDoc
            Set refDoc = occ.Definition.Document
            
            If Err.Number = 0 And Not refDoc Is Nothing Then
                Dim ext
                ext = LCase(Right(refDoc.FullFileName, 4))
                
                If ext = ".ipt" Then
                    ' Check if already processed
                    If Not m_ProcessedParts.Exists(refDoc.FullFileName) Then
                        ProcessPart refDoc
                    End If
                ElseIf ext = ".iam" Then
                    ' Recurse into sub-assembly
                    ScanOccurrences occ.SubOccurrences
                End If
            End If
            Err.Clear
        End If
    Next
End Sub

Sub ProcessPart(partDoc)
    On Error Resume Next
    
    ' Get description
    Dim description
    description = ""
    description = partDoc.PropertySets.Item("Design Tracking Properties").Item("Description").Value
    If Err.Number <> 0 Then
        Err.Clear
        description = ""
    End If
    
    ' Check if it's a plate
    Dim upperDesc
    upperDesc = UCase(description)
    
    Dim isPlate
    isPlate = False
    If InStr(upperDesc, "PL") > 0 Then isPlate = True
    If InStr(upperDesc, "VRN") > 0 Then isPlate = True
    If InStr(upperDesc, "S355JR") > 0 Then isPlate = True
    
    If Not isPlate Then Exit Sub
    
    ' Check if sheet metal
    Dim isSheetMetal
    isSheetMetal = (partDoc.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
    
    If Not isSheetMetal Then
        LogMsg "SKIP (not sheet metal): " & partDoc.DisplayName
        Exit Sub
    End If
    
    LogMsg "Processing: " & partDoc.DisplayName & " (" & description & ")"
    
    ' Get flat pattern for dimensions
    Dim flatPattern
    Set flatPattern = partDoc.ComponentDefinition.FlatPattern
    If Err.Number <> 0 Or flatPattern Is Nothing Then
        LogMsg "  SKIP: No flat pattern or not sheet metal"
        Err.Clear
        Exit Sub
    End If
    
    ' Get dimensions
    Dim lengthVal, widthVal
    lengthVal = flatPattern.Length
    widthVal = flatPattern.Width
    
    ' Get units
    Dim units
    units = partDoc.UnitsOfMeasure.LengthUnitsAbbreviation
    lengthVal = CStr(Round(lengthVal, 2)) & " " & units
    widthVal = CStr(Round(widthVal, 2)) & " " & units
    
    LogMsg "  Dimensions: " & lengthVal & " x " & widthVal
    
    ' Get custom properties
    Dim customProps
    Set customProps = partDoc.PropertySets.Item("Inventor User Defined Properties")
    If customProps Is Nothing Then
        LogMsg "  ERROR: Cannot access custom properties"
        Exit Sub
    End If
    
    ' Clean up malformed properties
    LogMsg "  Cleaning up old properties..."
    CleanupMalformedProperties customProps
    
    ' Now set LENGTH and WIDTH with actual values
    SetCustomProperty customProps, "LENGTH", lengthVal
    SetCustomProperty customProps, "WIDTH", widthVal
    
    ' Save the part
    partDoc.Save
    If Err.Number <> 0 Then
        LogMsg "  WARNING: Save failed - " & Err.Description
        Err.Clear
    Else
        LogMsg "  Saved"
    End If
    
    ' Mark as processed
    m_ProcessedParts.Add partDoc.FullFileName, True
End Sub

Sub SetCustomProperty(customProps, propName, propValue)
    On Error Resume Next
    
    Dim prop
    Set prop = customProps.Item(propName)
    
    If Err.Number <> 0 Then
        ' Property doesn't exist - create it
        Err.Clear
        Set prop = customProps.Add(propName, propValue)
        If Err.Number <> 0 Then
            LogMsg "  ERROR creating " & propName & ": " & Err.Description
            Err.Clear
            Exit Sub
        End If
        LogMsg "  Created " & propName & " = " & propValue
    Else
        ' Property exists - update value
        prop.Value = propValue
        If Err.Number <> 0 Then
            LogMsg "  ERROR updating " & propName & ": " & Err.Description
            Err.Clear
        Else
            LogMsg "  Updated " & propName & " = " & propValue
        End If
    End If
End Sub

Sub DeleteCustomProperty(customProps, propName)
    On Error Resume Next
    
    Dim prop
    Set prop = customProps.Item(propName)
    
    If Err.Number = 0 And Not prop Is Nothing Then
        prop.Delete
        If Err.Number = 0 Then
            LogMsg "    Deleted: " & propName
        End If
    End If
    Err.Clear
End Sub

Sub CleanupMalformedProperties(customProps)
    On Error Resume Next
    
    ' Build list of properties to delete (can't delete while iterating)
    Dim toDelete
    Set toDelete = CreateObject("Scripting.Dictionary")
    
    Dim prop
    For Each prop In customProps
        Dim propName
        propName = prop.Name
        
        ' Delete if name contains formula syntax
        If InStr(propName, "<") > 0 Or InStr(propName, ">") > 0 Then
            If Not toDelete.Exists(propName) Then toDelete.Add propName, True
        End If
        
        ' Delete if name is just a number (like "0") or "0 mm"
        If IsNumeric(propName) Or propName = "0 mm" Then
            If Not toDelete.Exists(propName) Then toDelete.Add propName, True
        End If
        
        ' Delete HEIGHT (old leftover)
        If UCase(propName) = "HEIGHT" Then
            If Not toDelete.Exists(propName) Then toDelete.Add propName, True
        End If
    Next
    
    LogMsg "    Properties to delete: " & toDelete.Count
    
    ' Now delete each one
    Dim key
    For Each key In toDelete.Keys
        LogMsg "    Deleting: " & key
        Set prop = Nothing
        Set prop = customProps.Item(key)
        If Err.Number = 0 And Not prop Is Nothing Then
            prop.Delete
            If Err.Number <> 0 Then
                LogMsg "      FAILED: " & Err.Description
                Err.Clear
            Else
                LogMsg "      OK"
            End If
        Else
            LogMsg "      NOT FOUND"
            Err.Clear
        End If
    Next
End Sub

Sub LogMsg(msg)
    m_Log = m_Log & msg & vbCrLf
    WScript.Echo msg
End Sub

Sub SaveLog()
    On Error Resume Next
    Dim fso, f, logFolder, logPath
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    logFolder = fso.GetParentFolderName(WScript.ScriptFullName) & "\Logs"
    If Not fso.FolderExists(logFolder) Then fso.CreateFolder logFolder
    
    logPath = logFolder & "\Fix_BOM_Plate_" & Replace(Replace(Replace(Now, "/", "-"), ":", "-"), " ", "_") & ".log"
    Set f = fso.CreateTextFile(logPath, True)
    f.WriteLine m_Log
    f.Close
    
    WScript.Echo ""
    WScript.Echo "Log: " & logPath
End Sub

Main
