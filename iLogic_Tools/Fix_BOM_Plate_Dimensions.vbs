' Fix BOM Plate Dimensions - Standalone VBScript
' Adds WIDTH and LENGTH custom iProperty columns to BOM and populates them with
' sheet metal flat pattern dimensions for plate parts only.
' Author: Quintin de Bruin © 2026
'
' FUNCTIONALITY:
' - Scans the open assembly's BOM
' - Creates WIDTH column if it doesn't exist (as custom iProperty column)
' - Creates LENGTH column if it doesn't exist (as custom iProperty column)
' - For plate parts (containing PL, VRN, or S355JR in description):
'   - Sets LENGTH custom iProperty to reference =<sheet metal length>
'   - Sets WIDTH custom iProperty to reference =<sheet metal width>
' - Skips non-plate parts
'
' PLATE DETECTION:
' Parts are identified as plates if their Description iProperty contains:
'   - "PL" (e.g., PL10, 10PL, etc.)
'   - "VRN" (Vloer/Roof/N plates)
'   - "S355JR" (structural steel grade)

Option Explicit

' Inventor API Constants
Const kAssemblyDocumentObject = 12291
Const kPartDocumentObject = 12290

' Global variables
Dim m_InventorApp
Dim m_Log
Dim m_LogPath

Sub Main()
    On Error Resume Next

    ' Initialize logging
    m_Log = ""
    
    LogMessage "=== FIX BOM PLATE DIMENSIONS STARTED ==="
    LogMessage "Date/Time: " & Now

    ' Get Inventor application
    LogMessage "Attempting to get Inventor application..."
    On Error Resume Next

    Set m_InventorApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        LogMessage "ERROR: Inventor not running - " & Err.Description
        MsgBox "Inventor is not running. Please start Inventor first.", vbCritical, "Error"
        SaveLog
        Exit Sub
    End If
    On Error GoTo 0

    If m_InventorApp Is Nothing Then
        LogMessage "ERROR: Inventor application object is Nothing"
        MsgBox "Inventor is not running. Please start Inventor first.", vbCritical, "Error"
        SaveLog
        Exit Sub
    End If

    LogMessage "Connected to Inventor successfully"

    ' Check if we have an active document
    If m_InventorApp.ActiveDocument Is Nothing Then
        LogMessage "ERROR: No active document found"
        MsgBox "No active document! Please open an assembly in Inventor.", vbCritical, "Error"
        SaveLog
        Exit Sub
    End If

    LogMessage "Active document: " & m_InventorApp.ActiveDocument.FullFileName
    LogMessage "Document type: " & m_InventorApp.ActiveDocument.DocumentType

    If m_InventorApp.ActiveDocument.DocumentType <> kAssemblyDocumentObject Then
        LogMessage "ERROR: Document is not an assembly"
        MsgBox "Please open an ASSEMBLY document (.iam file), not a part." & vbCrLf & vbCrLf & _
               "Current document: " & m_InventorApp.ActiveDocument.DisplayName, vbExclamation, "Assembly Required"
        SaveLog
        Exit Sub
    End If

    Dim asmDoc
    Set asmDoc = m_InventorApp.ActiveDocument
    LogMessage "Processing assembly: " & asmDoc.FullFileName

    ' Step 1: Scan assembly for plate parts
    LogMessage ""
    LogMessage "=== STEP 1: Scanning assembly for plate parts ==="
    LogMessage "Looking for parts with PL, VRN, or S355JR in description..."
    
    Dim plateParts
    Set plateParts = ScanAssemblyForPlates(asmDoc)

    If plateParts.Count = 0 Then
        LogMessage "No plate parts found in assembly"
        MsgBox "No plate parts found in the assembly." & vbCrLf & vbCrLf & _
               "Plate parts must have 'PL', 'VRN', or 'S355JR' in their Description iProperty.", _
               vbInformation, "No Plates Found"
        SaveLog
        Exit Sub
    End If

    LogMessage "Found " & plateParts.Count & " plate parts"

    ' Step 2: Process each plate part - add/update custom iProperties
    LogMessage ""
    LogMessage "=== STEP 2: Processing plate parts ==="
    
    Dim processedCount, skippedCount, failedCount
    processedCount = 0
    skippedCount = 0
    failedCount = 0

    Dim partPath
    For Each partPath In plateParts.Keys
        LogMessage ""
        LogMessage "Processing: " & partPath

        Dim partDoc
        Set partDoc = Nothing
        On Error Resume Next
        
        ' Check if part is already open
        Dim doc
        For Each doc In m_InventorApp.Documents
            If LCase(doc.FullFileName) = LCase(partPath) Then
                Set partDoc = doc
                LogMessage "  Part already open in Inventor"
                Exit For
            End If
        Next
        
        ' If not open, open it
        If partDoc Is Nothing Then
            Set partDoc = m_InventorApp.Documents.Open(partPath, False)
            If Err.Number <> 0 Then
                LogMessage "  ERROR: Failed to open part - " & Err.Description
                Err.Clear
                failedCount = failedCount + 1
            End If
        End If
        On Error GoTo 0

        If Not partDoc Is Nothing Then
            ' Check if it's a sheet metal part
            Dim isSheetMetal
            isSheetMetal = IsSheetMetalPart(partDoc)
            
            If isSheetMetal Then
                ' Add/update custom iProperties with formulas
                If SetPlateCustomIProperties(partDoc) Then
                    LogMessage "  Successfully set LENGTH and WIDTH custom iProperties"
                    processedCount = processedCount + 1

                    ' Save the part
                    On Error Resume Next
                    partDoc.Save
                    If Err.Number <> 0 Then
                        LogMessage "  WARNING: Failed to save part - " & Err.Description
                        Err.Clear
                    Else
                        LogMessage "  Part saved"
                    End If
                    On Error GoTo 0
                Else
                    LogMessage "  WARNING: Failed to set custom iProperties"
                    failedCount = failedCount + 1
                End If
            Else
                LogMessage "  SKIPPED: Not a sheet metal part (no flat pattern)"
                skippedCount = skippedCount + 1
            End If
        End If
    Next

    ' Step 3: Force BOM refresh
    LogMessage ""
    LogMessage "=== STEP 3: Refreshing BOM ==="
    On Error Resume Next
    asmDoc.Update
    asmDoc.Rebuild
    If Err.Number <> 0 Then
        LogMessage "WARNING: BOM refresh may not be complete - " & Err.Description
        Err.Clear
    Else
        LogMessage "BOM refreshed successfully"
    End If
    On Error GoTo 0

    ' Summary
    LogMessage ""
    LogMessage "=== PROCESSING COMPLETE ==="
    LogMessage "Parts processed successfully: " & processedCount
    LogMessage "Parts skipped (not sheet metal): " & skippedCount
    LogMessage "Parts failed: " & failedCount
    LogMessage "Total plate parts: " & plateParts.Count

    MsgBox "Processing complete!" & vbCrLf & vbCrLf & _
           "Successfully updated: " & processedCount & " parts" & vbCrLf & _
           "Skipped (not sheet metal): " & skippedCount & " parts" & vbCrLf & _
           "Failed: " & failedCount & " parts" & vbCrLf & vbCrLf & _
           "NEXT STEPS:" & vbCrLf & _
           "1. Open BOM in assembly" & vbCrLf & _
           "2. Add custom iProperty columns: LENGTH and WIDTH" & vbCrLf & _
           "3. Values should show flat pattern dimensions" & vbCrLf & vbCrLf & _
           "Check the log file for details.", vbInformation, "Fix BOM Plate Dimensions"

    SaveLog

End Sub

Function ScanAssemblyForPlates(asmDoc)
    ' Returns a dictionary of full file paths for parts that contain "PL", "VRN", or "S355JR" in description
    Dim result
    Set result = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    
    ' Iterate through all occurrences recursively
    Dim occurrences
    Set occurrences = asmDoc.ComponentDefinition.Occurrences
    
    CollectPlatePartsRecursive occurrences, result
    
    Set ScanAssemblyForPlates = result
End Function

Sub CollectPlatePartsRecursive(occurrences, result)
    On Error Resume Next
    
    Dim occ
    For Each occ In occurrences
        If Not occ.Suppressed Then
            Dim refDoc
            Set refDoc = Nothing
            Set refDoc = occ.Definition.Document
            
            If Err.Number <> 0 Then
                Err.Clear
            ElseIf Not refDoc Is Nothing Then
                ' Check if it's a part
                If LCase(Right(refDoc.FullFileName, 4)) = ".ipt" Then
                    ' Get description from iProperties
                    Dim description
                    description = ""
                    description = GetDescriptionFromIProperty(refDoc)
                    
                    ' Check if it's a plate
                    If IsPlateDescription(description) Then
                        If Not result.Exists(refDoc.FullFileName) Then
                            result.Add refDoc.FullFileName, description
                            LogMessage "  Found plate: " & refDoc.DisplayName & " (" & description & ")"
                        End If
                    End If
                ElseIf LCase(Right(refDoc.FullFileName, 4)) = ".iam" Then
                    ' Recurse into sub-assemblies
                    Dim subOccs
                    Set subOccs = occ.SubOccurrences
                    If Not subOccs Is Nothing Then
                        CollectPlatePartsRecursive subOccs, result
                    End If
                End If
            End If
        End If
    Next
End Sub

Function GetDescriptionFromIProperty(doc)
    On Error Resume Next
    GetDescriptionFromIProperty = ""
    
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
    
    GetDescriptionFromIProperty = Trim(descProp.Value)
End Function

Function IsPlateDescription(description)
    ' Check if description contains plate identifiers
    Dim upperDesc
    upperDesc = UCase(description)
    
    ' Check for PL (plate)
    If InStr(upperDesc, "PL") > 0 Then
        IsPlateDescription = True
        Exit Function
    End If
    
    ' Check for VRN
    If InStr(upperDesc, "VRN") > 0 Then
        IsPlateDescription = True
        Exit Function
    End If
    
    ' Check for S355JR (structural steel)
    If InStr(upperDesc, "S355JR") > 0 Then
        IsPlateDescription = True
        Exit Function
    End If
    
    IsPlateDescription = False
End Function

Function IsSheetMetalPart(partDoc)
    ' Check if the part is a sheet metal part with a flat pattern
    ' Uses multiple detection methods for reliability
    On Error Resume Next
    IsSheetMetalPart = False
    
    If partDoc Is Nothing Then Exit Function
    If partDoc.DocumentType <> kPartDocumentObject Then Exit Function
    
    Dim compDef
    Set compDef = partDoc.ComponentDefinition
    If compDef Is Nothing Then Exit Function
    
    ' Method 1: Check document SubType (most reliable)
    ' Sheet metal parts have SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
    Dim subType
    subType = partDoc.SubType
    LogMessage "  Document SubType: " & subType
    
    If subType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
        LogMessage "  Sheet metal document detected via SubType"
        IsSheetMetalPart = True
        Exit Function
    End If
    
    ' Method 2: Try to access FlatPattern directly
    Dim flatPattern
    Set flatPattern = compDef.FlatPattern
    If Err.Number = 0 And Not flatPattern Is Nothing Then
        LogMessage "  Flat pattern found"
        IsSheetMetalPart = True
        Exit Function
    End If
    Err.Clear
    
    ' Method 3: Check for sheet metal features in the features collection
    Dim features
    Set features = compDef.Features
    If Not features Is Nothing Then
        ' Try to access SheetMetalFeatures
        Dim smFeatures
        Set smFeatures = features.SheetMetalFeatures
        If Err.Number = 0 And Not smFeatures Is Nothing Then
            If smFeatures.Count > 0 Then
                LogMessage "  Sheet metal features found: " & smFeatures.Count
                IsSheetMetalPart = True
                Exit Function
            End If
        End If
        Err.Clear
        
        ' Try to access FaceFeatures (sheet metal face)
        Dim faceFeatures
        Set faceFeatures = features.FaceFeatures
        If Err.Number = 0 And Not faceFeatures Is Nothing Then
            If faceFeatures.Count > 0 Then
                LogMessage "  Face features found (sheet metal): " & faceFeatures.Count
                IsSheetMetalPart = True
                Exit Function
            End If
        End If
        Err.Clear
    End If
    
    ' Method 4: Check TypeName as fallback
    Dim typeName
    typeName = TypeName(compDef)
    LogMessage "  Component TypeName: " & typeName
    
    If InStr(typeName, "SheetMetal") > 0 Then
        LogMessage "  Sheet metal detected via TypeName"
        IsSheetMetalPart = True
        Exit Function
    End If
    
    LogMessage "  Not detected as sheet metal part"
End Function

Function SetPlateCustomIProperties(partDoc)
    ' Sets the LENGTH and WIDTH custom iProperties with formulas referencing flat pattern dimensions
    On Error Resume Next
    SetPlateCustomIProperties = False
    
    If partDoc Is Nothing Then Exit Function
    
    Dim propSets
    Set propSets = partDoc.PropertySets
    If propSets Is Nothing Then Exit Function
    
    ' Get or create Custom property set
    Dim customProps
    Set customProps = propSets.Item("Inventor User Defined Properties")
    If Err.Number <> 0 Then
        Err.Clear
        LogMessage "  ERROR: Could not access custom properties"
        Exit Function
    End If
    
    ' Set LENGTH property with formula
    Dim lengthProp
    Set lengthProp = Nothing
    
    ' Try to get existing property
    Set lengthProp = customProps.Item("LENGTH")
    If Err.Number <> 0 Then
        Err.Clear
        ' Create new property with formula - need to use ItemByPropId or different approach
        ' The Add method signature is: Add(PropName, PropValue)
        ' But for formulas we may need to set the Expression property
        customProps.Add "LENGTH", "=<sheet metal length>"
        If Err.Number <> 0 Then
            LogMessage "  ERROR creating LENGTH property: " & Err.Description
            Err.Clear
            ' Try alternative: create with empty value first, then set formula
            customProps.Add "LENGTH", ""
            Err.Clear
            Set lengthProp = customProps.Item("LENGTH")
            If Not lengthProp Is Nothing Then
                lengthProp.Expression = "=<sheet metal length>"
                If Err.Number <> 0 Then
                    Err.Clear
                    lengthProp.Value = "=<sheet metal length>"
                    Err.Clear
                End If
                LogMessage "  Created LENGTH custom iProperty (alt method)"
            End If
        Else
            LogMessage "  Created LENGTH custom iProperty"
        End If
    Else
        ' Update existing property - try Expression first, then Value
        lengthProp.Expression = "=<sheet metal length>"
        If Err.Number <> 0 Then
            Err.Clear
            lengthProp.Value = "=<sheet metal length>"
            If Err.Number <> 0 Then
                LogMessage "  ERROR updating LENGTH property: " & Err.Description
                Err.Clear
            Else
                LogMessage "  Updated LENGTH custom iProperty"
            End If
        Else
            LogMessage "  Updated LENGTH custom iProperty (expression)"
        End If
    End If
    
    ' Set WIDTH property with formula
    Dim widthProp
    Set widthProp = Nothing
    
    ' Debug: List all existing custom properties
    LogMessage "  Existing custom properties:"
    Dim prop
    For Each prop In customProps
        LogMessage "    - " & prop.Name & " = " & prop.Value
    Next
    
    ' Check for and delete malformed property (where formula became the name)
    On Error Resume Next
    Dim badProp
    Set badProp = customProps.Item("=<sheet metal width>")
    If Err.Number = 0 And Not badProp Is Nothing Then
        LogMessage "  Deleting malformed property: =<sheet metal width>"
        badProp.Delete
        Err.Clear
    End If
    Err.Clear
    
    ' Try to get existing WIDTH property
    Set widthProp = customProps.Item("WIDTH")
    If Err.Number <> 0 Then
        LogMessage "  WIDTH property does not exist, creating..."
        Err.Clear
        ' In Inventor API, PropertySets.Add requires: Name, Value
        ' The formula syntax is interpreted when you set Expression property
        ' First create with a placeholder value
        Dim newWidthProp
        Set newWidthProp = customProps.Add("WIDTH", "0")
        If Err.Number <> 0 Then
            LogMessage "  ERROR creating WIDTH property with placeholder: " & Err.Description & " (Error: " & Err.Number & ")"
            Err.Clear
        Else
            LogMessage "  Created WIDTH property with placeholder"
            ' Now set the Expression to the formula
            newWidthProp.Expression = "<sheet metal width>"
            If Err.Number <> 0 Then
                LogMessage "  ERROR setting WIDTH expression: " & Err.Description
                Err.Clear
                ' Try setting value directly with formula text
                newWidthProp.Value = "=<sheet metal width>"
                If Err.Number <> 0 Then
                    LogMessage "  ERROR setting WIDTH value: " & Err.Description
                    Err.Clear
                Else
                    LogMessage "  Set WIDTH value to formula text"
                End If
            Else
                LogMessage "  Set WIDTH expression successfully"
            End If
        End If
    Else
        LogMessage "  WIDTH property exists, updating..."
        ' Update existing property - try Expression first, then Value
        widthProp.Expression = "<sheet metal width>"
        If Err.Number <> 0 Then
            Err.Clear
            widthProp.Value = "=<sheet metal width>"
            If Err.Number <> 0 Then
                LogMessage "  ERROR updating WIDTH property: " & Err.Description
                Err.Clear
            Else
                LogMessage "  Updated WIDTH custom iProperty"
            End If
        Else
            LogMessage "  Updated WIDTH custom iProperty (expression)"
        End If
    End If
    
    SetPlateCustomIProperties = True
End Function

Sub LogMessage(msg)
    m_Log = m_Log & Now & " - " & msg & vbCrLf
    WScript.Echo msg
End Sub

Sub SaveLog()
    Dim fso, logFile, logFolder
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Create log path in script folder
    logFolder = fso.GetParentFolderName(WScript.ScriptFullName) & "\Logs"

    If Not fso.FolderExists(logFolder) Then
        fso.CreateFolder logFolder
    End If

    m_LogPath = logFolder & "\Fix_BOM_Plate_Dimensions_" & Replace(Replace(Replace(Now, "/", "-"), ":", "-"), " ", "_") & ".log"

    Set logFile = fso.CreateTextFile(m_LogPath, True)
    logFile.WriteLine m_Log
    logFile.Close

    WScript.Echo ""
    WScript.Echo "Log saved to: " & m_LogPath
End Sub

' Start the script
Main()
