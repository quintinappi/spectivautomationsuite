' Detailed Assembly Inventory - Debug Version
' Lists ALL parts found in the assembly with detailed parameter info
' Author: Quintin de Bruin © 2026

Option Explicit

' Inventor API Constants
Const kAssemblyDocumentObject = 12291

' Global variables
Dim m_InventorApp
Dim m_Log
Dim m_LogPath

Sub Main()
    On Error Resume Next

    ' Initialize logging
    m_Log = ""
    m_LogPath = CreateTempLogFile()
    
    LogMessage "=== DETAILED ASSEMBLY INVENTORY ==="

    ' Get Inventor application
    LogMessage "Attempting to get Inventor application..."
    Set m_InventorApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        MsgBox "Inventor is not running. Please start Inventor first.", vbCritical, "Error"
        Exit Sub
    End If
    
    LogMessage "Connected to Inventor application"

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
        MsgBox "Please open an ASSEMBLY document (.iam file), not a part.", vbExclamation, "Assembly Required"
        SaveLog
        Exit Sub
    End If

    Dim asmDoc
    Set asmDoc = m_InventorApp.ActiveDocument
    LogMessage "Processing assembly: " & asmDoc.FullFileName

    ' Scan assembly for ALL parts
    LogMessage ""
    LogMessage "SCANNING: Looking for ALL parts in assembly..."
    
    Dim allParts
    Set allParts = CreateObject("Scripting.Dictionary")
    
    Dim uniqueParts
    Set uniqueParts = CreateObject("Scripting.Dictionary")
    
    ' Start recursive traversal
    Call ProcessAssemblyForAllParts(asmDoc, uniqueParts, allParts, "ROOT")
    
    ' Display results
    LogMessage ""
    LogMessage "SCAN RESULTS:"
    LogMessage "============="
    LogMessage "Total parts found: " & allParts.Count
    LogMessage ""
    
    If allParts.Count = 0 Then
        LogMessage "No parts found in assembly!"
        MsgBox "No parts found in assembly!", vbInformation, "Scan Complete"
    Else
        Dim partKey
        Dim partIndex
        partIndex = 1
        
        Dim reportText
        reportText = "DETAILED ASSEMBLY INVENTORY:" & vbCrLf & vbCrLf
        
        For Each partKey In allParts.Keys
            Dim partInfo
            Set partInfo = allParts(partKey)
            
            Dim isPlate
            isPlate = partInfo("isPlate")
            
            Dim hasLength
            hasLength = partInfo("hasLength")
            
            Dim plateStatus
            If isPlate Then
                plateStatus = "[PLATE]"
            Else
                plateStatus = "[NON-PLATE]"
            End If
            
            Dim lengthStatus
            If hasLength Then
                lengthStatus = "[HAS LENGTH]"
            Else
                lengthStatus = "[NO LENGTH]"
            End If
            
            Dim logLine
            logLine = partIndex & ". " & partInfo("fileName") & " " & plateStatus & " " & lengthStatus & vbCrLf & _
                     "   Description: " & partInfo("description") & vbCrLf & _
                     "   Path: " & partInfo("fullPath")
            
            LogMessage logLine
            reportText = reportText & logLine & vbCrLf & vbCrLf
            
            partIndex = partIndex + 1
        Next
        
        MsgBox reportText & vbCrLf & _
               "Total: " & allParts.Count & " parts" & vbCrLf & vbCrLf & _
               "See log file for full details: " & m_LogPath, vbInformation, "Inventory Results"
    End If
    
    SaveLog
    
End Sub

Sub ProcessAssemblyForAllParts(asmDoc, uniqueParts, allParts, asmLevel)
    LogMessage "Processing assembly - " & asmDoc.DisplayName & " (Level: " & asmLevel & ")"

    Dim occurrences
    Set occurrences = asmDoc.ComponentDefinition.Occurrences

    LogMessage "  Found " & occurrences.Count & " occurrences in this assembly"

    ' Process each occurrence
    Dim i
    For i = 1 To occurrences.Count
        Dim occ
        Set occ = occurrences.Item(i)

        ' Check suppressed status
        Dim suppressed
        suppressed = occ.Suppressed
        
        If suppressed Then
            LogMessage "  SKIPPING: Suppressed occurrence"
        Else
            Dim doc
            Set doc = occ.Definition.Document

            Dim fileName
            fileName = GetFileNameFromPath(doc.FullFileName)

            Dim fullPath
            fullPath = doc.FullFileName

            ' Check if this is a part file
            If LCase(Right(fileName, 4)) = ".ipt" Then
                LogMessage "  Found part: " & fileName
                
                ' Only process once (by full path)
                If Not uniqueParts.Exists(fullPath) Then
                    uniqueParts.Add fullPath, True

                    ' Get description from iProperty
                    Dim description
                    description = GetDescriptionFromIProperty(doc)

                    If description = "" Then
                        description = "(no description)"
                    End If

                    ' Check if it's a plate part
                    Dim isPlate
                    isPlate = (InStr(1, UCase(description), "PL", vbTextCompare) > 0 Or _
                               InStr(1, UCase(description), "S355JR", vbTextCompare) > 0)

                    ' Check if it has Length parameter
                    Dim hasLength
                    hasLength = HasLengthParameter(doc)
                    
                    ' Store all info
                    Dim partInfo
                    Set partInfo = CreateObject("Scripting.Dictionary")
                    partInfo.Add "fullPath", fullPath
                    partInfo.Add "fileName", fileName
                    partInfo.Add "description", description
                    partInfo.Add "isPlate", isPlate
                    partInfo.Add "hasLength", hasLength
                    
                    allParts.Add fullPath, partInfo
                    
                    LogMessage "    -> Description: " & description & " | Plate: " & isPlate & " | Has Length: " & hasLength
                Else
                    LogMessage "    -> DUPLICATE (already processed)"
                End If

            ElseIf LCase(Right(fileName, 4)) = ".iam" Then
                ' This is a sub-assembly - recurse into it
                LogMessage "  Recursing into sub-assembly - " & fileName
                Call ProcessAssemblyForAllParts(doc, uniqueParts, allParts, asmLevel & ">" & fileName)
            End If
        End If
    Next
End Sub

Function GetFileNameFromPath(fullPath)
    GetFileNameFromPath = Mid(fullPath, InStrRev(fullPath, "\") + 1)
End Function

Function GetDescriptionFromIProperty(doc)
    ' Read Description from Design Tracking Properties
    On Error Resume Next

    Dim propertySet
    Set propertySet = doc.PropertySets.Item("Design Tracking Properties")

    If Err.Number <> 0 Then
        GetDescriptionFromIProperty = ""
        Err.Clear
        Exit Function
    End If

    Dim descriptionProp
    Set descriptionProp = propertySet.Item("Description")

    If Err.Number <> 0 Then
        GetDescriptionFromIProperty = ""
        Err.Clear
        Exit Function
    End If

    GetDescriptionFromIProperty = Trim(descriptionProp.Value)
    Err.Clear
End Function

Function HasLengthParameter(doc)
    ' Check if the part has a Length parameter
    ' Returns True if Length parameter exists, False otherwise
    
    On Error Resume Next
    
    Dim compDef
    Set compDef = doc.ComponentDefinition
    
    If compDef Is Nothing Then
        HasLengthParameter = False
        Exit Function
    End If
    
    Dim params
    Set params = compDef.Parameters.UserParameters
    
    Dim lengthParam
    Set lengthParam = params.Item("Length")
    
    If Err.Number = 0 And Not lengthParam Is Nothing Then
        HasLengthParameter = True
    Else
        HasLengthParameter = False
    End If
    
    Err.Clear
End Function

Sub LogMessage(message)
    m_Log = m_Log & message & vbCrLf
    WScript.Echo message
End Sub

Function CreateTempLogFile()
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim tempPath
    tempPath = fso.GetSpecialFolder(2) ' TemporaryFolder
    
    Dim logFile
    logFile = tempPath & "\AssemblyInventory_" & Format(Now, "yyyyMMdd_HHmmss") & ".txt"
    
    CreateTempLogFile = logFile
End Function

Sub SaveLog()
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim fileHandle
    Set fileHandle = fso.CreateTextFile(m_LogPath, True)
    fileHandle.Write m_Log
    fileHandle.Close
    
    LogMessage ""
    LogMessage "Log saved to: " & m_LogPath
End Sub

Main
