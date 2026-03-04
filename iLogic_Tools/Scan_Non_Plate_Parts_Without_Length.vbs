' Scan Non-Plate Parts Without Length Parameter
' Scans open assembly and lists all non-plate parts that don't have a Length parameter
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
    
    LogMessage "=== NON-PLATE PARTS WITHOUT LENGTH PARAMETER SCANNER ==="

    ' Get Inventor application
    LogMessage "Attempting to get Inventor application..."
    
    ' Try multiple connection methods
    On Error Resume Next
    Set m_InventorApp = GetObject(, "Inventor.Application")
    
    If Err.Number <> 0 Then
        LogMessage "First attempt failed: " & Err.Description & " (Error: " & Err.Number & ")"
        Err.Clear
        
        ' Try alternative ProgID
        LogMessage "Trying alternative ProgID..."
        Set m_InventorApp = CreateObject("Inventor.Application")
        If Err.Number <> 0 Then
            LogMessage "ERROR: Failed to connect to Inventor - " & Err.Description & " (Error: " & Err.Number & ")"
            MsgBox "Failed to connect to Inventor." & vbCrLf & vbCrLf & _
                   "Error Details:" & vbCrLf & _
                   "Error Number: " & Err.Number & vbCrLf & _
                   "Error Description: " & Err.Description & vbCrLf & vbCrLf & _
                   "Please make sure Inventor is running and try again.", vbCritical, "Connection Error"
            SaveLog
            Exit Sub
        End If
    End If
    
    On Error GoTo 0
    
    If m_InventorApp Is Nothing Then
        LogMessage "ERROR: Inventor application object is Nothing"
        MsgBox "Failed to get valid Inventor application object.", vbCritical, "Error"
        SaveLog
        Exit Sub
    End If
    
    LogMessage "Connected to Inventor application successfully"

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

    ' Scan assembly for non-plate parts without Length parameter
    LogMessage ""
    LogMessage "SCANNING: Looking for non-plate parts without Length parameter..."
    
    Dim nonPlateParts
    Set nonPlateParts = CreateObject("Scripting.Dictionary")
    
    Dim uniqueParts
    Set uniqueParts = CreateObject("Scripting.Dictionary")
    
    ' Start recursive traversal
    Call ProcessAssemblyForNonPlates(asmDoc, uniqueParts, nonPlateParts, "ROOT")
    
    ' Display results
    LogMessage ""
    LogMessage "SCAN RESULTS:"
    LogMessage "============="
    
    If nonPlateParts.Count = 0 Then
        LogMessage "No non-plate parts without Length parameter found!"
        MsgBox "Scan Complete!" & vbCrLf & vbCrLf & _
               "Result: No non-plate parts WITHOUT a Length parameter were found." & vbCrLf & vbCrLf & _
               "All parts either:" & vbCrLf & _
               "  • Are plate parts (contain 'PL' or 'S355JR'), OR" & vbCrLf & _
               "  • Have a Length parameter defined", vbInformation, "Scan Complete"
    Else
        LogMessage "Found " & nonPlateParts.Count & " non-plate parts WITHOUT Length parameter:"
        LogMessage ""
        
        Dim partKey
        Dim partIndex
        partIndex = 1
        
        Dim reportText
        reportText = "NON-PLATE PARTS WITHOUT LENGTH PARAMETER:" & vbCrLf & vbCrLf
        
        For Each partKey In nonPlateParts.Keys
            Dim partInfo
            Set partInfo = nonPlateParts(partKey)
            
            Dim logLine
            logLine = partIndex & ". " & partInfo("fileName") & vbCrLf & _
                     "   Description: " & partInfo("description") & vbCrLf & _
                     "   Path: " & partInfo("fullPath")
            
            LogMessage logLine
            reportText = reportText & logLine & vbCrLf & vbCrLf
            
            partIndex = partIndex + 1
        Next
        
        MsgBox reportText & vbCrLf & _
               "Total: " & nonPlateParts.Count & " parts" & vbCrLf & vbCrLf & _
               "See log file for full details: " & m_LogPath, vbInformation, "Scan Results"
    End If
    
    SaveLog
    
End Sub

Sub ProcessAssemblyForNonPlates(asmDoc, uniqueParts, nonPlateParts, asmLevel)
    On Error Resume Next
    
    LogMessage "Processing assembly - " & asmDoc.DisplayName & " (Level: " & asmLevel & ")"

    Dim occurrences
    Set occurrences = asmDoc.ComponentDefinition.Occurrences
    
    If Err.Number <> 0 Then
        LogMessage "  ERROR getting occurrences: " & Err.Description
        Err.Clear
        Exit Sub
    End If

    LogMessage "  Found " & occurrences.Count & " occurrences"

    ' Process each occurrence
    Dim i
    For i = 1 To occurrences.Count
        Err.Clear
        Dim occ
        Set occ = occurrences.Item(i)
        
        If Err.Number = 0 Then
            ' Skip suppressed occurrences
            If Not occ.Suppressed Then
                Err.Clear
                Dim doc
                Set doc = occ.Definition.Document

                Dim fileName
                fileName = GetFileNameFromPath(doc.FullFileName)

                Dim fullPath
                fullPath = doc.FullFileName

                ' Check if this is a part file
                If LCase(Right(fileName, 4)) = ".ipt" Then
                    ' Only process once (by full path)
                    If Not uniqueParts.Exists(fullPath) Then
                        uniqueParts.Add fullPath, True

                        ' Get description from iProperty
                        Dim description
                        description = GetDescriptionFromIProperty(doc)

                        If description = "" Then
                            description = "(no description)"
                        End If

                        ' Check if it's NOT a plate part
                        Dim isPlate
                        isPlate = (InStr(1, UCase(description), "PL", vbTextCompare) > 0 Or _
                                   InStr(1, UCase(description), "S355JR", vbTextCompare) > 0)

                        If Not isPlate Then
                            ' It's a non-plate part - check if it has Length parameter
                            Dim hasLength
                            hasLength = HasLengthParameter(doc)

                            If Not hasLength Then
                                ' This is what we're looking for!
                                LogMessage "  FOUND: " & fileName & " - " & description
                                
                                Dim partInfo
                                Set partInfo = CreateObject("Scripting.Dictionary")
                                partInfo.Add "fullPath", fullPath
                                partInfo.Add "fileName", fileName
                                partInfo.Add "description", description
                                
                                nonPlateParts.Add fullPath, partInfo
                            End If
                        End If
                    End If

                ElseIf LCase(Right(fileName, 4)) = ".iam" Then
                    ' This is a sub-assembly - recurse into it
                    LogMessage "  Recursing into sub-assembly - " & fileName
                    Call ProcessAssemblyForNonPlates(doc, uniqueParts, nonPlateParts, asmLevel & ">" & fileName)
                End If
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
    logFile = tempPath & "\NonPlateScan_" & Format(Now, "yyyyMMdd_HHmmss") & ".txt"
    
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
