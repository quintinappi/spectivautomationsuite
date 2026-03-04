' =========================================================
' DIAGNOSE THICKNESS EXPORT - Diagnostic Tool
' =========================================================
' This script diagnoses why Thickness export might not be working
' It scans the assembly and provides detailed information about:
' - Which parts are detected as plates
' - What parameters exist in each part
' - Whether Thickness parameters can be found
' =========================================================

Option Explicit

Dim m_InventorApp
Dim m_Log

Main()

Sub Main()
    On Error Resume Next
    
    m_Log = "=== THICKNESS EXPORT DIAGNOSTIC REPORT ===" & vbCrLf
    m_Log = m_Log & "Generated: " & Now & vbCrLf & vbCrLf
    
    WScript.Echo "=== THICKNESS EXPORT DIAGNOSTIC ==="
    WScript.Echo ""
    
    ' Connect to Inventor
    Set m_InventorApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Could not connect to Inventor."
        WScript.Echo "Please make sure Inventor is running."
        WScript.Quit 1
    End If
    
    Dim activeDoc
    Set activeDoc = m_InventorApp.ActiveDocument
    
    If activeDoc Is Nothing Then
        WScript.Echo "ERROR: No active document."
        WScript.Quit 1
    End If
    
    If activeDoc.DocumentType <> 12291 Then ' kAssemblyDocumentObject
        WScript.Echo "ERROR: Not an assembly document"
        WScript.Quit 1
    End If
    
    m_Log = m_Log & "Assembly: " & activeDoc.FullFileName & vbCrLf
    m_Log = m_Log & String(60, "=") & vbCrLf & vbCrLf
    
    WScript.Echo "Assembly: " & activeDoc.DisplayName
    WScript.Echo ""
    
    ' Get all unique parts from BOM
    Dim bom
    Set bom = activeDoc.ComponentDefinition.BOM
    bom.StructuredViewEnabled = True
    
    Dim bomView
    Set bomView = bom.BOMViews.Item("Structured")
    
    Dim uniqueParts
    Set uniqueParts = CreateObject("Scripting.Dictionary")
    
    Dim i
    For i = 1 To bomView.BOMRows.Count
        Err.Clear
        
        Dim bomRow
        Set bomRow = bomView.BOMRows.Item(i)
        
        If Err.Number = 0 And Not bomRow Is Nothing Then
            Dim compDef
            Set compDef = Nothing
            On Error Resume Next
            Set compDef = bomRow.ComponentDefinitions.Item(1)
            On Error GoTo 0
            
            If Not compDef Is Nothing Then
                Dim partDoc
                Set partDoc = Nothing
                On Error Resume Next
                Set partDoc = compDef.Document
                On Error GoTo 0
                
                If Not partDoc Is Nothing Then
                    Dim fullPath
                    fullPath = partDoc.FullFileName
                    
                    If LCase(Right(fullPath, 4)) = ".ipt" And Not uniqueParts.Exists(fullPath) Then
                        uniqueParts.Add fullPath, partDoc
                    End If
                End If
            End If
        End If
    Next
    
    m_Log = m_Log & "Total unique parts found: " & uniqueParts.Count & vbCrLf & vbCrLf
    WScript.Echo "Found " & uniqueParts.Count & " unique part(s)"
    WScript.Echo ""
    WScript.Echo "Analyzing each part..."
    WScript.Echo ""
    
    ' Analyze each part
    Dim partIndex
    partIndex = 1
    
    Dim partPath
    For Each partPath In uniqueParts.Keys
        AnalyzePart partPath, uniqueParts(partPath), partIndex
        partIndex = partIndex + 1
    Next
    
    ' Summary
    m_Log = m_Log & String(60, "=") & vbCrLf
    m_Log = m_Log & "END OF DIAGNOSTIC REPORT" & vbCrLf
    
    ' Save log
    SaveDiagnosticLog
    
    WScript.Echo ""
    WScript.Echo "Diagnostic complete! Check the log file for full details."
    WScript.Echo ""
    
End Sub

Sub AnalyzePart(filePath, partDoc, index)
    On Error Resume Next
    
    Dim fileName
    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)
    
    m_Log = m_Log & "[" & index & "] " & fileName & vbCrLf
    m_Log = m_Log & String(40, "-") & vbCrLf
    
    WScript.Echo "[" & index & "] " & fileName
    
    ' Get part properties
    Dim partNumber, description
    partNumber = ""
    description = ""
    
    On Error Resume Next
    partNumber = partDoc.PropertySets("Design Tracking Properties").Item("Part Number").Value
    description = partDoc.PropertySets("Design Tracking Properties").Item("Description").Value
    On Error GoTo 0
    
    m_Log = m_Log & "  Part Number: " & partNumber & vbCrLf
    m_Log = m_Log & "  Description: " & description & vbCrLf
    
    WScript.Echo "  Part Number: " & partNumber
    WScript.Echo "  Description: " & description
    
    ' Check if it's a plate
    Dim isPlate
    isPlate = IsPlatePart(partNumber, description)
    
    m_Log = m_Log & "  Is Plate: " & isPlate & vbCrLf
    WScript.Echo "  Is Plate: " & isPlate
    
    If Not isPlate Then
        m_Log = m_Log & "  ** SKIPPED: Not detected as plate part" & vbCrLf
        WScript.Echo "  ** SKIPPED: Not detected as plate part"
        m_Log = m_Log & vbCrLf
        WScript.Echo ""
        Exit Sub
    End If
    
    ' Get component definition
    Dim compDef
    Set compDef = partDoc.ComponentDefinition
    
    If compDef Is Nothing Then
        m_Log = m_Log & "  ERROR: No component definition" & vbCrLf
        m_Log = m_Log & vbCrLf
        WScript.Echo "  ERROR: No component definition"
        WScript.Echo ""
        Exit Sub
    End If
    
    ' Get parameters
    Dim params
    Set params = compDef.Parameters
    
    If params Is Nothing Then
        m_Log = m_Log & "  ERROR: No parameters collection" & vbCrLf
        m_Log = m_Log & vbCrLf
        WScript.Echo "  ERROR: No parameters collection"
        WScript.Echo ""
        Exit Sub
    End If
    
    ' Check ModelParameters for Thickness
    Dim modelParams
    Set modelParams = params.ModelParameters
    
    Dim foundInModelParams
    foundInModelParams = False
    
    m_Log = m_Log & "  ModelParameters:" & vbCrLf
    WScript.Echo "  ModelParameters:"
    
    If Not modelParams Is Nothing Then
        Dim i
        For i = 1 To modelParams.Count
            Dim mpName
            mpName = modelParams.Item(i).Name
            
            If UCase(mpName) = "THICKNESS" Then
                foundInModelParams = True
                Dim mpValue
                mpValue = modelParams.Item(i).Value
                m_Log = m_Log & "    -> THICKNESS = " & mpValue & " mm **" & vbCrLf
                WScript.Echo "    -> THICKNESS = " & mpValue & " mm **"
                
                ' Check export status
                CheckExportStatus modelParams.Item(i), "    "
            Else
                m_Log = m_Log & "       " & mpName & vbCrLf
            End If
        Next
    End If
    
    ' Check UserParameters for Thickness
    Dim userParams
    Set userParams = params.UserParameters
    
    Dim foundInUserParams
    foundInUserParams = False
    
    m_Log = m_Log & "  UserParameters:" & vbCrLf
    WScript.Echo "  UserParameters:"
    
    If Not userParams Is Nothing Then
        For i = 1 To userParams.Count
            Dim upName
            upName = userParams.Item(i).Name
            
            If UCase(upName) = "THICKNESS" Then
                foundInUserParams = True
                Dim upValue
                upValue = userParams.Item(i).Value
                m_Log = m_Log & "    -> THICKNESS = " & upValue & " mm **" & vbCrLf
                WScript.Echo "    -> THICKNESS = " & upValue & " mm **"
                
                ' Check export status
                CheckExportStatus userParams.Item(i), "    "
            Else
                m_Log = m_Log & "       " & upName & vbCrLf
            End If
        Next
    End If
    
    ' Summary for this part
    If Not foundInModelParams And Not foundInUserParams Then
        m_Log = m_Log & "  ** WARNING: No Thickness parameter found!" & vbCrLf
        WScript.Echo "  ** WARNING: No Thickness parameter found!"
    End If
    
    m_Log = m_Log & vbCrLf
    WScript.Echo ""
    
End Sub

Sub CheckExportStatus(param, indent)
    On Error Resume Next
    
    ' Check ExposedAsProperty
    Err.Clear
    Dim exposed
    exposed = param.ExposedAsProperty
    
    If Err.Number = 0 Then
        m_Log = m_Log & indent & "    ExposedAsProperty: " & exposed & vbCrLf
        WScript.Echo indent & "    ExposedAsProperty: " & exposed
    Else
        m_Log = m_Log & indent & "    ExposedAsProperty: (not available)" & vbCrLf
    End If
    
    ' Check ExportParameter
    Err.Clear
    Dim exportParam
    exportParam = param.ExportParameter
    
    If Err.Number = 0 Then
        m_Log = m_Log & indent & "    ExportParameter: " & exportParam & vbCrLf
        WScript.Echo indent & "    ExportParameter: " & exportParam
    Else
        m_Log = m_Log & indent & "    ExportParameter: (not available)" & vbCrLf
    End If
    
    ' Check ExportedToSheet
    Err.Clear
    Dim exportedSheet
    exportedSheet = param.ExportedToSheet
    
    If Err.Number = 0 Then
        m_Log = m_Log & indent & "    ExportedToSheet: " & exportedSheet & vbCrLf
        WScript.Echo indent & "    ExportedToSheet: " & exportedSheet
    Else
        m_Log = m_Log & indent & "    ExportedToSheet: (not available)" & vbCrLf
    End If
    
End Sub

Function IsPlatePart(partNumber, description)
    Dim checkString
    checkString = UCase(partNumber & " " & description)
    
    IsPlatePart = False
    
    If InStr(checkString, "PL ") > 0 Then IsPlatePart = True
    If InStr(checkString, "PLATE") > 0 Then IsPlatePart = True
    If InStr(checkString, "S355JR") > 0 Then IsPlatePart = True
    If InStr(checkString, "VRN") > 0 Then IsPlatePart = True
    
End Function

Sub SaveDiagnosticLog()
    On Error Resume Next
    
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim logFolder
    logFolder = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%USERPROFILE%") & "\Documents\Inventor_Logs"
    
    If Not fso.FolderExists(logFolder) Then
        fso.CreateFolder logFolder
    End If
    
    Dim logPath
    logPath = logFolder & "\Thickness_Diagnostic_" & Replace(Replace(Replace(Now, "/", "-"), ":", "-"), " ", "_") & ".log"
    
    Dim logFile
    Set logFile = fso.CreateTextFile(logPath, True)
    logFile.WriteLine m_Log
    logFile.Close
    
    WScript.Echo "Diagnostic log saved to:"
    WScript.Echo logPath
    
End Sub
