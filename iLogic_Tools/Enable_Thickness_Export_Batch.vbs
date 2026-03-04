' =========================================================
' ENABLE THICKNESS EXPORT - ALL PLATE PARTS IN ASSEMBLY
' =========================================================
' For each plate part: opens part, enables Thickness export,
' saves, then returns to assembly
' =========================================================

Option Explicit

Dim m_InventorApp
Dim m_Shell
Dim m_LogFile
Dim m_LogPath

Main()

Sub Main()
    On Error Resume Next
    
    WScript.Echo "=== ENABLE THICKNESS EXPORT - ASSEMBLY BATCH ==="
    WScript.Echo ""
    
    Set m_InventorApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Could not connect to Inventor."
        WScript.Quit 1
    End If
    
    Dim activeDoc
    Set activeDoc = m_InventorApp.ActiveDocument
    
    If activeDoc Is Nothing Then
        WScript.Echo "ERROR: No active document."
        WScript.Quit 1
    End If
    
    ' Check if it's an assembly
    If activeDoc.DocumentType <> 12291 Then ' kAssemblyDocumentObject
        WScript.Echo "ERROR: Not an assembly"
        WScript.Quit 1
    End If
    
    WScript.Echo "Assembly: " & activeDoc.DisplayName
    WScript.Echo ""
    
    ' Initialize Shell for SendKeys
    Set m_Shell = CreateObject("WScript.Shell")
    
    ' Setup log file
    m_LogPath = m_InventorApp.FileManager.GetDefaultProjectPath
    If m_LogPath = "" Then
        m_LogPath = m_InventorApp.DesignProjectManager.ActiveDesignProject.WorkspacePath
    End If
    If m_LogPath = "" Then
        m_LogPath = CreateObject("WScript.Shell").SpecialFolders("Documents")
    End If
    
    ' Get BOM to find plate parts
    Dim bom
    Set bom = activeDoc.ComponentDefinition.BOM
    bom.StructuredViewEnabled = True
    bom.StructuredViewFirstLevelOnly = False
    
    Dim bomView
    Set bomView = bom.BOMViews.Item("Structured")
    
    WScript.Echo "Scanning BOM for plate parts..."
    WScript.Echo ""
    
    ' Collect all unique plate part documents
    Dim plateParts()
    Dim platePartCount
    platePartCount = 0
    
    Dim i
    For i = 1 To bomView.BOMRows.Count
        Dim bomRow
        Set bomRow = bomView.BOMRows.Item(i)
        
        Dim compDef
        Set compDef = bomRow.ComponentDefinitions.Item(1)
        Dim partDoc
        Set partDoc = compDef.Document
        
        Dim partNum
        partNum = ""
        On Error Resume Next
        partNum = partDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value
        On Error Goto 0
        
        If InStr(1, partNum, "PL", vbTextCompare) > 0 Then
            ' Check if already in list
            Dim j
            Dim isNewPart
            isNewPart = True
            
            For j = 0 To platePartCount - 1
                If plateParts(j).FullDocumentName = partDoc.FullDocumentName Then
                    isNewPart = False
                    Exit For
                End If
            Next
            
            If isNewPart Then
                If platePartCount = 0 Then
                    ReDim plateParts(0)
                Else
                    ReDim Preserve plateParts(platePartCount)
                End If
                Set plateParts(platePartCount) = partDoc
                platePartCount = platePartCount + 1
            End If
        End If
    Next
    
    If platePartCount = 0 Then
        WScript.Echo "No plate parts found."
        WScript.Quit 0
    End If
    
    WScript.Echo "Found " & platePartCount & " unique plate part(s)"
    WScript.Echo ""
    WScript.Echo "Processing parts..."
    WScript.Echo ""
    
    ' Process each plate part
    Dim processedCount
    processedCount = 0
    
    For i = 0 To platePartCount - 1
        Dim plateDoc
        Set plateDoc = plateParts(i)
        
        Dim plateName
        plateName = ""
        On Error Resume Next
        plateName = plateDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value
        On Error Goto 0
        
        WScript.Echo "  [" & (i + 1) & "/" & platePartCount & "] " & plateName & "..."
        
        ' Open the part if not already open
        Dim isAlreadyOpen
        isAlreadyOpen = True
        
        On Error Resume Next
        Dim checkDoc
        Set checkDoc = m_InventorApp.Documents.Item(plateDoc.FullDocumentName)
        If Err.Number <> 0 Then
            isAlreadyOpen = False
        End If
        Err.Clear
        On Error Goto 0
        
        If Not isAlreadyOpen Then
            WScript.Echo "    Opening part..."
            m_InventorApp.Documents.Open plateDoc.FullDocumentName, False
            WScript.Sleep 2000
        End If
        
        ' Get the open part document
        Set plateDoc = m_InventorApp.ActiveDocument
        
        ' Check if it has Thickness parameter
        Dim params
        Set params = plateDoc.ComponentDefinition.Parameters
        
        Dim thicknessParam
        Set thicknessParam = Nothing
        
        On Error Resume Next
        Set thicknessParam = params.Item("Thickness")
        On Error Goto 0
        
        If thicknessParam Is Nothing Then
            WScript.Echo "    WARNING: No Thickness parameter found, skipping"
        Else
            ' Enable export via UI automation
            WScript.Echo "    Enabling Thickness export..."
            EnableThicknessExport
            
            ' Save
            plateDoc.Save
            WScript.Echo "    Saved"
            processedCount = processedCount + 1
        End If
        
        WScript.Echo ""
    Next
    
    WScript.Echo "================================"
    WScript.Echo "Completed: " & processedCount & " part(s) updated"
    WScript.Echo "================================"
    WScript.Echo ""
    
End Sub

Sub EnableThicknessExport()
    On Error Resume Next
    
    ' Open Parameters dialog: Tools > Parameters
    m_Shell.SendKeys "%t"
    WScript.Sleep 800
    
    m_Shell.SendKeys "p"
    WScript.Sleep 1500
    
    ' Search for Thickness
    m_Shell.SendKeys "t"
    WScript.Sleep 200
    m_Shell.SendKeys "h"
    WScript.Sleep 300
    
    ' Tab to Export checkbox
    m_Shell.SendKeys "{TAB}"
    WScript.Sleep 200
    m_Shell.SendKeys "{TAB}"
    WScript.Sleep 200
    
    ' Check the Export checkbox
    m_Shell.SendKeys " "
    WScript.Sleep 300
    
    ' Click OK
    m_Shell.SendKeys "{ENTER}"
    WScript.Sleep 1000
    
End Sub
