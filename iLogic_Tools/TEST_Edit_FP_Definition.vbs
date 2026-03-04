' TEST SCRIPT - Edit Flat Pattern Definition
' Author: Quintin de Bruin © 2026

Option Explicit

Const kSheetMetalSubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"

Sub Main()
    On Error Resume Next

    WScript.Echo "=== EDIT FLAT PATTERN DEFINITION ==="
    WScript.Echo ""

    Dim invApp
    Set invApp = GetObject(, "Inventor.Application")

    Dim partDoc
    Set partDoc = invApp.ActiveDocument
    
    WScript.Echo "Part: " & partDoc.DisplayName

    Dim smDef
    Set smDef = partDoc.ComponentDefinition

    If Not smDef.HasFlatPattern Then
        WScript.Echo "No flat pattern exists - creating one first..."
        smDef.Unfold
        WScript.Sleep 500
    End If
    
    Dim fp
    Set fp = smDef.FlatPattern
    
    WScript.Echo "Current dimensions: " & FormatNumber(fp.Length * 10, 1) & " x " & FormatNumber(fp.Width * 10, 1) & " mm"

    ' Find the largest face
    WScript.Echo ""
    WScript.Echo "Finding largest face..."
    
    Dim body
    Set body = smDef.SurfaceBodies.Item(1)
    
    Dim largestFace
    Dim largestArea
    largestArea = 0
    
    Dim i
    For i = 1 To body.Faces.Count
        Dim face
        Set face = body.Faces.Item(i)
        
        Dim faceArea
        faceArea = face.Evaluator.Area
        
        If faceArea > largestArea Then
            Set largestFace = face
            largestArea = faceArea
        End If
    Next
    
    WScript.Echo "Largest face area: " & FormatNumber(largestArea * 100, 0) & " mm²"

    ' Select the flat pattern feature first
    WScript.Echo ""
    WScript.Echo "Selecting flat pattern feature..."
    
    Dim selectSet
    Set selectSet = partDoc.SelectSet
    selectSet.Clear
    selectSet.Select fp
    
    If Err.Number = 0 Then
        WScript.Echo "Flat pattern selected"
    Else
        WScript.Echo "Could not select: " & Err.Description
        Err.Clear
    End If

    ' Search for Edit Flat Pattern Definition command
    WScript.Echo ""
    WScript.Echo "Searching for 'Edit Flat Pattern Definition' command..."
    
    Dim cmdMgr
    Set cmdMgr = invApp.CommandManager
    
    Dim ctrlDefs
    Set ctrlDefs = cmdMgr.ControlDefinitions
    
    Dim foundCmds
    foundCmds = 0
    
    Dim j
    For j = 1 To ctrlDefs.Count
        Dim ctrl
        Set ctrl = ctrlDefs.Item(j)
        
        Dim internalName
        internalName = LCase(ctrl.InternalName)
        
        If (InStr(internalName, "flat") > 0 And InStr(internalName, "edit") > 0) Or _
           (InStr(internalName, "flat") > 0 And InStr(internalName, "def") > 0) Then
            WScript.Echo "  Found: " & ctrl.InternalName
            If ctrl.Enabled Then
                WScript.Echo "    -> ENABLED"
            End If
            foundCmds = foundCmds + 1
        End If
        Err.Clear
        
        If foundCmds >= 15 Then Exit For
    Next
    
    ' Try specific command names
    WScript.Echo ""
    WScript.Echo "=== TRYING MULTIPLE COMMANDS ==="
    
    ' First try PartChangeFlatPatternBaseFaceCmd
    WScript.Echo ""
    WScript.Echo "1) Trying PartChangeFlatPatternBaseFaceCmd..."
    
    Dim changeCmd
    Set changeCmd = Nothing
    On Error Resume Next
    Set changeCmd = cmdMgr.ControlDefinitions.Item("PartChangeFlatPatternBaseFaceCmd")
    
    If Not changeCmd Is Nothing Then
        WScript.Echo "   Found! Enabled: " & changeCmd.Enabled
        
        ' Pre-select the largest face
        selectSet.Clear
        selectSet.Select largestFace
        WScript.Echo "   Largest face pre-selected"
        
        WScript.Echo "   Executing..."
        changeCmd.Execute
        WScript.Echo "   Execute result: " & Err.Description
        Err.Clear
        
        WScript.Sleep 2000
        partDoc.Update
        
        ' Check new dimensions
        WScript.Echo ""
        WScript.Echo "   New dimensions: " & FormatNumber(fp.Length * 10, 1) & " x " & FormatNumber(fp.Width * 10, 1) & " mm"
    Else
        WScript.Echo "   Not found"
        Err.Clear
    End If
    
    ' Try accessing the flat pattern environment first
    WScript.Echo ""
    WScript.Echo "2) Entering Flat Pattern edit mode..."
    
    Dim editFPCmd
    Set editFPCmd = Nothing
    On Error Resume Next
    Set editFPCmd = cmdMgr.ControlDefinitions.Item("PartEditFlatPatternCmd")
    
    If Not editFPCmd Is Nothing And editFPCmd.Enabled Then
        WScript.Echo "   Found PartEditFlatPatternCmd! Executing..."
        editFPCmd.Execute
        WScript.Echo "   Result: " & Err.Description
        Err.Clear
        WScript.Sleep 1000
        
        ' Now try the edit command
        WScript.Echo ""
        WScript.Echo "3) Now trying SheetMetalEditFlatCtxCmd..."
        
        Dim editCmd
        Set editCmd = cmdMgr.ControlDefinitions.Item("SheetMetalEditFlatCtxCmd")
        
        If editCmd.Enabled Then
            selectSet.Clear
            selectSet.Select largestFace
            editCmd.Execute
            WScript.Echo "   Result: " & Err.Description
            Err.Clear
        End If
        
        ' Exit flat pattern mode
        Dim exitCmd
        Set exitCmd = cmdMgr.ControlDefinitions.Item("PartExitFlatPatternCmd")
        If Not exitCmd Is Nothing Then
            exitCmd.Execute
            Err.Clear
        End If
    Else
        WScript.Echo "   PartEditFlatPatternCmd not found or disabled"
        Err.Clear
    End If
    
    ' Final dimensions check
    WScript.Echo ""
    WScript.Echo "=== FINAL DIMENSIONS ==="
    WScript.Echo "Flat pattern: " & FormatNumber(fp.Length * 10, 1) & " x " & FormatNumber(fp.Width * 10, 1) & " mm"
    
    WScript.Echo ""
    WScript.Echo "=== DONE ==="
End Sub

Main
