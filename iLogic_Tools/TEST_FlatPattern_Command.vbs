' TEST SCRIPT - Create Flat Pattern via Command with Face Selection
' Author: Quintin de Bruin © 2026

Option Explicit

Const kSheetMetalSubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"

Sub Main()
    On Error Resume Next

    WScript.Echo "=== CREATE FLAT PATTERN VIA COMMAND ==="
    WScript.Echo ""

    Dim invApp
    Set invApp = GetObject(, "Inventor.Application")

    If invApp Is Nothing Then
        WScript.Echo "ERROR: Inventor is not running"
        Exit Sub
    End If

    Dim partDoc
    Set partDoc = invApp.ActiveDocument
    
    WScript.Echo "Part: " & partDoc.DisplayName
    
    ' Check if it's sheet metal
    If partDoc.SubType <> kSheetMetalSubType Then
        WScript.Echo "ERROR: Part is not sheet metal type"
        WScript.Echo "SubType: " & partDoc.SubType
        Exit Sub
    End If
    
    WScript.Echo "Part is sheet metal - good"

    Dim smDef
    Set smDef = partDoc.ComponentDefinition

    ' Step 1: Check current state
    WScript.Echo ""
    WScript.Echo "=== CURRENT STATE ==="
    WScript.Echo "HasFlatPattern: " & smDef.HasFlatPattern
    
    If smDef.HasFlatPattern Then
        Dim fp
        Set fp = smDef.FlatPattern
        WScript.Echo "FlatPattern Length: " & fp.Length * 10 & " mm"
        WScript.Echo "FlatPattern Width: " & fp.Width * 10 & " mm"
    End If

    ' Step 2: Find the largest face
    WScript.Echo ""
    WScript.Echo "=== FINDING LARGEST FACE ==="
    
    Dim body
    Set body = smDef.SurfaceBodies.Item(1)
    
    WScript.Echo "Surface bodies count: " & smDef.SurfaceBodies.Count
    WScript.Echo "Faces in body 1: " & body.Faces.Count
    
    Dim largestFace
    Dim largestArea
    largestArea = 0
    Dim largestFaceIndex
    largestFaceIndex = 0
    
    Dim i
    For i = 1 To body.Faces.Count
        Dim face
        Set face = body.Faces.Item(i)
        
        Dim faceArea
        faceArea = face.Evaluator.Area
        
        If faceArea > largestArea Then
            Set largestFace = face
            largestArea = faceArea
            largestFaceIndex = i
        End If
    Next
    
    WScript.Echo "Largest face is #" & largestFaceIndex & " with area " & FormatNumber(largestArea * 100, 0) & " mm²"

    ' Step 3: Delete existing flat pattern if exists
    If smDef.HasFlatPattern Then
        WScript.Echo ""
        WScript.Echo "=== REMOVING EXISTING FLAT PATTERN ==="
        
        ' Try Refold first
        smDef.Refold
        If Err.Number <> 0 Then
            WScript.Echo "Refold failed: " & Err.Description
            Err.Clear
            
            ' Try FlatPattern.Delete
            smDef.FlatPattern.Delete
            If Err.Number <> 0 Then
                WScript.Echo "Delete failed: " & Err.Description
                Err.Clear
            Else
                WScript.Echo "Deleted via FlatPattern.Delete"
            End If
        Else
            WScript.Echo "Refolded successfully"
        End If
        
        partDoc.Update
        WScript.Sleep 500
    End If

    ' Step 4: Select the largest face
    WScript.Echo ""
    WScript.Echo "=== SELECTING LARGEST FACE ==="
    
    Dim selectSet
    Set selectSet = partDoc.SelectSet
    selectSet.Clear
    
    selectSet.Select largestFace
    
    If Err.Number = 0 Then
        WScript.Echo "Largest face selected"
    Else
        WScript.Echo "Could not select face: " & Err.Description
        Err.Clear
    End If

    ' Step 5: Execute Create Flat Pattern command
    WScript.Echo ""
    WScript.Echo "=== EXECUTING CREATE FLAT PATTERN COMMAND ==="
    
    Dim cmdMgr
    Set cmdMgr = invApp.CommandManager
    
    ' Try different command names
    Dim cmdNames
    cmdNames = Array("PartFlatPatternCmd", "SheetMetalFlatPatternCmd", _
                     "SMFlatPatternCmd", "CreateFlatPatternCmd", _
                     "PartCreateFlatPatternCmd", "SheetMetalCreateFlatPatternCmd")
    
    Dim cmdName
    For Each cmdName In cmdNames
        Dim cmd
        Set cmd = Nothing
        
        Set cmd = cmdMgr.ControlDefinitions.Item(cmdName)
        
        If Err.Number = 0 And Not cmd Is Nothing Then
            WScript.Echo "Found: " & cmdName
            WScript.Echo "  Enabled: " & cmd.Enabled
            
            If cmd.Enabled Then
                WScript.Echo "  Executing with face pre-selected..."
                cmd.Execute
                
                If Err.Number = 0 Then
                    WScript.Echo "  Command executed!"
                    WScript.Sleep 1000
                    Exit For
                Else
                    WScript.Echo "  Execute failed: " & Err.Description
                    Err.Clear
                End If
            End If
        End If
        Err.Clear
    Next
    
    partDoc.Update
    WScript.Sleep 500

    ' Step 6: Check result
    WScript.Echo ""
    WScript.Echo "=== RESULT ==="
    
    If smDef.HasFlatPattern Then
        Set fp = smDef.FlatPattern
        Dim fpLength, fpWidth
        fpLength = fp.Length * 10
        fpWidth = fp.Width * 10
        
        WScript.Echo "Flat pattern created!"
        WScript.Echo "Dimensions: " & FormatNumber(fpLength, 2) & " x " & FormatNumber(fpWidth, 2) & " mm"
        
        If fpLength >= 50 And fpWidth >= 50 Then
            WScript.Echo ""
            WScript.Echo "SUCCESS! Orientation looks correct!"
        Else
            WScript.Echo ""
            WScript.Echo "WARNING: One dimension is small - may be edge view"
            WScript.Echo "The pre-selected face may not have been used."
        End If
    Else
        WScript.Echo "No flat pattern was created"
    End If

    WScript.Echo ""
    WScript.Echo "=== DONE ==="
End Sub

Main
