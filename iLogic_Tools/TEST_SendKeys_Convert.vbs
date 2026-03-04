' TEST_SendKeys_Convert.vbs
' Use SendKeys to confirm the pre-selected face
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, selectSet, cmdMgr, WshShell

Set WshShell = CreateObject("WScript.Shell")

WScript.Echo "=== SENDKEYS CONVERT TEST ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
Set partDoc = invApp.ActiveDocument
Set compDef = partDoc.ComponentDefinition
Set selectSet = partDoc.SelectSet
Set cmdMgr = invApp.CommandManager

WScript.Echo "Part: " & partDoc.DisplayName
WScript.Echo ""

' Step 1: Revert if sheet metal
If partDoc.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
    WScript.Echo "=== REVERTING TO STANDARD ==="
    
    If compDef.HasFlatPattern Then
        compDef.FlatPattern.Delete
        partDoc.Update
    End If
    
    Dim revertCmd
    Set revertCmd = cmdMgr.ControlDefinitions.Item("PartConvertToStandardPartCmd")
    
    If revertCmd.Enabled Then
        revertCmd.Execute
        WScript.Sleep 1500
        partDoc.Update
        Set compDef = partDoc.ComponentDefinition
        WScript.Echo "Reverted"
    End If
    WScript.Echo ""
End If

' Step 2: Find largest face
WScript.Echo "=== FINDING LARGEST FACE ==="

Dim body, faces, face, largestFace, largestArea
largestArea = 0

Set body = compDef.SurfaceBodies.Item(1)
Set faces = body.Faces

For Each face In faces
    Dim area
    area = face.Evaluator.Area * 100
    If Err.Number = 0 And area > largestArea Then
        largestArea = area
        Set largestFace = face
    End If
    Err.Clear
Next

WScript.Echo "Largest face: " & FormatNumber(largestArea, 0) & " mm²"
WScript.Echo ""

' Step 3: Pre-select face
WScript.Echo "=== PRE-SELECTING FACE ==="
selectSet.Clear
WScript.Sleep 500
selectSet.Select largestFace
WScript.Echo "SelectSet.Count: " & selectSet.Count

' Give Inventor time to highlight the selection
invApp.ActiveView.Update
WScript.Sleep 2000

' Step 4: Start convert command
WScript.Echo ""
WScript.Echo "=== STARTING CONVERT COMMAND ==="

Dim convertCmd
Set convertCmd = cmdMgr.ControlDefinitions.Item("PartConvertToSheetMetalCmd")

If Not convertCmd.Enabled Then
    WScript.Echo "Convert command not enabled!"
    WScript.Quit
End If

' Bring Inventor to foreground
invApp.WindowState = 1 ' kNormalWindow
invApp.Visible = True

' Activate the Inventor window
WshShell.AppActivate "Autodesk Inventor"
WScript.Sleep 1000

' Execute the command
WScript.Echo "Executing Convert command..."
convertCmd.Execute

' Wait longer for dialog to appear and process selection
WScript.Echo "Waiting for dialog (5 seconds)..."
WScript.Sleep 5000

' Try sending Enter to confirm the face
WScript.Echo "Sending Enter to confirm face..."
WshShell.SendKeys "{ENTER}"
WScript.Sleep 2000

' Send another Enter for the settings dialog
WScript.Echo "Sending Enter for settings..."
WshShell.SendKeys "{ENTER}"
WScript.Sleep 2000

' Check result
WScript.Echo ""
WScript.Echo "=== CHECKING RESULT ==="

partDoc.Update
Set compDef = partDoc.ComponentDefinition

WScript.Echo "SubType: " & partDoc.SubType

Dim isSheetMetal
isSheetMetal = (partDoc.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
WScript.Echo "Is Sheet Metal: " & isSheetMetal

If isSheetMetal Then
    If Not compDef.HasFlatPattern Then
        WScript.Echo "Creating flat pattern..."
        compDef.Unfold
        partDoc.Update
    End If
    
    ' Add custom properties with formulas
    WScript.Echo ""
    WScript.Echo "=== ADDING CUSTOM PROPERTIES ==="
    Call AddPlateCustomProperties(partDoc)
    
    If compDef.HasFlatPattern Then
        Dim fp
        Set fp = compDef.FlatPattern
        WScript.Echo ""
        WScript.Echo "Flat Pattern: " & FormatNumber(fp.Length * 10, 1) & " x " & FormatNumber(fp.Width * 10, 1) & " mm"
        
        If fp.Length * 10 > 100 And fp.Width * 10 > 100 Then
            WScript.Echo "*** SUCCESS! ***"
        Else
            WScript.Echo "*** WRONG ORIENTATION ***"
        End If
    End If
End If

WScript.Echo ""
WScript.Echo "=== DONE ==="

' Function to add PLATE LENGTH and PLATE WIDTH custom iProperties with formulas
Sub AddPlateCustomProperties(partDoc)
    On Error Resume Next

    WScript.Echo "Adding PLATE LENGTH and PLATE WIDTH formulas..."

    Dim customPropSet
    Set customPropSet = partDoc.PropertySets.Item("Inventor User Defined Properties")

    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Could not get custom property set"
        Err.Clear
        Exit Sub
    End If

    ' Add or update PLATE LENGTH
    Dim lengthProp
    Set lengthProp = customPropSet.Item("PLATE LENGTH")

    If Err.Number <> 0 Then
        Err.Clear
        customPropSet.Add "=<SHEET METAL LENGTH>", "PLATE LENGTH"
        If Err.Number = 0 Then WScript.Echo "  PLATE LENGTH added"
        Err.Clear
    Else
        lengthProp.Value = "=<SHEET METAL LENGTH>"
        If Err.Number = 0 Then WScript.Echo "  PLATE LENGTH updated"
        Err.Clear
    End If

    ' Add or update PLATE WIDTH
    Dim widthProp
    Set widthProp = customPropSet.Item("PLATE WIDTH")

    If Err.Number <> 0 Then
        Err.Clear
        customPropSet.Add "=<SHEET METAL WIDTH>", "PLATE WIDTH"
        If Err.Number = 0 Then WScript.Echo "  PLATE WIDTH added"
        Err.Clear
    Else
        widthProp.Value = "=<SHEET METAL WIDTH>"
        If Err.Number = 0 Then WScript.Echo "  PLATE WIDTH updated"
        Err.Clear
    End If
    
    partDoc.Update
End Sub
