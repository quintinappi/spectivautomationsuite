' TEST_New_Convert_Single_Part.vbs
' Test the new ConvertPartToSheetMetal function on the active part
Option Explicit
On Error Resume Next

Const kSheetMetalSubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
Const kStandardPartSubType = "{4D29B490-49B2-11D0-93C3-7E0706000000}"

WScript.Echo "=== TEST NEW CONVERT FUNCTION ==="
WScript.Echo ""

Dim invApp
Set invApp = GetObject(, "Inventor.Application")

If invApp Is Nothing Then
    WScript.Echo "ERROR: Inventor not running"
    WScript.Quit
End If

Dim partDoc
Set partDoc = invApp.ActiveDocument

If partDoc Is Nothing Then
    WScript.Echo "ERROR: No document open"
    WScript.Quit
End If

WScript.Echo "Part: " & partDoc.DisplayName
WScript.Echo "Initial SubType: " & partDoc.SubType
WScript.Echo ""

' Call the conversion function (same logic as main script)
Call ConvertPartToSheetMetal(partDoc, "6") ' Assume 6mm thickness

WScript.Echo ""
WScript.Echo "=== TEST COMPLETE ==="

' ========== FUNCTION DEFINITION ==========
Sub ConvertPartToSheetMetal(partDoc, thickness)
    On Error Resume Next

    WScript.Echo "Converting to sheet metal with thickness: " & thickness & "mm"

    ' Part must be the active document for command to work
    partDoc.Activate
    
    ' Force view update and wait for activation to complete
    If Not invApp.ActiveView Is Nothing Then
        invApp.ActiveView.Update
    End If
    WScript.Sleep 500
    
    ' Create WshShell for SendKeys
    Dim WshShell
    Set WshShell = CreateObject("WScript.Shell")
    
    Dim compDef
    Set compDef = partDoc.ComponentDefinition
    
    ' Step 1: Revert to standard part if already sheet metal
    If partDoc.SubType = kSheetMetalSubType Then
        WScript.Echo "Part is already sheet metal - reverting to standard first..."
        
        ' Delete flat pattern if exists
        If compDef.HasFlatPattern Then
            compDef.FlatPattern.Delete
            If Err.Number <> 0 Then
                WScript.Echo "Warning: Could not delete flat pattern: " & Err.Description
                Err.Clear
            Else
                partDoc.Update
                WScript.Echo "Flat pattern deleted"
            End If
        End If
        
        ' Execute revert command
        Dim cmdMgr
        Set cmdMgr = invApp.CommandManager
        
        Dim revertCmd
        Set revertCmd = cmdMgr.ControlDefinitions.Item("PartConvertToStandardPartCmd")
        
        If Not revertCmd Is Nothing And revertCmd.Enabled Then
            WScript.Echo "Executing revert to standard part..."
            revertCmd.Execute
            WScript.Sleep 1500
            partDoc.Update
            Set compDef = partDoc.ComponentDefinition
            WScript.Echo "Reverted to standard part"
        Else
            WScript.Echo "Warning: Could not revert - command not available"
        End If
    End If
    
    ' Re-get component definition after possible revert
    Set compDef = partDoc.ComponentDefinition
    
    ' Step 2: Find the largest face for correct flat pattern orientation
    WScript.Echo "Finding largest face for correct flat pattern orientation..."
    
    Dim body, faces, face, largestFace, largestArea
    largestArea = 0
    
    Set body = compDef.SurfaceBodies.Item(1)
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Could not access surface body: " & Err.Description
        Exit Sub
    End If
    
    Set faces = body.Faces
    
    For Each face In faces
        Dim area
        area = face.Evaluator.Area * 100 ' Convert to mm²
        If Err.Number = 0 And area > largestArea Then
            largestArea = area
            Set largestFace = face
        End If
        Err.Clear
    Next
    
    If largestFace Is Nothing Then
        WScript.Echo "ERROR: Could not find largest face"
        Exit Sub
    End If
    
    WScript.Echo "Largest face area: " & FormatNumber(largestArea, 0) & " mm²"
    
    ' Step 3: Pre-select the largest face
    WScript.Echo "Pre-selecting largest face..."
    
    Dim selectSet
    Set selectSet = partDoc.SelectSet
    selectSet.Clear
    selectSet.Select largestFace
    
    If Err.Number <> 0 Then
        WScript.Echo "Warning: Could not pre-select face: " & Err.Description
        Err.Clear
    Else
        WScript.Echo "Face pre-selected (SelectSet.Count = " & selectSet.Count & ")"
    End If
    
    ' Step 4: Execute Convert to Sheet Metal command
    WScript.Echo "Executing Convert to Sheet Metal command..."
    
    Set cmdMgr = invApp.CommandManager
    
    Dim convertCmd
    Set convertCmd = cmdMgr.ControlDefinitions.Item("PartConvertToSheetMetalCmd")
    
    If convertCmd Is Nothing Or Not convertCmd.Enabled Then
        WScript.Echo "ERROR: Convert command not available"
        Exit Sub
    End If
    
    ' Bring Inventor to foreground
    invApp.WindowState = 1 ' kNormalWindow
    invApp.Visible = True
    
    ' Activate the Inventor window
    WshShell.AppActivate "Autodesk Inventor"
    WScript.Sleep 500
    
    ' Execute the command
    WScript.Echo "Executing: " & convertCmd.DisplayName
    convertCmd.Execute
    
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Command execution failed: " & Err.Description
        Exit Sub
    End If
    
    ' Step 5: Use SendKeys to confirm the pre-selected face
    WScript.Echo "Sending Enter to confirm face selection..."
    WScript.Sleep 1000
    
    WshShell.SendKeys "{ENTER}"
    WScript.Sleep 500
    
    ' Send another Enter for the Sheet Metal Defaults dialog
    WScript.Echo "Sending Enter for Sheet Metal Defaults..."
    WshShell.SendKeys "{ENTER}"
    WScript.Sleep 1000
    
    ' Step 6: Verify conversion
    partDoc.Update
    Set compDef = partDoc.ComponentDefinition
    
    If partDoc.SubType = kSheetMetalSubType Then
        WScript.Echo "VERIFIED: Part is now sheet metal type"
        
        ' Re-get component definition as SheetMetalComponentDefinition
        Set compDef = partDoc.ComponentDefinition
        
        ' Step 7: Create flat pattern if not exists
        If Not compDef.HasFlatPattern Then
            WScript.Echo "Creating flat pattern..."
            Err.Clear
            compDef.Unfold
            If Err.Number <> 0 Then
                WScript.Echo "ERROR creating flat pattern: " & Err.Description
                Err.Clear
            Else
                partDoc.Update
                WScript.Echo "Flat pattern created"
            End If
        Else
            WScript.Echo "Flat pattern already exists"
        End If
        
        ' Verify flat pattern dimensions
        If compDef.HasFlatPattern Then
            Dim fp
            Set fp = compDef.FlatPattern
            Dim fpLength, fpWidth
            fpLength = fp.Length * 10
            fpWidth = fp.Width * 10
            
            WScript.Echo "Flat pattern dimensions: " & FormatNumber(fpLength, 1) & " x " & FormatNumber(fpWidth, 1) & " mm"
            
            ' Check if orientation is correct
            If fpLength > 100 And fpWidth > 100 Then
                WScript.Echo "SUCCESS: Flat pattern shows correct orientation (large face)"
            Else
                WScript.Echo "WARNING: Flat pattern may still show edge view - check manually"
            End If
        End If
        
        ' Add PLATE LENGTH and PLATE WIDTH custom iProperties with formulas
        WScript.Echo ""
        WScript.Echo "Adding custom iProperties..."
        Call AddPlateCustomProperties(partDoc)
        
        ' Save the part
        WScript.Echo "Saving part..."
        partDoc.Save
        If Err.Number <> 0 Then
            WScript.Echo "ERROR saving part: " & Err.Description
            Err.Clear
        Else
            WScript.Echo "Part saved successfully"
        End If
    Else
        WScript.Echo "ERROR: Conversion failed - part SubType is: " & partDoc.SubType
    End If
End Sub

' Function to add PLATE LENGTH and PLATE WIDTH custom iProperties with formulas
Sub AddPlateCustomProperties(partDoc)
    On Error Resume Next

    WScript.Echo "Adding PLATE LENGTH and PLATE WIDTH custom iProperties..."

    ' Get custom property set
    Dim customPropSet
    Set customPropSet = partDoc.PropertySets.Item("Inventor User Defined Properties")

    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Could not get custom property set: " & Err.Description
        Err.Clear
        Exit Sub
    End If

    ' Add or update PLATE LENGTH with formula
    Dim lengthProp
    On Error Resume Next
    Set lengthProp = customPropSet.Item("PLATE LENGTH")

    If Err.Number <> 0 Then
        ' Property doesn't exist, add it with formula
        Err.Clear
        WScript.Echo "Adding PLATE LENGTH = =<SHEET METAL LENGTH>"
        customPropSet.Add "=<SHEET METAL LENGTH>", "PLATE LENGTH"

        If Err.Number <> 0 Then
            WScript.Echo "ERROR: Could not add PLATE LENGTH: " & Err.Description
            Err.Clear
        Else
            WScript.Echo "PLATE LENGTH added successfully"
        End If
    Else
        ' Property exists - update with formula
        WScript.Echo "Updating PLATE LENGTH with formula"
        lengthProp.Value = "=<SHEET METAL LENGTH>"
        If Err.Number <> 0 Then
            WScript.Echo "WARNING: Could not update PLATE LENGTH: " & Err.Description
            Err.Clear
        Else
            WScript.Echo "PLATE LENGTH updated successfully"
        End If
    End If
    Err.Clear

    ' Add or update PLATE WIDTH with formula
    Dim widthProp
    On Error Resume Next
    Set widthProp = customPropSet.Item("PLATE WIDTH")

    If Err.Number <> 0 Then
        ' Property doesn't exist, add it with formula
        Err.Clear
        WScript.Echo "Adding PLATE WIDTH = =<SHEET METAL WIDTH>"
        customPropSet.Add "=<SHEET METAL WIDTH>", "PLATE WIDTH"

        If Err.Number <> 0 Then
            WScript.Echo "ERROR: Could not add PLATE WIDTH: " & Err.Description
            Err.Clear
        Else
            WScript.Echo "PLATE WIDTH added successfully"
        End If
    Else
        ' Property exists - update with formula
        WScript.Echo "Updating PLATE WIDTH with formula"
        widthProp.Value = "=<SHEET METAL WIDTH>"
        If Err.Number <> 0 Then
            WScript.Echo "WARNING: Could not update PLATE WIDTH: " & Err.Description
            Err.Clear
        Else
            WScript.Echo "PLATE WIDTH updated successfully"
        End If
    End If
    Err.Clear
End Sub
