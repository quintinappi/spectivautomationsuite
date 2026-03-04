' EMERGENCY SCRIPT - Revert Sheet Metal to Normal Part
' This converts a sheet metal part back to a standard part
' Author: Quintin de Bruin © 2026

Option Explicit

Const kSheetMetalSubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
Const kStandardPartSubType = "{4D29B490-49B2-11D0-93C3-7E0706000000}"

Sub Main()
    On Error Resume Next

    WScript.Echo "=== REVERT SHEET METAL TO NORMAL PART ==="
    WScript.Echo ""

    Dim invApp
    Set invApp = GetObject(, "Inventor.Application")

    If invApp Is Nothing Then
        WScript.Echo "ERROR: Inventor is not running"
        Exit Sub
    End If

    Dim partDoc
    Set partDoc = invApp.ActiveDocument
    
    If partDoc Is Nothing Then
        WScript.Echo "ERROR: No active document"
        Exit Sub
    End If
    
    WScript.Echo "Part: " & partDoc.DisplayName
    WScript.Echo "Current SubType: " & partDoc.SubType
    
    If partDoc.SubType = kStandardPartSubType Then
        WScript.Echo "Part is already a STANDARD part"
        Exit Sub
    End If
    
    If partDoc.SubType <> kSheetMetalSubType Then
        WScript.Echo "Part is not sheet metal type - unknown type"
        Exit Sub
    End If
    
    WScript.Echo "Part is SHEET METAL type"
    WScript.Echo ""

    Dim smDef
    Set smDef = partDoc.ComponentDefinition

    ' Step 1: Remove flat pattern if exists
    WScript.Echo "=== STEP 1: REMOVE FLAT PATTERN ==="
    
    If smDef.HasFlatPattern Then
        WScript.Echo "Deleting flat pattern..."
        smDef.FlatPattern.Delete
        
        If Err.Number = 0 Then
            WScript.Echo "Deleted"
        Else
            WScript.Echo "Delete failed: " & Err.Description
            Err.Clear
        End If
    Else
        WScript.Echo "No flat pattern to remove"
    End If

    ' Step 2: Try to convert back to standard part
    WScript.Echo ""
    WScript.Echo "=== STEP 2: CONVERT TO STANDARD PART ==="
    
    ' Method 1: Try using ConvertToStandardPart method
    WScript.Echo "Method 1: Trying smDef.ConvertToStandardPart..."
    smDef.ConvertToStandardPart
    
    If Err.Number = 0 Then
        WScript.Echo "ConvertToStandardPart succeeded!"
    Else
        WScript.Echo "Failed: " & Err.Description
        Err.Clear
        
        ' Method 2: Try via command
        WScript.Echo ""
        WScript.Echo "Method 2: Looking for conversion command..."
        
        Dim cmdMgr
        Set cmdMgr = invApp.CommandManager
        
        ' Search for relevant commands
        Dim ctrlDefs
        Set ctrlDefs = cmdMgr.ControlDefinitions
        
        Dim i
        For i = 1 To ctrlDefs.Count
            Dim ctrl
            Set ctrl = ctrlDefs.Item(i)
            
            Dim internalName
            internalName = LCase(ctrl.InternalName)
            
            If (InStr(internalName, "convert") > 0 And InStr(internalName, "standard") > 0) Or _
               (InStr(internalName, "revert") > 0) Then
                WScript.Echo "  Found: " & ctrl.InternalName
                If ctrl.Enabled Then
                    WScript.Echo "    -> ENABLED"
                End If
            End If
            Err.Clear
        Next
        
        ' Try specific command names
        Dim cmdNames
        cmdNames = Array("PartConvertToStandardPartCmd", "ConvertToStandardPartCmd", _
                         "SheetMetalConvertToStandardCmd", "SMConvertToStandardCmd", _
                         "PartRevertToStandardCmd")
        
        Dim cmdName
        For Each cmdName In cmdNames
            Dim cmd
            Set cmd = Nothing
            
            Set cmd = cmdMgr.ControlDefinitions.Item(cmdName)
            
            If Err.Number = 0 And Not cmd Is Nothing Then
                WScript.Echo ""
                WScript.Echo "Found command: " & cmdName
                WScript.Echo "  Enabled: " & cmd.Enabled
                
                If cmd.Enabled Then
                    WScript.Echo "  Executing..."
                    cmd.Execute
                    
                    If Err.Number = 0 Then
                        WScript.Echo "  Command executed!"
                        WScript.Sleep 1000
                        Exit For
                    Else
                        WScript.Echo "  Failed: " & Err.Description
                        Err.Clear
                    End If
                End If
            End If
            Err.Clear
        Next
    End If

    partDoc.Update

    ' Step 3: Verify
    WScript.Echo ""
    WScript.Echo "=== STEP 3: VERIFY ==="
    WScript.Echo "Final SubType: " & partDoc.SubType
    
    If partDoc.SubType = kStandardPartSubType Then
        WScript.Echo ""
        WScript.Echo "SUCCESS! Part is now a STANDARD part."
    ElseIf partDoc.SubType = kSheetMetalSubType Then
        WScript.Echo ""
        WScript.Echo "Part is still SHEET METAL."
        WScript.Echo ""
        WScript.Echo "Manual steps to revert:"
        WScript.Echo "1. In Inventor, go to 3D Model tab"
        WScript.Echo "2. In the Convert panel, click 'Convert to Standard Part'"
        WScript.Echo "   OR"
        WScript.Echo "1. Right-click on the part in the browser"
        WScript.Echo "2. Look for 'Convert to Standard Part' option"
    End If
    
    WScript.Echo ""
    WScript.Echo "=== DONE ==="
End Sub

Main
