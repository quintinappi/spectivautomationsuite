' TEST SCRIPT - Break Derived Part Link (Simple Version)
' Author: Quintin de Bruin © 2026

Option Explicit

Sub Main()
    On Error Resume Next

    WScript.Echo "=== BREAK DERIVED PART LINK ==="
    WScript.Echo ""

    ' Get Inventor application
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

    ' Get component definition and features
    Dim compDef
    Set compDef = partDoc.ComponentDefinition
    
    Dim features
    Set features = compDef.Features

    ' Check for reference features
    Dim refFeatures
    Set refFeatures = features.ReferenceFeatures
    
    WScript.Echo "ReferenceFeatures count: " & refFeatures.Count
    
    If refFeatures.Count = 0 Then
        WScript.Echo "No derived features found - part is already independent"
        Exit Sub
    End If
    
    ' Select the reference feature
    WScript.Echo ""
    WScript.Echo "Selecting reference feature..."
    
    Dim selectSet
    Set selectSet = partDoc.SelectSet
    selectSet.Clear
    
    Dim refFeature
    Set refFeature = refFeatures.Item(1)
    WScript.Echo "Feature name: " & refFeature.Name
    
    selectSet.Select refFeature
    
    If Err.Number <> 0 Then
        WScript.Echo "Could not select feature: " & Err.Description
        Err.Clear
    Else
        WScript.Echo "Feature selected"
    End If
    
    ' Execute Break Link command
    WScript.Echo ""
    WScript.Echo "Executing PartBreakLinkDerivedPartCtxCmd..."
    
    Dim cmdMgr
    Set cmdMgr = invApp.CommandManager
    
    Dim breakLinkCmd
    Set breakLinkCmd = cmdMgr.ControlDefinitions.Item("PartBreakLinkDerivedPartCtxCmd")
    
    If Err.Number <> 0 Or breakLinkCmd Is Nothing Then
        WScript.Echo "Could not find command: " & Err.Description
        Err.Clear
        Exit Sub
    End If
    
    WScript.Echo "Command found. Enabled: " & breakLinkCmd.Enabled
    
    If Not breakLinkCmd.Enabled Then
        WScript.Echo "Command is not enabled!"
        WScript.Echo "Trying BreakLinkCmd instead..."
        
        Set breakLinkCmd = cmdMgr.ControlDefinitions.Item("BreakLinkCmd")
        If Err.Number = 0 And Not breakLinkCmd Is Nothing Then
            WScript.Echo "BreakLinkCmd enabled: " & breakLinkCmd.Enabled
        End If
    End If
    
    If breakLinkCmd.Enabled Then
        WScript.Echo ""
        WScript.Echo "Executing command..."
        
        ' Try to suppress any dialogs
        Dim oldSilent
        oldSilent = invApp.SilentOperation
        invApp.SilentOperation = True
        
        breakLinkCmd.Execute
        
        If Err.Number <> 0 Then
            WScript.Echo "Execution error: " & Err.Description
            Err.Clear
        Else
            WScript.Echo "Command executed!"
        End If
        
        ' Restore silent mode
        invApp.SilentOperation = oldSilent
        
        ' Wait longer for any dialogs/processing
        WScript.Sleep 2000
        partDoc.Update
        
        ' Check result
        WScript.Echo ""
        WScript.Echo "Checking result..."
        
        Dim newRefCount
        On Error Resume Next
        newRefCount = features.ReferenceFeatures.Count
        
        WScript.Echo "ReferenceFeatures count now: " & newRefCount
        
        If newRefCount = 0 Then
            WScript.Echo ""
            WScript.Echo "SUCCESS! Derived link has been broken!"
            WScript.Echo "The part is now independent."
        Else
            WScript.Echo ""
            WScript.Echo "Reference features still exist."
            WScript.Echo ""
            WScript.Echo "Manual steps required:"
            WScript.Echo "1. In the Model tree, expand 'Folded Model'"
            WScript.Echo "2. Right-click on 'Solid4::DM Underpan...'"  
            WScript.Echo "3. Select 'Break Link'"
            WScript.Echo "4. Click Yes to confirm"
            WScript.Echo ""
            WScript.Echo "Then run the flat pattern fix script again."
        End If
    End If
    
    WScript.Echo ""
    WScript.Echo "=== DONE ==="
End Sub

Main
