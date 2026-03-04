' TEST SCRIPT - Break Derived Part Link
' Converts a derived part to an independent part by breaking the link
' Author: Quintin de Bruin © 2026

Option Explicit

' Inventor API Constants
Const kPartDocumentObject = 12290

Sub Main()
    On Error Resume Next

    WScript.Echo "=== BREAK DERIVED PART LINK TEST ==="
    WScript.Echo ""

    ' Get Inventor application
    Dim invApp
    Set invApp = GetObject(, "Inventor.Application")

    If invApp Is Nothing Then
        WScript.Echo "ERROR: Inventor is not running"
        Exit Sub
    End If

    WScript.Echo "Connected to Inventor"

    ' Check if we have an active document
    If invApp.ActiveDocument Is Nothing Then
        WScript.Echo "ERROR: No active document"
        Exit Sub
    End If

    Dim partDoc
    Set partDoc = invApp.ActiveDocument

    ' Verify it's a part document
    If partDoc.DocumentType <> kPartDocumentObject Then
        WScript.Echo "ERROR: Active document is not a part"
        Exit Sub
    End If

    WScript.Echo "Part: " & partDoc.DisplayName
    WScript.Echo ""

    ' Get component definition
    Dim compDef
    Set compDef = partDoc.ComponentDefinition

    ' Look for derived features
    WScript.Echo "=== CHECKING FOR DERIVED FEATURES ==="
    
    Dim features
    Set features = compDef.Features
    
    ' Check for DerivePartFeatures
    Dim derivePartFeatures
    Set derivePartFeatures = Nothing
    
    On Error Resume Next
    Set derivePartFeatures = features.DerivePartFeatures
    
    If Err.Number = 0 And Not derivePartFeatures Is Nothing Then
        WScript.Echo "DerivePartFeatures count: " & derivePartFeatures.Count
        
        If derivePartFeatures.Count > 0 Then
            Dim i
            For i = 1 To derivePartFeatures.Count
                Dim dpf
                Set dpf = derivePartFeatures.Item(i)
                
                WScript.Echo ""
                WScript.Echo "Derived Feature " & i & ":"
                WScript.Echo "  Name: " & dpf.Name
                WScript.Echo "  Linked: " & dpf.Linked
                
                ' Try to get the source file
                Dim sourceFile
                On Error Resume Next
                sourceFile = dpf.ReferencedFile.FullFileName
                If Err.Number = 0 Then
                    WScript.Echo "  Source: " & sourceFile
                End If
                Err.Clear
                
                ' Check if we can break the link
                If dpf.Linked Then
                    WScript.Echo ""
                    WScript.Echo "=== ATTEMPTING TO BREAK LINK ==="
                    
                    ' Method 1: Try BreakLink method
                    WScript.Echo "Method 1: Trying dpf.BreakLink..."
                    dpf.BreakLink
                    
                    If Err.Number = 0 Then
                        WScript.Echo "BreakLink succeeded!"
                    Else
                        WScript.Echo "BreakLink failed: " & Err.Description
                        Err.Clear
                        
                        ' Method 2: Try setting Linked = False
                        WScript.Echo "Method 2: Trying dpf.Linked = False..."
                        dpf.Linked = False
                        
                        If Err.Number = 0 Then
                            WScript.Echo "Setting Linked = False succeeded!"
                        Else
                            WScript.Echo "Setting Linked failed: " & Err.Description
                            Err.Clear
                        End If
                    End If
                    
                    ' Verify
                    partDoc.Update
                    WScript.Echo ""
                    WScript.Echo "After break attempt:"
                    WScript.Echo "  Linked: " & dpf.Linked
                Else
                    WScript.Echo "Link is already broken"
                End If
            Next
        Else
            WScript.Echo "No DerivePartFeatures found"
        End If
    Else
        WScript.Echo "Could not access DerivePartFeatures: " & Err.Description
        Err.Clear
    End If
    
    ' Also check ReferenceFeatures
    WScript.Echo ""
    WScript.Echo "=== CHECKING REFERENCE FEATURES ==="
    
    Dim refFeatures
    Set refFeatures = Nothing
    
    On Error Resume Next
    Set refFeatures = features.ReferenceFeatures
    
    If Err.Number = 0 And Not refFeatures Is Nothing Then
        WScript.Echo "ReferenceFeatures count: " & refFeatures.Count
        
        Dim j
        For j = 1 To refFeatures.Count
            Dim rf
            Set rf = refFeatures.Item(j)
            WScript.Echo ""
            WScript.Echo "Reference Feature " & j & ":"
            WScript.Echo "  Name: " & rf.Name
            
            ' Check if it's linked
            Dim isLinked
            isLinked = False
            
            On Error Resume Next
            isLinked = rf.Linked
            If Err.Number <> 0 Then
                Err.Clear
                ' Try LinkedFile property
                Dim linkedFile
                Set linkedFile = rf.LinkedFile
                If Err.Number = 0 And Not linkedFile Is Nothing Then
                    isLinked = True
                    WScript.Echo "  Linked to: " & linkedFile.FullFileName
                End If
                Err.Clear
            Else
                WScript.Echo "  Linked: " & isLinked
            End If
            
            ' Try to break the link
            WScript.Echo ""
            WScript.Echo "  Attempting to break link..."
            
            ' Method 1: BreakLink
            rf.BreakLink
            If Err.Number = 0 Then
                WScript.Echo "  BreakLink() succeeded!"
            Else
                WScript.Echo "  BreakLink() failed: " & Err.Description
                Err.Clear
                
                ' Method 2: BreakLinkToFile
                rf.BreakLinkToFile
                If Err.Number = 0 Then
                    WScript.Echo "  BreakLinkToFile() succeeded!"
                Else
                    WScript.Echo "  BreakLinkToFile() failed: " & Err.Description
                    Err.Clear
                    
                    ' Method 3: Set Linked = False
                    rf.Linked = False
                    If Err.Number = 0 Then
                        WScript.Echo "  Setting Linked=False succeeded!"
                    Else
                        WScript.Echo "  Setting Linked=False failed: " & Err.Description
                        Err.Clear
                    End If
                End If
            End If
        Next
    Else
        WScript.Echo "Could not access ReferenceFeatures: " & Err.Description
        Err.Clear
    End If
    
    ' Try using command to break link
    WScript.Echo ""
    WScript.Echo "=== TRYING BREAK LINK COMMAND ==="
    
    Dim cmdMgr
    Set cmdMgr = invApp.CommandManager
    
    ' First, let's search for any command containing "break" or "link"
    WScript.Echo "Searching for break/link related commands..."
    
    Dim ctrlDefs
    Set ctrlDefs = cmdMgr.ControlDefinitions
    
    Dim foundCmds
    foundCmds = 0
    
    Dim k
    For k = 1 To ctrlDefs.Count
        Dim ctrl
        Set ctrl = ctrlDefs.Item(k)
        
        On Error Resume Next
        Dim internalName
        internalName = LCase(ctrl.InternalName)
        
        If InStr(internalName, "break") > 0 Or InStr(internalName, "derive") > 0 Then
            WScript.Echo "  Found: " & ctrl.InternalName
            If ctrl.Enabled Then
                WScript.Echo "    -> ENABLED"
            End If
            foundCmds = foundCmds + 1
            
            If foundCmds >= 20 Then Exit For ' Limit output
        End If
        Err.Clear
    Next
    
    ' Try specific command names for breaking links
    WScript.Echo ""
    WScript.Echo "Trying specific break link commands..."
    
    Dim cmdNames
    cmdNames = Array("PartBreakLinkDerivedPartCtxCmd", "PartBreakLinkPromoteCtxCmd", _
                     "BreakLinkCmd", "PartBreakLinkCmd", "DerivedBreakLinkCmd")
    
    Dim cmdName
    For Each cmdName In cmdNames
        Dim cmd
        Set cmd = Nothing
        
        On Error Resume Next
        Set cmd = cmdMgr.ControlDefinitions.Item(cmdName)
        
        If Err.Number = 0 And Not cmd Is Nothing Then
            WScript.Echo ""
            WScript.Echo "Found command: " & cmdName
            WScript.Echo "  Display Name: " & cmd.DisplayName
            WScript.Echo "  Enabled: " & cmd.Enabled
            
            If cmd.Enabled Then
                ' First, select the reference feature so the context command knows what to break
                WScript.Echo "  Selecting reference feature first..."
                
                Dim selectSet
                Set selectSet = partDoc.SelectSet
                selectSet.Clear
                
                ' Select the first reference feature
                Dim firstRef
                Set firstRef = features.ReferenceFeatures.Item(1)
                selectSet.Select firstRef
                
                If Err.Number = 0 Then
                    WScript.Echo "  Reference feature selected"
                Else
                    WScript.Echo "  Could not select: " & Err.Description
                    Err.Clear
                End If
                
                WScript.Echo "  Executing command..."
                cmd.Execute
                WScript.Sleep 1000
                
                If Err.Number = 0 Then
                    WScript.Echo "  Command executed successfully!"
                    partDoc.Update
                    
                    ' Check if reference feature is still there
                    WScript.Echo ""
                    WScript.Echo "  Checking result..."
                    WScript.Echo "  ReferenceFeatures count now: " & features.ReferenceFeatures.Count
                    
                    Exit For ' Stop trying other commands
                Else
                    WScript.Echo "  Execution failed: " & Err.Description
                    Err.Clear
                End If
            End If
        End If
        Err.Clear
    Next
    
    WScript.Echo ""
    WScript.Echo "=== TEST COMPLETE ==="
    WScript.Echo ""
    WScript.Echo "Check if the part is now independent (no derive link)."
    WScript.Echo "If successful, you can now fix the flat pattern orientation."
    
End Sub

' Run main
Main
