' List all Sheet Metal related commands in Inventor
' This will help find the correct command ID for conversion

Option Explicit

Sub Main()
    On Error Resume Next
    
    Dim invApp
    Set invApp = GetObject(, "Inventor.Application")
    
    If invApp Is Nothing Then
        WScript.Echo "ERROR: Inventor is not running"
        Exit Sub
    End If
    
    WScript.Echo "Connected to Inventor"
    WScript.Echo ""
    
    Dim cmdMgr
    Set cmdMgr = invApp.CommandManager
    
    Dim controlDefs
    Set controlDefs = cmdMgr.ControlDefinitions
    
    WScript.Echo "Total control definitions: " & controlDefs.Count
    WScript.Echo ""
    WScript.Echo "Searching for Sheet Metal related commands..."
    WScript.Echo "=============================================="
    WScript.Echo ""
    
    Dim i
    Dim ctrlDef
    Dim foundCount
    foundCount = 0
    
    For i = 1 To controlDefs.Count
        Set ctrlDef = controlDefs.Item(i)
        
        ' Check if the internal name contains sheet metal related keywords
        Dim internalName
        internalName = LCase(ctrlDef.InternalName)
        
        If InStr(internalName, "sheet") > 0 Or _
           InStr(internalName, "metal") > 0 Or _
           InStr(internalName, "convert") > 0 Or _
           InStr(internalName, "flange") > 0 Or _
           InStr(internalName, "unfold") > 0 Or _
           InStr(internalName, "flat") > 0 Then
           
            WScript.Echo "InternalName: " & ctrlDef.InternalName
            WScript.Echo "DisplayName:  " & ctrlDef.DisplayName
            WScript.Echo "DescriptionText: " & ctrlDef.DescriptionText
            WScript.Echo "---"
            foundCount = foundCount + 1
        End If
    Next
    
    WScript.Echo ""
    WScript.Echo "Found " & foundCount & " related commands"
    WScript.Echo ""
    
    ' Also try to specifically get the conversion command
    WScript.Echo "Attempting to get specific commands:"
    WScript.Echo "====================================="
    
    Dim testCmds
    testCmds = Array("PartConvertToSheetMetalCmd", "ConvertToSheetMetalCmd", _
                     "SMConvertCmd", "SheetMetalConvertCmd", _
                     "PartSheetMetalConvertCmd", "SheetMetalPartConvertCmd", _
                     "PartFlatPatternCmd", "SheetMetalUnfoldCmd")
    
    Dim cmdName
    For Each cmdName In testCmds
        Err.Clear
        Dim cmd
        Set cmd = controlDefs.Item(cmdName)
        If Err.Number = 0 And Not cmd Is Nothing Then
            WScript.Echo cmdName & " = FOUND (" & cmd.DisplayName & ")"
        Else
            WScript.Echo cmdName & " = NOT FOUND"
        End If
    Next
    
    WScript.Echo ""
    WScript.Echo "Done!"
End Sub

Main
