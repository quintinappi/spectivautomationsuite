' TEST_Manual_Command.vbs
' Use Inventor commands to change flat pattern orientation
Option Explicit
On Error Resume Next

Dim invApp, partDoc, cmdMgr, selectSet

WScript.Echo "=== MANUAL COMMAND APPROACH ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Then
    WScript.Echo "Cannot connect to Inventor"
    WScript.Quit
End If

Set partDoc = invApp.ActiveDocument
WScript.Echo "Part: " & partDoc.DisplayName

Set cmdMgr = invApp.CommandManager
Set selectSet = partDoc.SelectSet

' List all available sheet metal commands
WScript.Echo ""
WScript.Echo "=== AVAILABLE SHEET METAL COMMANDS ==="

Dim cmdNames, cmd, cmdName
cmdNames = Array( _
    "PartChangeFlatPatternBaseFaceCmd", _
    "SheetMetalChangeFlatPatternBaseFaceCmd", _
    "ChangeFlatPatternBaseFaceCmd", _
    "SMChangeFlatPatternBaseFaceCmd", _
    "PartRedefineBaseFaceCmd", _
    "SheetMetalRedefineBaseFaceCmd", _
    "FlatPatternRedefineBaseCmd", _
    "PartFlatPatternFlipCmd", _
    "SheetMetalFlatPatternFlipCmd", _
    "FlatPatternFlipCmd", _
    "PartFlatPatternOrientationCmd", _
    "SheetMetalOrientationCmd", _
    "SMOrientationCmd" _
)

For Each cmdName In cmdNames
    Set cmd = Nothing
    Set cmd = cmdMgr.ControlDefinitions.Item(cmdName)
    
    If Not cmd Is Nothing Then
        WScript.Echo cmdName & ": Enabled=" & cmd.Enabled
    End If
    Err.Clear
Next

' Try to find commands that contain certain keywords
WScript.Echo ""
WScript.Echo "=== SCANNING ALL CONTROL DEFINITIONS ==="
WScript.Echo "Looking for: base, face, flip, orient, flat..."

Dim ctrlDefs, ctrlDef, foundCount
Set ctrlDefs = cmdMgr.ControlDefinitions
foundCount = 0

For Each ctrlDef In ctrlDefs
    Dim name
    name = LCase(ctrlDef.InternalName)
    
    If InStr(name, "flat") > 0 And (InStr(name, "base") > 0 Or InStr(name, "face") > 0 Or InStr(name, "flip") > 0 Or InStr(name, "orient") > 0) Then
        WScript.Echo "  " & ctrlDef.InternalName & " (Enabled: " & ctrlDef.Enabled & ")"
        foundCount = foundCount + 1
    End If
    Err.Clear
Next

WScript.Echo ""
WScript.Echo "Found " & foundCount & " matching commands"

' Also look for "redefine" commands
WScript.Echo ""
WScript.Echo "=== REDEFINE COMMANDS ==="

For Each ctrlDef In ctrlDefs
    name = LCase(ctrlDef.InternalName)
    
    If InStr(name, "redefine") > 0 Or InStr(name, "change") > 0 Then
        If InStr(name, "sheet") > 0 Or InStr(name, "flat") > 0 Or InStr(name, "part") > 0 Then
            WScript.Echo "  " & ctrlDef.InternalName & " (Enabled: " & ctrlDef.Enabled & ")"
        End If
    End If
    Err.Clear
Next

WScript.Echo ""
WScript.Echo "=== DONE ==="
