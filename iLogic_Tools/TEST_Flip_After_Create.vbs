' TEST_Flip_After_Create.vbs
' Try FlipBaseFace after flat pattern is created
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, smDef, fp

WScript.Echo "=== FLIP BASE FACE TEST ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Then
    WScript.Echo "Cannot connect to Inventor"
    WScript.Quit
End If

Set partDoc = invApp.ActiveDocument
WScript.Echo "Part: " & partDoc.DisplayName

' Get to the flat pattern
Set compDef = partDoc.ComponentDefinition

' Try different ways to access sheet metal
WScript.Echo ""
WScript.Echo "Checking flat pattern..."

' Method 1: Direct HasFlatPattern
WScript.Echo "Method 1 - HasFlatPattern: " & compDef.HasFlatPattern
If Err.Number <> 0 Then
    WScript.Echo "  Error: " & Err.Description
    Err.Clear
End If

' Try to get flat pattern
Set fp = compDef.FlatPattern
If Err.Number <> 0 Then
    WScript.Echo "FlatPattern access error: " & Err.Description
    Err.Clear
End If

If Not fp Is Nothing Then
    WScript.Echo ""
    WScript.Echo "Current flat pattern:"
    WScript.Echo "  Length: " & FormatNumber(fp.Length * 10, 1) & " mm"
    WScript.Echo "  Width: " & FormatNumber(fp.Width * 10, 1) & " mm"
    
    ' Try FlipBaseFace
    WScript.Echo ""
    WScript.Echo "Trying FlipBaseFace..."
    fp.FlipBaseFace
    
    If Err.Number = 0 Then
        WScript.Echo "FlipBaseFace executed!"
        partDoc.Update
        
        WScript.Echo "After flip:"
        WScript.Echo "  Length: " & FormatNumber(fp.Length * 10, 1) & " mm"
        WScript.Echo "  Width: " & FormatNumber(fp.Width * 10, 1) & " mm"
    Else
        WScript.Echo "FlipBaseFace failed: " & Err.Description
        Err.Clear
    End If
    
    ' Try multiple flips
    WScript.Echo ""
    WScript.Echo "Trying 3 consecutive flips..."
    Dim i
    For i = 1 To 3
        fp.FlipBaseFace
        partDoc.Update
        WScript.Sleep 200
        WScript.Echo "Flip " & i & ": " & FormatNumber(fp.Length * 10, 1) & " x " & FormatNumber(fp.Width * 10, 1) & " mm"
        Err.Clear
    Next
Else
    WScript.Echo "No flat pattern object"
End If

' Try via SheetMetalComponentDefinition explicitly
WScript.Echo ""
WScript.Echo "=== Alternative access ==="

Set smDef = compDef
WScript.Echo "SheetMetal type: " & smDef.Type
WScript.Echo "Thickness: " & smDef.Thickness.Value * 10 & " mm"
If Err.Number <> 0 Then
    WScript.Echo "  Error: " & Err.Description
    Err.Clear
End If

WScript.Echo ""
WScript.Echo "=== DONE ==="
