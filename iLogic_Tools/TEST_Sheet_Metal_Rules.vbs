' TEST_Sheet_Metal_Rules.vbs
' Check sheet metal style/rules that might affect base face selection
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, smDef

WScript.Echo "=== SHEET METAL RULES AND STYLES ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
Set partDoc = invApp.ActiveDocument
Set compDef = partDoc.ComponentDefinition
Set smDef = compDef

WScript.Echo "Part: " & partDoc.DisplayName

' Check sheet metal style
WScript.Echo ""
WScript.Echo "=== SHEET METAL STYLE ==="

Dim smStyle
Set smStyle = smDef.SheetMetalStyle
If Not smStyle Is Nothing Then
    WScript.Echo "Style Name: " & smStyle.Name
    WScript.Echo "Style Type: " & TypeName(smStyle)
    
    ' Check properties
    WScript.Echo ""
    WScript.Echo "Style Properties:"
    WScript.Echo "  Thickness: " & smStyle.Thickness * 10 & " mm"
    Err.Clear
    WScript.Echo "  BendRadius: " & smStyle.BendRadius * 10 & " mm"
    Err.Clear
    WScript.Echo "  MinimumRemnant: " & smStyle.MinimumRemnant * 10 & " mm"
    Err.Clear
    
    ' Check for flat pattern related properties
    WScript.Echo ""
    WScript.Echo "  FlatPatternBaseFaceDir: " & smStyle.FlatPatternBaseFaceDirection
    Err.Clear
    WScript.Echo "  FlatPatternPunchDir: " & smStyle.FlatPatternPunchDirection
    Err.Clear
Else
    WScript.Echo "No sheet metal style: " & Err.Description
End If
Err.Clear

' Check UnfoldMethod
WScript.Echo ""
WScript.Echo "=== UNFOLD METHOD ==="

WScript.Echo "smDef.UnfoldMethod: " & smDef.UnfoldMethod
Err.Clear

' Check FlatPatternOrientation
WScript.Echo "smDef.FlatPatternOrientation: " & smDef.FlatPatternOrientation
Err.Clear

' Check FlatPattern properties
WScript.Echo ""
WScript.Echo "=== FLAT PATTERN PROPERTIES ==="

If smDef.HasFlatPattern Then
    Dim fp
    Set fp = smDef.FlatPattern
    
    WScript.Echo "fp.Length: " & fp.Length * 10 & " mm"
    WScript.Echo "fp.Width: " & fp.Width * 10 & " mm"
    WScript.Echo "fp.Area: " & fp.Area * 100 & " mm²"
    Err.Clear
    
    WScript.Echo ""
    WScript.Echo "fp.PunchDirection: " & fp.PunchDirection
    Err.Clear
    WScript.Echo "fp.BaseFaceDirection: " & fp.BaseFaceDirection
    Err.Clear
    
    ' Try changing direction
    WScript.Echo ""
    WScript.Echo "Trying to change BaseFaceDirection..."
    fp.BaseFaceDirection = 1
    WScript.Echo "  Result: " & Err.Description
    Err.Clear
    
    fp.BaseFaceDirection = 0
    WScript.Echo "  Result 0: " & Err.Description
    Err.Clear
    
    ' Check if FlipDirection exists
    WScript.Echo ""
    WScript.Echo "fp.FlipDirection..."
    fp.FlipDirection
    WScript.Echo "  Result: " & Err.Description
    Err.Clear
    
    WScript.Echo "fp.FlipPunchDirection..."
    fp.FlipPunchDirection
    WScript.Echo "  Result: " & Err.Description
    Err.Clear
    
    ' Check orientation
    WScript.Echo ""
    WScript.Echo "fp.Orientation: " & fp.Orientation
    Err.Clear
    
    ' Try changing orientation
    WScript.Echo "Trying to set Orientation..."
    fp.Orientation = 1
    WScript.Echo "  Result: " & Err.Description
    Err.Clear
    
    ' After any changes, check dimensions
    partDoc.Update
    WScript.Echo ""
    WScript.Echo "After changes:"
    WScript.Echo "  Dimensions: " & FormatNumber(fp.Length * 10, 1) & " x " & FormatNumber(fp.Width * 10, 1) & " mm"
End If

' Check for any setting related to face selection
WScript.Echo ""
WScript.Echo "=== DOCUMENT SETTINGS ==="

Dim docSettings
Set docSettings = partDoc.DocumentSettings
If Not docSettings Is Nothing Then
    WScript.Echo "DocumentSettings found"
    
    ' Check for relevant properties
    WScript.Echo "ModelingSettings: " & TypeName(docSettings.ModelingSettings)
    Err.Clear
End If

WScript.Echo ""
WScript.Echo "=== DONE ==="
