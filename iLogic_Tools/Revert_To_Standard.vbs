On Error Resume Next

Dim invApp
Set invApp = GetObject(, "Inventor.Application")

If Err.Number <> 0 Then
    WScript.Echo "ERROR: Cannot connect to Inventor"
    WScript.Quit 1
End If

Dim doc
Set doc = invApp.ActiveDocument

If doc Is Nothing Then
    WScript.Echo "ERROR: No active document"
    WScript.Quit 1
End If

WScript.Echo "Active Document: " & doc.DisplayName

If doc.DocumentType <> 12290 Then
    WScript.Echo "ERROR: Not a part document"
    WScript.Quit 1
End If

Dim compDef
Set compDef = doc.ComponentDefinition

If compDef.DocumentSubType.UniqueID <> "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
    WScript.Echo "Part is already a standard part"
    WScript.Quit 0
End If

WScript.Echo "Converting from sheet metal back to standard part..."

' Delete flat pattern first
Dim flatPattern
Set flatPattern = compDef.FlatPattern

If Not flatPattern Is Nothing Then
    flatPattern.Delete
    WScript.Echo "Flat pattern deleted"
End If

' Get sheet metal definition
Dim smDef
Set smDef = compDef.SheetMetalDefinition

' Delete sheet metal definition
smDef.Delete

WScript.Echo "Sheet metal definition deleted"
WScript.Echo "Part is now a standard part"
WScript.Echo ""
WScript.Echo "Ready for conversion test."
