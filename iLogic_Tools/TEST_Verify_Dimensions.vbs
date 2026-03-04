' TEST_Verify_Dimensions.vbs
' Verify the flat pattern dimensions after SendKeys conversion
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef

WScript.Echo "=== VERIFY FLAT PATTERN DIMENSIONS ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
Set partDoc = invApp.ActiveDocument
Set compDef = partDoc.ComponentDefinition

WScript.Echo "Part: " & partDoc.DisplayName
WScript.Echo "SubType: " & partDoc.SubType
WScript.Echo "Is Sheet Metal: " & (partDoc.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
WScript.Echo ""

If compDef.HasFlatPattern Then
    Dim fp
    Set fp = compDef.FlatPattern
    
    WScript.Echo "=== FLAT PATTERN DETAILS ==="
    WScript.Echo "Length: " & FormatNumber(fp.Length * 10, 1) & " mm"
    WScript.Echo "Width: " & FormatNumber(fp.Width * 10, 1) & " mm"
    WScript.Echo "Area: " & FormatNumber(fp.Area * 100, 0) & " mm²"
    
    ' Check BaseFace
    Dim baseFace
    Set baseFace = fp.BaseFace
    If Not baseFace Is Nothing Then
        Dim bfArea
        bfArea = baseFace.Evaluator.Area * 100
        WScript.Echo ""
        WScript.Echo "BaseFace area: " & FormatNumber(bfArea, 0) & " mm²"
        
        If bfArea > 6000000 Then
            WScript.Echo ""
            WScript.Echo "*** CORRECT! BaseFace is the large face! ***"
        Else
            WScript.Echo ""
            WScript.Echo "*** WRONG! BaseFace is still the edge face! ***"
        End If
    End If
Else
    WScript.Echo "No flat pattern exists"
End If

WScript.Echo ""
WScript.Echo "=== DONE ==="
