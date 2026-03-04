On Error Resume Next

Dim invApp
Set invApp = GetObject(, "Inventor.Application")

If invApp Is Nothing Then
    WScript.Echo "ERROR: Cannot connect to Inventor"
    WScript.Quit 1
End If

Dim doc
Set doc = invApp.ActiveDocument

If doc Is Nothing Then
    WScript.Echo "ERROR: No active document"
    WScript.Quit 1
End If

WScript.Echo "=== ACTIVE DOCUMENT INFO ==="
WScript.Echo "Display Name: " & doc.DisplayName
WScript.Echo "Full Path: " & doc.FullFileName
WScript.Echo "Document Type: " & doc.DocumentType & " (12291=Assembly, 12290=Part)"
WScript.Echo "Is Dirty: " & doc.Dirty
WScript.Echo ""

If doc.DocumentType = 12290 Then
    Dim compDef
    Set compDef = doc.ComponentDefinition
    
    WScript.Echo "SubType: " & doc.SubType
    
    If doc.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
        WScript.Echo "Part is: SHEET METAL"
        
        If compDef.HasFlatPattern Then
            Dim fp
            Set fp = compDef.FlatPattern
            WScript.Echo "Flat Pattern: " & Round(fp.Length / 10, 1) & " x " & Round(fp.Width / 10, 1) & " mm"
            
            If Not fp.BaseFace Is Nothing Then
                WScript.Echo "BaseFace Area: " & FormatNumber(Round(fp.BaseFace.Evaluator.Area * 100, 0), 0) & " mm²"
            End If
        End If
    Else
        WScript.Echo "Part is: STANDARD (not sheet metal)"
    End If
End If
