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
WScript.Echo "Document Type: " & doc.DocumentType
WScript.Echo "Full Path: " & doc.FullFileName

If doc.DocumentType = 12290 Then ' Part document
    Dim compDef
    Set compDef = doc.ComponentDefinition
    
    WScript.Echo "Document SubType: " & compDef.DocumentSubType.UniqueID
    
    If compDef.DocumentSubType.UniqueID = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
        WScript.Echo "*** PART IS SHEET METAL ***"
        
        ' Check flat pattern
        Dim flatPattern
        Set flatPattern = compDef.FlatPattern
        
        If Not flatPattern Is Nothing Then
            WScript.Echo "Flat Pattern Length: " & Round(flatPattern.Length / 10, 1) & " mm"
            WScript.Echo "Flat Pattern Width: " & Round(flatPattern.Width / 10, 1) & " mm"
            
            If Not flatPattern.BaseFace Is Nothing Then
                WScript.Echo "BaseFace Area: " & Round(flatPattern.BaseFace.Evaluator.Area * 100, 0) & " mm²"
            End If
        End If
    Else
        WScript.Echo "*** PART IS STANDARD (NOT SHEET METAL) ***"
    End If
End If

WScript.Echo ""
WScript.Echo "Ready to convert."
