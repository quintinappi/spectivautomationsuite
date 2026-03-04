' TEST SCRIPT - Quick Flat Pattern Check and Fix
' Author: Quintin de Bruin © 2026

Option Explicit

Const kSheetMetalSubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"

Sub Main()
    On Error Resume Next

    Dim invApp
    Set invApp = GetObject(, "Inventor.Application")

    Dim partDoc
    Set partDoc = invApp.ActiveDocument
    
    WScript.Echo "Part: " & partDoc.DisplayName
    
    If partDoc.SubType <> kSheetMetalSubType Then
        WScript.Echo "NOT sheet metal - SubType: " & partDoc.SubType
        Exit Sub
    End If

    Dim smDef
    Set smDef = partDoc.ComponentDefinition

    ' Check flat pattern state
    WScript.Echo ""
    If smDef.HasFlatPattern Then
        Dim fp
        Set fp = smDef.FlatPattern
        Dim fpL, fpW
        fpL = fp.Length * 10
        fpW = fp.Width * 10
        
        WScript.Echo "Flat Pattern: " & FormatNumber(fpL, 1) & " x " & FormatNumber(fpW, 1) & " mm"
        
        If fpL >= 50 And fpW >= 50 Then
            WScript.Echo "ORIENTATION: CORRECT (both dimensions > 50mm)"
        Else
            WScript.Echo "ORIENTATION: WRONG (edge view - one dimension is thickness)"
        End If
    Else
        WScript.Echo "No flat pattern exists"
        
        ' Create one
        WScript.Echo "Creating flat pattern..."
        smDef.Unfold
        
        If Err.Number = 0 Then
            WScript.Echo "Created!"
            If smDef.HasFlatPattern Then
                Set fp = smDef.FlatPattern
                fpL = fp.Length * 10
                fpW = fp.Width * 10
                WScript.Echo "Dimensions: " & FormatNumber(fpL, 1) & " x " & FormatNumber(fpW, 1) & " mm"
            End If
        Else
            WScript.Echo "Unfold failed: " & Err.Description
            Err.Clear
        End If
    End If
    
    WScript.Echo ""
End Sub

Main
