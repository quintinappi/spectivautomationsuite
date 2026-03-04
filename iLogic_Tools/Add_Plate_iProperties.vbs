' Add PLATE LENGTH and PLATE WIDTH custom iProperties to sheet metal part
' These will be formula-based, referencing the flat pattern dimensions

Option Explicit

Sub Main()
    On Error Resume Next

    Dim invApp
    Set invApp = GetObject(, "Inventor.Application")

    If invApp.ActiveDocument Is Nothing Then
        MsgBox "No active document"
        Exit Sub
    End If

    Dim partDoc
    Set partDoc = invApp.ActiveDocument

    WScript.Echo "Adding PLATE LENGTH and PLATE WIDTH iProperties to: " & partDoc.DisplayName

    ' Check if it's a sheet metal part
    Dim smDef
    Set smDef = partDoc.ComponentDefinition

    If Not smDef.HasFlatPattern Then
        WScript.Echo "ERROR: Part does not have a flat pattern"
        Exit Sub
    End If

    ' Get the actual flat pattern dimensions first
    Dim flatPattern
    Set flatPattern = smDef.FlatPattern

    Dim length, width
    length = flatPattern.Length * 10 ' cm to mm
    width = flatPattern.Width * 10

    WScript.Echo "Flat pattern dimensions: " & FormatNumber(length, 2) & "mm x " & FormatNumber(width, 2) & "mm"

    ' Get custom property set
    Dim customPropSet
    Set customPropSet = partDoc.PropertySets.Item("Inventor User Defined Properties")

    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Could not get custom property set: " & Err.Description
        Exit Sub
    End If

    ' Add or update PLATE LENGTH
    Dim lengthProp
    On Error Resume Next
    Set lengthProp = customPropSet.Item("PLATE LENGTH")

    If Err.Number <> 0 Then
        ' Property doesn't exist, add it
        Err.Clear
        WScript.Echo "Adding PLATE LENGTH property..."

        ' Try formula first: =<Sheet Metal Length>
        customPropSet.Add "=<Sheet Metal Length>", "PLATE LENGTH"

        If Err.Number <> 0 Then
            WScript.Echo "Formula 1 failed, trying: =<Flat Pattern Length>"
            Err.Clear
            customPropSet.Add "=<Flat Pattern Length>", "PLATE LENGTH"
        End If

        If Err.Number <> 0 Then
            WScript.Echo "Formula 2 failed, trying: " & FormatNumber(length, 3) & " mm"
            Err.Clear
            customPropSet.Add FormatNumber(length, 3) & " mm", "PLATE LENGTH"
        End If

        If Err.Number = 0 Then
            WScript.Echo "PLATE LENGTH added successfully"
        Else
            WScript.Echo "ERROR adding PLATE LENGTH: " & Err.Description
        End If
    Else
        WScript.Echo "PLATE LENGTH already exists: " & lengthProp.Value
    End If
    Err.Clear

    ' Add or update PLATE WIDTH
    Dim widthProp
    On Error Resume Next
    Set widthProp = customPropSet.Item("PLATE WIDTH")

    If Err.Number <> 0 Then
        ' Property doesn't exist, add it
        Err.Clear
        WScript.Echo "Adding PLATE WIDTH property..."

        ' Try formula first: =<Sheet Metal Width>
        customPropSet.Add "=<Sheet Metal Width>", "PLATE WIDTH"

        If Err.Number <> 0 Then
            WScript.Echo "Formula 1 failed, trying: =<Flat Pattern Width>"
            Err.Clear
            customPropSet.Add "=<Flat Pattern Width>", "PLATE WIDTH"
        End If

        If Err.Number <> 0 Then
            WScript.Echo "Formula 2 failed, trying: " & FormatNumber(width, 3) & " mm"
            Err.Clear
            customPropSet.Add FormatNumber(width, 3) & " mm", "PLATE WIDTH"
        End If

        If Err.Number = 0 Then
            WScript.Echo "PLATE WIDTH added successfully"
        Else
            WScript.Echo "ERROR adding PLATE WIDTH: " & Err.Description
        End If
    Else
        WScript.Echo "PLATE WIDTH already exists: " & widthProp.Value
    End If
    Err.Clear

    ' Save the part
    partDoc.Save
    WScript.Echo ""
    WScript.Echo "Part saved successfully"

    ' Show all custom properties
    WScript.Echo ""
    WScript.Echo "=== CUSTOM iPROPERTIES ==="
    Dim i
    For i = 1 To customPropSet.Count
        Dim prop
        Set prop = customPropSet.Item(i)
        WScript.Echo prop.Name & " = " & prop.Value
    Next

End Sub

Main
