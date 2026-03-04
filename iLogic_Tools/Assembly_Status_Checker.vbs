' Assembly Parts Status Checker
' Author: Quintin de Bruin © 2026

Option Explicit

Sub Main()
    On Error Resume Next

    ' Get Inventor application
    Dim inventorApp
    Set inventorApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Inventor not running"
        Exit Sub
    End If

    ' Get active document
    Dim doc
    Set doc = inventorApp.ActiveDocument

    If doc.DocumentType <> 12291 Then ' kAssemblyDocumentObject
        WScript.Echo "Please open an assembly document"
        Exit Sub
    End If

    WScript.Echo "=== ASSEMBLY PARTS CONVERSION STATUS ==="
    WScript.Echo "Assembly: " & doc.FullFileName
    WScript.Echo ""

    ' Get component definition
    Dim compDef
    Set compDef = doc.ComponentDefinition

    ' Get all occurrences
    Dim occurrences
    Set occurrences = compDef.Occurrences

    Dim convertedCount
    convertedCount = 0
    Dim totalParts
    totalParts = 0

    WScript.Echo "PART CONVERSION STATUS:"
    WScript.Echo "======================="

    Dim i
    For i = 1 To occurrences.Count
        Dim occ
        Set occ = occurrences.Item(i)

        ' Skip suppressed occurrences
        If Not occ.Suppressed Then
            Dim refDoc
            Set refDoc = occ.Definition.Document

            Dim fileName
            fileName = GetFileNameFromPath(refDoc.FullFileName)

            ' Only check .ipt files
            If LCase(Right(fileName, 4)) = ".ipt" Then
                totalParts = totalParts + 1

                ' Check if part is sheet metal
                Dim isSheetMetal
                isSheetMetal = CheckIfSheetMetal(refDoc)

                If isSheetMetal Then
                    convertedCount = convertedCount + 1
                    WScript.Echo "[CONVERTED] " & fileName & " - Sheet Metal ✓"

                    ' Get thickness if possible
                    Dim thickness
                    thickness = GetSheetMetalThickness(refDoc)
                    If thickness <> "" Then
                        WScript.Echo "           Thickness: " & thickness
                    End If

                    ' Get flat pattern dimensions if available
                    Dim flatPatternDims
                    flatPatternDims = GetFlatPatternDimensions(refDoc)
                    If flatPatternDims <> "" Then
                        WScript.Echo "           Flat Pattern: " & flatPatternDims
                    End If
                Else
                    WScript.Echo "[NOT CONVERTED] " & fileName & " - Regular Part ✗"
                End If

                WScript.Echo ""
            End If
        End If
    Next

    WScript.Echo "SUMMARY:"
    WScript.Echo "========"
    WScript.Echo "Total parts checked: " & totalParts
    WScript.Echo "Successfully converted: " & convertedCount
    WScript.Echo "Failed conversion: " & (totalParts - convertedCount)
    WScript.Echo "Success rate: " & FormatPercent(convertedCount / totalParts)

    ' Check assembly parameters
    WScript.Echo ""
    WScript.Echo "ASSEMBLY PARAMETERS:"
    WScript.Echo "===================="

    Dim userParams
    Set userParams = compDef.Parameters.UserParameters

    Dim hasPlateLength
    hasPlateLength = False
    Dim hasPlateWidth
    hasPlateWidth = False

    For i = 1 To userParams.Count
        Dim param
        Set param = userParams.Item(i)
        WScript.Echo param.Name & " = " & param.Value & " " & param.Units

        If param.Name = "PLATE LENGTH" Then hasPlateLength = True
        If param.Name = "PLATE WIDTH" Then hasPlateWidth = True
    Next

    If hasPlateLength And hasPlateWidth Then
        WScript.Echo "✓ Assembly parameters created successfully"
    Else
        WScript.Echo "✗ Assembly parameters missing or incomplete"
    End If

    WScript.Echo ""
    WScript.Echo "=== END REPORT ==="
End Sub

Function CheckIfSheetMetal(partDoc)
    On Error Resume Next

    Dim compDef
    Set compDef = partDoc.ComponentDefinition

    ' Try to cast to sheet metal component definition
    Dim sheetMetalCompDef
    Set sheetMetalCompDef = compDef

    If Err.Number = 0 Then
        ' Check if it has sheet metal parameters
        Dim params
        Set params = compDef.Parameters

        Dim thicknessParam
        Set thicknessParam = params.Item("Thickness")

        If Not thicknessParam Is Nothing Then
            CheckIfSheetMetal = True
        Else
            CheckIfSheetMetal = False
        End If
    Else
        CheckIfSheetMetal = False
    End If

    Err.Clear
End Function

Function GetSheetMetalThickness(partDoc)
    On Error Resume Next

    Dim compDef
    Set compDef = partDoc.ComponentDefinition
    Dim params
    Set params = compDef.Parameters

    Dim thicknessParam
    Set thicknessParam = params.Item("Thickness")

    If Not thicknessParam Is Nothing Then
        GetSheetMetalThickness = thicknessParam.Expression
    Else
        GetSheetMetalThickness = ""
    End If

    Err.Clear
End Function

Function GetFlatPatternDimensions(partDoc)
    On Error Resume Next

    Dim compDef
    Set compDef = partDoc.ComponentDefinition

    Dim sheetMetalCompDef
    Set sheetMetalCompDef = compDef

    If Err.Number = 0 Then
        Dim flatPattern
        Set flatPattern = sheetMetalCompDef.FlatPattern

        If Not flatPattern Is Nothing Then
            Dim length
            length = flatPattern.Length * 10 ' Convert to mm
            Dim width
            width = flatPattern.Width * 10 ' Convert to mm

            GetFlatPatternDimensions = FormatNumber(length, 2) & "mm × " & FormatNumber(width, 2) & "mm"
        Else
            GetFlatPatternDimensions = "No flat pattern"
        End If
    Else
        GetFlatPatternDimensions = "N/A"
    End If

    Err.Clear
End Function

Function GetFileNameFromPath(fullPath)
    GetFileNameFromPath = Mid(fullPath, InStrRev(fullPath, "\") + 1)
End Function

Function FormatPercent(value)
    FormatPercent = FormatNumber(value * 100, 1) & "%"
End Function

Main