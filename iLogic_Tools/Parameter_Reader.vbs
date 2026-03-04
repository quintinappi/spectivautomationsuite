' Parameter Reader for Lug DM-UP.ipt
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

    ' Path to the Part4 DM-UP.ipt file
    Dim partPath
    partPath = "C:\Users\Quintin\Documents\Spectiv\3. Working\22. DMS plant\22. SSCR05 - Primary Sinks D&R Screen Station\722 Underpan\DM Underpan\Part4 DM-UP.ipt"

    ' Open the part document
    Dim partDoc
    Set partDoc = inventorApp.Documents.Open(partPath, False)
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Could not open part file: " & Err.Description
        Exit Sub
    End If

    WScript.Echo "=== PART4 DM-UP.IPT PARAMETERS ==="
    WScript.Echo "File: " & partDoc.FullFileName
    WScript.Echo ""

    ' Get component definition
    Dim compDef
    Set compDef = partDoc.ComponentDefinition

    ' Get all parameters
    Dim params
    Set params = compDef.Parameters

    WScript.Echo "ALL PARAMETERS:"
    WScript.Echo "---------------"

    Dim i
    For i = 1 To params.Count
        Dim param
        Set param = params.Item(i)
        WScript.Echo param.Name & " = " & param.Value & " " & param.Units & " (" & GetParamType(param.ParameterType) & ")"
    Next

    WScript.Echo ""
    WScript.Echo "USER PARAMETERS:"
    WScript.Echo "----------------"

    ' Get user parameters specifically
    Dim userParams
    Set userParams = compDef.Parameters.UserParameters

    For i = 1 To userParams.Count
        Set param = userParams.Item(i)
        WScript.Echo param.Name & " = " & param.Value & " " & param.Units
    Next

    ' Check thickness in different ways
    WScript.Echo "THICKNESS ANALYSIS:"
    WScript.Echo "-------------------"

    ' Check the thickness parameter value and units
    Dim thicknessParam
    Set thicknessParam = params.Item("Thickness")
    If Not thicknessParam Is Nothing Then
        WScript.Echo "Thickness parameter value: " & thicknessParam.Value
        WScript.Echo "Thickness parameter units: " & thicknessParam.Units
        WScript.Echo "Thickness parameter expression: " & thicknessParam.Expression

        ' Check internal units
        WScript.Echo "Thickness in document units: " & partDoc.UnitsOfMeasure.ConvertUnits(thicknessParam.Value, thicknessParam.Units, "mm") & " mm"
    End If

    ' Check sheet metal style thickness
    If Not sheetMetalCompDef Is Nothing Then
        WScript.Echo "Sheet metal style thickness value: " & sheetMetalCompDef.ActiveSheetMetalStyle.Thickness.Value
        WScript.Echo "Sheet metal style thickness expression: " & sheetMetalCompDef.ActiveSheetMetalStyle.Thickness.Expression

        ' Convert to mm for clarity
        Dim thicknessInMm
        thicknessInMm = partDoc.UnitsOfMeasure.ConvertUnits(sheetMetalCompDef.ActiveSheetMetalStyle.Thickness.Value, "cm", "mm")
        WScript.Echo "Sheet metal thickness converted to mm: " & thicknessInMm & " mm"
    End If

    ' Close without saving
    partDoc.Close False

    WScript.Echo ""
    WScript.Echo "=== END REPORT ==="
End Sub

Function GetParamType(paramType)
    Select Case paramType
        Case 1: GetParamType = "Number"
        Case 2: GetParamType = "Boolean"
        Case 3: GetParamType = "Text"
        Case Else: GetParamType = "Unknown"
    End Select
End Function

Main