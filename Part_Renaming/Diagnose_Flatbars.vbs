' Diagnostic script to find why flatbars are being missed
' Run this with Inventor open and an assembly loaded

On Error Resume Next

Dim invApp
Set invApp = GetObject(, "Inventor.Application")

If invApp Is Nothing Then
    MsgBox "ERROR: Cannot connect to Inventor. Make sure Inventor is running.", vbCritical
    WScript.Quit 1
End If

Dim doc
Set doc = invApp.ActiveDocument

If doc Is Nothing Then
    MsgBox "ERROR: No active document. Please open an assembly.", vbCritical
    WScript.Quit 1
End If

If doc.DocumentType <> 12291 Then
    MsgBox "ERROR: Active document is not an assembly. Please open an assembly.", vbCritical
    WScript.Quit 1
End If

' Create output string
Dim report
report = "=== FLATBAR DIAGNOSTIC REPORT ===" & vbCrLf & vbCrLf
report = report & "Assembly: " & doc.DisplayName & vbCrLf & vbCrLf

' Dictionary to track unique parts
Dim uniqueParts
Set uniqueParts = CreateObject("Scripting.Dictionary")

' Find all parts
Call ScanAssembly(doc, uniqueParts, report)

' Show results
MsgBox report, vbInformation, "Flatbar Diagnosis"

' Save to file
Dim fso, file
Set fso = CreateObject("Scripting.FileSystemObject")
Dim logPath
logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\Flatbar_Diagnosis.txt"
Set file = fso.CreateTextFile(logPath, True)
file.WriteLine report
file.Close

MsgBox "Report saved to: " & logPath, vbInformation

Sub ScanAssembly(asmDoc, uniqueParts, ByRef report)
    Dim occurrences
    Set occurrences = asmDoc.ComponentDefinition.Occurrences

    Dim i
    For i = 1 To occurrences.Count
        Dim occ
        Set occ = occurrences.Item(i)

        If Not occ.Suppressed Then
            Dim partDoc
            Set partDoc = occ.Definition.Document

            Dim fileName
            fileName = partDoc.DisplayName

            ' Only process .ipt files
            If LCase(Right(fileName, 4)) = ".ipt" Then
                Dim fullPath
                fullPath = partDoc.FullFileName

                ' Skip duplicates
                If Not uniqueParts.Exists(fullPath) Then
                    uniqueParts.Add fullPath, True

                    ' Get description
                    Dim desc
                    desc = GetDescription(partDoc)

                    If desc <> "" Then
                        Dim descUpper
                        descUpper = UCase(Trim(desc))

                        ' Check if it looks like a flatbar
                        If InStr(descUpper, "FL") > 0 Or InStr(descUpper, "FLAT") > 0 Or InStr(descUpper, "STIFF") > 0 Then
                            report = report & "FILE: " & fileName & vbCrLf
                            report = report & "  Description: " & desc & vbCrLf
                            report = report & "  Upper: " & descUpper & vbCrLf
                            report = report & "  First 2 chars: [" & Left(descUpper, 2) & "]" & vbCrLf
                            report = report & "  Contains FLOOR: " & CBool(InStr(descUpper, "FLOOR") > 0) & vbCrLf
                            report = report & "  Classification: " & ClassifyDescription(desc) & vbCrLf
                            report = report & vbCrLf
                        End If
                    End If
                End If
            ElseIf LCase(Right(fileName, 4)) = ".iam" Then
                ' Recurse into sub-assemblies
                Call ScanAssembly(partDoc, uniqueParts, report)
            End If
        End If
    Next
End Sub

Function GetDescription(doc)
    On Error Resume Next
    Dim propertySet
    Set propertySet = doc.PropertySets.Item("Design Tracking Properties")
    If Err.Number <> 0 Then
        GetDescription = ""
        Exit Function
    End If

    Dim descProp
    Set descProp = propertySet.Item("Description")
    If Err.Number <> 0 Then
        GetDescription = ""
        Exit Function
    End If

    GetDescription = Trim(descProp.Value)
End Function

Function ClassifyDescription(description)
    Dim desc
    desc = UCase(Trim(description))

    If InStr(desc, "BOLT") > 0 Or InStr(desc, "SCREW") > 0 Or InStr(desc, "WASHER") > 0 Or InStr(desc, "NUT") > 0 Then
        ClassifyDescription = "SKIP (hardware)"
        Exit Function
    End If

    If Left(desc, 2) = "UB" Then
        ClassifyDescription = "B (UB beam)"
    ElseIf Left(desc, 2) = "UC" Then
        ClassifyDescription = "B (UC column)"
    ElseIf Left(desc, 2) = "PL" Then
        If InStr(desc, "S355JR") > 0 Then
            ClassifyDescription = "PL (platework)"
        Else
            ClassifyDescription = "LPL (liner)"
        End If
    ElseIf Left(desc, 1) = "L" And (InStr(desc, "X") > 0 Or InStr(desc, " X ") > 0) Then
        ClassifyDescription = "A (angle)"
    ElseIf Left(desc, 3) = "PFC" Then
        ClassifyDescription = "CH (PFC)"
    ElseIf Left(desc, 3) = "TFC" Then
        ClassifyDescription = "CH (TFC)"
    ElseIf Left(desc, 4) = "PIPE" Then
        ClassifyDescription = "P (pipe)"
    ElseIf Left(desc, 3) = "SHS" Then
        ClassifyDescription = "SQ (SHS)"
    ElseIf Left(desc, 2) = "FL" And Not InStr(desc, "FLOOR") > 0 Then
        ClassifyDescription = "FL (flatbar) ✓"
    ElseIf Left(desc, 3) = "IPE" Then
        ClassifyDescription = "IPE (European I-beam)"
    Else
        ClassifyDescription = "OTHER (unclassified)"
    End If
End Function
