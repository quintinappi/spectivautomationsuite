Option Explicit

' Dump all occurrence names from the currently open assembly

Dim invApp
Set invApp = GetObject(, "Inventor.Application")

Dim activeDoc
Set activeDoc = invApp.ActiveDocument

If activeDoc Is Nothing Then
    MsgBox "No document is open!", vbCritical
    WScript.Quit
End If

If LCase(Right(activeDoc.FullFileName, 4)) <> ".iam" Then
    MsgBox "Please open an assembly (.iam) file!", vbCritical
    WScript.Quit
End If

Dim output
output = "OCCURRENCE NAMES IN: " & activeDoc.DisplayName & vbCrLf & vbCrLf

Dim occurrences
Set occurrences = activeDoc.ComponentDefinition.Occurrences

output = output & "Total occurrences: " & occurrences.Count & vbCrLf & vbCrLf

Dim i
For i = 1 To occurrences.Count
    Dim occ
    Set occ = occurrences.Item(i)

    Dim refPath
    refPath = occ.ReferencedFileDescriptor.FullFileName

    Dim fileName
    fileName = Mid(refPath, InStrRev(refPath, "\") + 1)

    output = output & i & ". """ & occ.Name & """ --> " & fileName & vbCrLf
Next

' Save to file
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Dim outputPath
outputPath = fso.GetParentFolderName(activeDoc.FullFileName) & "\OccurrenceNames.txt"

Dim file
Set file = fso.CreateTextFile(outputPath, True)
file.Write output
file.Close

MsgBox "Occurrence names dumped to:" & vbCrLf & outputPath, vbInformation

WScript.Quit
