' UPDATE SAME-FOLDER DERIVED PARTS - POST-RENAMING FIX
' =============================================================================
' Author: Quintin de Bruin © 2025
' PART RENAMING WORKFLOW - STEP 1c: Fix derived parts (if needed)
'
' This script updates same-folder derived parts to reference renamed files
' Run this AFTER the main renamer (STEP 1b) if you have derived parts that broke
' Reads STEP_1_MAPPING.txt to find new file names and update references
'
' WHEN TO USE:
' - After running the main renamer (STEP 1b)
' - If derived parts are showing errors or broken links
' - To fix parts that derive from renamed components
' =============================================================================

Const kPartDocumentObject = 12290
Const kAssemblyDocumentObject = 12291

Dim invApp, activeDoc, asmFolder
Dim fso, logFile, logPath
Dim fixCount

' Initialize
Set fso = CreateObject("Scripting.FileSystemObject")
logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\Update_Derived_Log.txt"

' Connect to Inventor
On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Then
    MsgBox "Inventor is not running!", vbCritical
    WScript.Quit
End If
On Error GoTo 0

Set activeDoc = invApp.ActiveDocument
If activeDoc Is Nothing Then
    MsgBox "No document open!", vbCritical
    WScript.Quit
End If

If activeDoc.DocumentType <> kAssemblyDocumentObject Then
    MsgBox "Assembly not open!", vbCritical
    WScript.Quit
End If

asmFolder = fso.GetParentFolderName(activeDoc.FullFileName)

' Read mapping file
Dim mappingPath
mappingPath = asmFolder & "\STEP_1_MAPPING.txt"

If Not fso.FileExists(mappingPath) Then
    MsgBox "STEP_1_MAPPING.txt not found!", vbCritical
    WScript.Quit
End If

' Create log
Set logFile = fso.CreateTextFile(logPath, True)
WriteLog "=========================================="
WriteLog " UPDATE SAME-FOLDER DERIVED PARTS"
WriteLog "=========================================="
WriteLog "Assembly: " & activeDoc.DisplayName
WriteLog "Mapping: " & mappingPath
WriteLog ""

' Scan and fix
fixCount = 0
ScanAndFixDerivedParts

' Done
WriteLog ""
WriteLog "=========================================="
WriteLog "COMPLETE"
WriteLog "=========================================="
WriteLog "Total fixes: " & fixCount
WriteLog ""

logFile.Close

MsgBox "Complete!" & vbCrLf & vbCrLf & "Fixed: " & fixCount & " derived parts", vbInformation

Sub ScanAndFixDerivedParts()
    Dim doc, partDoc

    For Each doc In activeDoc.AllReferencedDocuments
        If doc.DocumentType = kPartDocumentObject Then
            FixPartDerivedRefs doc
        End If
    Next
End Sub

Sub FixPartDerivedRefs(partDoc)
    On Error Resume Next

    Dim partDef, refComps, derivedParts
    Dim derivedComp, docDesc, basePath
    Dim partPath, partName
    Dim newBasePath

    partPath = partDoc.FullFileName
    partName = fso.GetFileName(partPath)

    Set partDef = partDoc.ComponentDefinition
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If

    Set refComps = partDef.ReferenceComponents
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If

    Set derivedParts = refComps.DerivedPartComponents
    If Err.Number <> 0 Or derivedParts Is Nothing Then
        Err.Clear
        Exit Sub
    End If

    If derivedParts.Count = 0 Then
        Exit Sub
    End If

    Dim i
    For i = 1 To derivedParts.Count
        Set derivedComp = derivedParts.Item(i)

        If derivedComp.LinkedToFile Then
            Set docDesc = derivedComp.ReferencedDocumentDescriptor
            If Not docDesc Is Nothing Then
                basePath = docDesc.FullDocumentName
                Dim baseName
                baseName = fso.GetBaseName(basePath)

                ' Check if this base file was renamed
                newBasePath = FindNewBasePath(baseName, asmFolder, mappingPath)

                If newBasePath <> "" Then
                    WriteLog "Fixing: " & partName
                    WriteLog "  Old Base: " & fso.GetFileName(basePath)
                    WriteLog "  New Base: " & fso.GetFileName(newBasePath)

                    derivedComp.Replace newBasePath, Nothing

                    If Err.Number = 0 Then
                        WriteLog "  SUCCESS"
                        fixCount = fixCount + 1
                    Else
                        WriteLog "  ERROR: " & Err.Description
                        Err.Clear
                    End If
                    WriteLog ""
                End If
            End If
        End If
    Next

    On Error GoTo 0
End Sub

Function FindNewBasePath(baseName, asmFolder, mappingPath)
    ' Read mapping file to find new base file name
    Dim mappingFile
    Set mappingFile = fso.OpenTextFile(mappingPath, 1)

    Dim newBasePath
    newBasePath = ""

    Dim line, done, parts, originalFile, newFile
    done = False
    While Not mappingFile.AtEndOfStream And Not done
        line = mappingFile.ReadLine

        If Trim(line) <> "" And Left(Trim(line), 1) <> "#" Then
            parts = Split(line, "|")

            If UBound(parts) >= 3 Then
                originalFile = fso.GetBaseName(parts(2))

                If originalFile = baseName Then
                    newFile = fso.GetBaseName(parts(3))
                    newBasePath = asmFolder & "\" & newFile & ".ipt"
                    done = True
                End If
            End If
        End If
    Wend

    mappingFile.Close

    FindNewBasePath = newBasePath
End Function

Sub WriteLog(text)
    logFile.WriteLine text
End Sub
