' List ALL Part Names - Simple output
Option Explicit
Const kAssemblyDocumentObject = 12291

Dim invApp, asmDoc, uniqueParts

On Error Resume Next

Set invApp = GetObject(, "Inventor.Application")
Set asmDoc = invApp.ActiveDocument

WScript.Echo "Assembly: " & asmDoc.DisplayName
WScript.Echo ""

Set uniqueParts = CreateObject("Scripting.Dictionary")
ListAll asmDoc, uniqueParts

WScript.Echo ""
WScript.Echo "Total unique parts: " & uniqueParts.Count

Sub ListAll(doc, seen)
    Dim occs, idx, occ, subdoc, fname, fpath
    
    Err.Clear
    Set occs = doc.ComponentDefinition.Occurrences
    If Err.Number <> 0 Then Exit Sub
    
    For idx = 1 To occs.Count
        Err.Clear
        Set occ = occs.Item(idx)
        If Err.Number = 0 And Not occ.Suppressed Then
            Set subdoc = occ.Definition.Document
            fname = Mid(subdoc.FullFileName, InStrRev(subdoc.FullFileName, "\") + 1)
            fpath = subdoc.FullFileName
            
            If LCase(Right(fname, 4)) = ".ipt" Then
                If Not seen.Exists(fpath) Then
                    seen.Add fpath, True
                    WScript.Echo fname
                End If
            ElseIf LCase(Right(fname, 4)) = ".iam" Then
                ListAll subdoc, seen
            End If
        End If
    Next
End Sub
