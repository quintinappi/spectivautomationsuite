' Check ALL FL Parts for Length Parameter
Option Explicit
Const kAssemblyDocumentObject = 12291

Dim invApp, asmDoc, uniqueParts, flParts

On Error Resume Next

Set invApp = GetObject(, "Inventor.Application")
Set asmDoc = invApp.ActiveDocument

WScript.Echo "Assembly: " & asmDoc.DisplayName
WScript.Echo "Searching for ALL FL parts (FL1, FL2, etc.)..."
WScript.Echo ""

Set uniqueParts = CreateObject("Scripting.Dictionary")
Set flParts = CreateObject("Scripting.Dictionary")

ScanForFL asmDoc, uniqueParts, flParts

WScript.Echo ""
WScript.Echo "============================================"
WScript.Echo "TOTAL FL PARTS FOUND: " & flParts.Count
WScript.Echo "============================================"
WScript.Echo ""

Dim withLength, withoutLength
withLength = 0
withoutLength = 0

If flParts.Count > 0 Then
    Dim partKey
    For Each partKey In flParts.Keys
        Dim pinfo
        Set pinfo = flParts(partKey)
        
        If pinfo("hasLen") Then
            withLength = withLength + 1
        Else
            withoutLength = withoutLength + 1
            WScript.Echo "NO LENGTH: " & pinfo("name")
            WScript.Echo "   Path: " & pinfo("path")
            WScript.Echo "   Description: " & pinfo("desc")
            WScript.Echo ""
        End If
    Next
    
    WScript.Echo "--------------------------------------------"
    WScript.Echo "FL Parts WITH Length parameter: " & withLength
    WScript.Echo "FL Parts WITHOUT Length parameter: " & withoutLength
    WScript.Echo "--------------------------------------------"
End If

Sub ScanForFL(doc, seen, found)
    Dim occs, idx, occ, subdoc, fname, fpath, desc, haslen, pinfo
    
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
                    
                    ' Check if filename contains -FL followed by a number
                    If InStr(1, UCase(fname), "-FL", 1) > 0 Then
                        Err.Clear
                        desc = GetDesc(subdoc)
                        haslen = HasLen(subdoc)
                        
                        Set pinfo = CreateObject("Scripting.Dictionary")
                        pinfo.Add "name", fname
                        pinfo.Add "path", fpath
                        pinfo.Add "desc", desc
                        pinfo.Add "hasLen", haslen
                        
                        found.Add fpath, pinfo
                    End If
                End If
            ElseIf LCase(Right(fname, 4)) = ".iam" Then
                ScanForFL subdoc, seen, found
            End If
        End If
    Next
End Sub

Function GetDesc(doc)
    Dim ps, dp
    On Error Resume Next
    Err.Clear
    
    Set ps = doc.PropertySets.Item("Design Tracking Properties")
    If Err.Number = 0 Then
        Err.Clear
        Set dp = ps.Item("Description")
        If Err.Number = 0 Then
            GetDesc = Trim(dp.Value)
            Exit Function
        End If
    End If
    GetDesc = "(none)"
End Function

Function HasLen(doc)
    Dim cd, params, lenParam
    On Error Resume Next
    Err.Clear
    
    Set cd = doc.ComponentDefinition
    If Err.Number <> 0 Or cd Is Nothing Then
        HasLen = False
        Exit Function
    End If
    
    Err.Clear
    Set params = cd.Parameters.UserParameters
    If Err.Number <> 0 Then
        HasLen = False
        Exit Function
    End If
    
    Err.Clear
    Set lenParam = params.Item("Length")
    If Err.Number = 0 And Not lenParam Is Nothing Then
        HasLen = True
    Else
        HasLen = False
    End If
End Function
