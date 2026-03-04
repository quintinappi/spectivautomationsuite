' Find FL25 Parts - Check for Length parameter
Option Explicit
Const kAssemblyDocumentObject = 12291

Dim invApp, asmDoc, uniqueParts, fl25Parts

On Error Resume Next

Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Or invApp Is Nothing Then
    WScript.Echo "ERROR: Inventor not running"
    WScript.Quit 1
End If

Set asmDoc = invApp.ActiveDocument
If Err.Number <> 0 Or asmDoc Is Nothing Then
    WScript.Echo "ERROR: No active document"
    WScript.Quit 1
End If

Set uniqueParts = CreateObject("Scripting.Dictionary")
Set fl25Parts = CreateObject("Scripting.Dictionary")

WScript.Echo "Assembly: " & asmDoc.DisplayName
WScript.Echo "Searching for FL25 parts..."
WScript.Echo ""

ScanForFL25 asmDoc, uniqueParts, fl25Parts

WScript.Echo ""
WScript.Echo "FL25 PARTS FOUND: " & fl25Parts.Count
WScript.Echo ""

If fl25Parts.Count = 0 Then
    WScript.Echo "No parts with 'FL25' in the name were found."
Else
    Dim i, partKey
    i = 1
    For Each partKey In fl25Parts.Keys
        Dim pinfo
        Set pinfo = fl25Parts(partKey)
        
        WScript.Echo i & ". " & pinfo("name")
        WScript.Echo "   Full Path: " & pinfo("path")
        WScript.Echo "   Description: " & pinfo("desc")
        
        If pinfo("hasLen") Then
            WScript.Echo "   >>> HAS Length parameter"
        Else
            WScript.Echo "   >>> NO Length parameter <<<<<"
        End If
        
        WScript.Echo ""
        i = i + 1
    Next
End If

Sub ScanForFL25(doc, seen, found)
    Dim occs, idx, occ, subdoc, fname, fpath, desc, haslen, pinfo
    
    Err.Clear
    Set occs = doc.ComponentDefinition.Occurrences
    If Err.Number <> 0 Then Exit Sub
    
    For idx = 1 To occs.Count
        Err.Clear
        Set occ = occs.Item(idx)
        If Err.Number = 0 Then
            If Not occ.Suppressed Then
                Set subdoc = occ.Definition.Document
                fname = Mid(subdoc.FullFileName, InStrRev(subdoc.FullFileName, "\") + 1)
                fpath = subdoc.FullFileName
                
                If LCase(Right(fname, 4)) = ".ipt" Then
                    If Not seen.Exists(fpath) Then
                        seen.Add fpath, True
                        
                        ' Check if filename contains FL25
                        If InStr(1, UCase(fname), "FL25", 1) > 0 Then
                            Err.Clear
                            desc = GetDesc(subdoc)
                            haslen = HasLen(subdoc)
                            
                            Set pinfo = CreateObject("Scripting.Dictionary")
                            pinfo.Add "name", fname
                            pinfo.Add "path", fpath
                            pinfo.Add "desc", desc
                            pinfo.Add "hasLen", haslen
                            
                            found.Add fpath, pinfo
                            
                            WScript.Echo "Found: " & fname & " (Has Length: " & haslen & ")"
                        End If
                    End If
                ElseIf LCase(Right(fname, 4)) = ".iam" Then
                    ScanForFL25 subdoc, seen, found
                End If
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
