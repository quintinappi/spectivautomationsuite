' Simple Part Lister - Fixed version
Option Explicit
Const kAssemblyDocumentObject = 12291

Dim invApp, asmDoc, uniqueParts, allParts

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

If asmDoc.DocumentType <> kAssemblyDocumentObject Then
    WScript.Echo "ERROR: Not an assembly"
    WScript.Quit 1
End If

Set uniqueParts = CreateObject("Scripting.Dictionary")
Set allParts = CreateObject("Scripting.Dictionary")

WScript.Echo "Assembly: " & asmDoc.DisplayName
WScript.Echo "Scanning parts..."
WScript.Echo ""

ScanParts asmDoc, uniqueParts, allParts

WScript.Echo ""
WScript.Echo "TOTAL PARTS FOUND: " & allParts.Count
WScript.Echo ""

Dim i, partKey
i = 1
For Each partKey In allParts.Keys
    Dim pinfo, plateStr, lenStr
    Set pinfo = allParts(partKey)
    
    If pinfo("isPlate") Then
        plateStr = "Yes"
    Else
        plateStr = "No"
    End If
    
    If pinfo("hasLen") Then
        lenStr = "Yes"
    Else
        lenStr = "No"
    End If
    
    WScript.Echo i & ". " & pinfo("name")
    WScript.Echo "   Desc: " & pinfo("desc")
    WScript.Echo "   Plate: " & plateStr & " | Has Length: " & lenStr
    WScript.Echo ""
    
    i = i + 1
Next

Sub ScanParts(doc, seen, found)
    Dim occs, idx, occ, subdoc, fname, fpath, desc, ispl, haslen, pinfo
    
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
                        
                        Err.Clear
                        desc = GetDesc(subdoc)
                        ispl = (InStr(1, UCase(desc), "PL", 1) > 0 Or InStr(1, UCase(desc), "S355JR", 1) > 0)
                        haslen = HasLen(subdoc)
                        
                        Set pinfo = CreateObject("Scripting.Dictionary")
                        pinfo.Add "name", fname
                        pinfo.Add "desc", desc
                        pinfo.Add "isPlate", ispl
                        pinfo.Add "hasLen", haslen
                        
                        found.Add fpath, pinfo
                    End If
                ElseIf LCase(Right(fname, 4)) = ".iam" Then
                    ScanParts subdoc, seen, found
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
    Dim cd, p
    On Error Resume Next
    Err.Clear
    
    Set cd = doc.ComponentDefinition
    If Err.Number = 0 And Not cd Is Nothing Then
        Err.Clear
        Set p = cd.Parameters.UserParameters.Item("Length")
        If Err.Number = 0 Then
            HasLen = True
            Exit Function
        End If
    End If
    HasLen = False
End Function
