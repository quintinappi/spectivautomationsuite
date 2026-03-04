Option Explicit
On Error Resume Next
Dim invApp
Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Or invApp Is Nothing Then
    WScript.Echo "ERROR: Inventor is not running or cannot be reached."
    WScript.Quit 1
End If
Err.Clear

Dim docs
Set docs = invApp.Documents
If docs.Count = 0 Then
    WScript.Echo "No documents are open in Inventor."
    WScript.Quit 0
End If

Dim doc
WScript.Echo "Open Inventor Documents:"
WScript.Echo "-------------------------"
Dim activeDoc
Set activeDoc = invApp.ActiveDocument
For Each doc In docs
    Dim t
    t = doc.DocumentType
    Dim tname
    Select Case t
        Case 12290
            tname = "Part (.ipt)"
        Case 12291
            tname = "Assembly (.iam)"
        Case 12292
            tname = "Drawing (.idw/.dwg)"
        Case Else
            tname = "Other (Type=" & t & ")"
    End Select

    Dim activeMark
    If Not activeDoc Is Nothing And LCase(doc.FullFileName) = LCase(activeDoc.FullFileName) Then
        activeMark = "[ACTIVE]"
    Else
        activeMark = ""
    End If

    WScript.Echo tname & " - " & doc.DisplayName & " " & activeMark
    WScript.Echo "    Path: " & doc.FullFileName
Next

' Summarize active assembly and drawing if present
Dim activeAsm, activeDrw
Set activeAsm = Nothing
Set activeDrw = Nothing
For Each doc In docs
    If doc.DocumentType = 12291 Then Set activeAsm = doc
    If doc.DocumentType = 12292 Then Set activeDrw = doc
Next
WScript.Echo ""
If Not activeAsm Is Nothing Then
    WScript.Echo "Active Assembly (first found): " & activeAsm.DisplayName & " - " & activeAsm.FullFileName
Else
    WScript.Echo "No assembly open."
End If
If Not activeDrw Is Nothing Then
    WScript.Echo "Open Drawing (first found): " & activeDrw.DisplayName & " - " & activeDrw.FullFileName
Else
    WScript.Echo "No drawing open."
End If
