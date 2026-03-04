' Test_Single_IDW_Refs.vbs
' Scans a single IDW file using Inventor with silent options
' Shows ALL references to diagnose cross-folder issues

Option Explicit

Dim oFSO, oInventor, oDoc
Dim sIDWPath

' The specific IDW to test
sIDWPath = "C:\Users\Quintin\Documents\Spectiv\3. Working\22. DMS plant\21. SSCR04 - Primary Floats D&R Screen Station\000 Structure & Walkway\Knee Brace-1\MGY-200-DRP-01-12.idw"

Set oFSO = CreateObject("Scripting.FileSystemObject")

If Not oFSO.FileExists(sIDWPath) Then
    WScript.Echo "ERROR: File not found: " & sIDWPath
    WScript.Quit 1
End If

WScript.Echo "IDW: " & oFSO.GetFileName(sIDWPath)
WScript.Echo ""

' Connect to running Inventor
On Error Resume Next
Set oInventor = GetObject(, "Inventor.Application")
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Inventor must be running!"
    WScript.Quit 1
End If
On Error GoTo 0

' Enable silent operation
oInventor.SilentOperation = True

' Open invisibly - don't use OpenWithOptions, use standard Open
On Error Resume Next
Set oDoc = oInventor.Documents.Open(sIDWPath, False)
If Err.Number <> 0 Then
    WScript.Echo "ERROR opening: " & Err.Description
    oInventor.SilentOperation = False
    WScript.Quit 1
End If
On Error GoTo 0

WScript.Echo "ALL REFERENCED FILES:"
WScript.Echo "=" & String(70, "=")

Dim oRefFiles, oRefFile, sRefPath, nCount
Set oRefFiles = oDoc.File.ReferencedFileDescriptors
nCount = 0

For Each oRefFile In oRefFiles
    nCount = nCount + 1
    sRefPath = ""
    On Error Resume Next
    sRefPath = oRefFile.FullFileName
    On Error GoTo 0
    
    ' Check if path contains SCR04 or SCR06
    Dim sFlag
    sFlag = ""
    If InStr(1, sRefPath, "SCR04", vbTextCompare) > 0 Then
        sFlag = " [SCR04 - OK]"
    ElseIf InStr(1, sRefPath, "SCR06", vbTextCompare) > 0 Then
        sFlag = " [SCR06] *** CROSS-REF ***"
    ElseIf InStr(1, sRefPath, "SCR05", vbTextCompare) > 0 Then
        sFlag = " [SCR05] *** CROSS-REF ***"
    End If
    
    WScript.Echo nCount & ". " & oFSO.GetFileName(sRefPath) & sFlag
    WScript.Echo "   " & sRefPath
Next

WScript.Echo ""
WScript.Echo "Total references: " & nCount

oDoc.Close True
Set oDoc = Nothing

oInventor.SilentOperation = False

WScript.Echo ""
WScript.Echo "Done!"
