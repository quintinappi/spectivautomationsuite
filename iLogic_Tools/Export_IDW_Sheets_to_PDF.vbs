Option Explicit

' Export_IDW_Sheets_to_PDF.vbs - DETAILING WORKFLOW STEP 14: Export IDW sheets to PDF
' DETAILING WORKFLOW - STEP 14: Export IDW sheets to PDF with correct numbering
' Export each sheet of the open IDW to a separate PDF with numbered file names
' All Colors As Black enabled

Const kDrawingDocumentObject = 12292

Sub Main()
    On Error Resume Next

    Dim invApp
    Set invApp = GetObject(, "Inventor.Application")

    If invApp Is Nothing Then
        WScript.Echo "ERROR: Inventor is not running!"
        Exit Sub
    End If

    If invApp.ActiveDocument Is Nothing Then
        WScript.Echo "ERROR: No active document!"
        Exit Sub
    End If

    Dim oDoc
    Set oDoc = invApp.ActiveDocument

    If oDoc.DocumentType <> kDrawingDocumentObject Then
        WScript.Echo "ERROR: Active document is not a drawing!"
        Exit Sub
    End If

    ' Get base name: full path without .idw
    Dim fullPath
    fullPath = oDoc.FullFileName
    Dim baseName
    baseName = Left(fullPath, Len(fullPath) - 4) ' remove .idw

    WScript.Echo "Exporting sheets from: " & oDoc.DisplayName
    WScript.Echo "Base name: " & baseName

    ' Enable All Colors As Black at application level
    Dim oPDFOptions
    Set oPDFOptions = invApp.ApplicationOptions.ExportOptions.PDFExportOptions

    If Not oPDFOptions Is Nothing Then
        ' Try to set All Colors As Black
        On Error Resume Next
        oPDFOptions.AllColorsAsBlack = True
        If Err.Number = 0 Then
            WScript.Echo "All Colors As Black: ENABLED"
        Else
            WScript.Echo "Note: Could not set All Colors As Black option"
            Err.Clear
        End If
        On Error GoTo 0
    End If

    ' Export each sheet
    Dim i
    For i = 1 To oDoc.Sheets.Count
        Dim filePath
        filePath = baseName & "-" & Right("0" & CStr(i), 2) & ".pdf"

        ' Set active sheet
        oDoc.Sheets(i).Activate

        Err.Clear

        ' Export using simple SaveAs method
        oDoc.SaveAs filePath, 12294

        If Err.Number <> 0 Then
            WScript.Echo "ERROR exporting sheet " & i & ": " & Err.Description
            Err.Clear
        Else
            ' Verify file exists
            Dim fso
            Set fso = CreateObject("Scripting.FileSystemObject")
            If fso.FileExists(filePath) Then
                WScript.Echo "Exported sheet " & i & " to: " & filePath
            Else
                WScript.Echo "WARNING: Sheet " & i & " reported success but file not found"
            End If
        End If
    Next

    WScript.Echo "Export complete."
End Sub

Call Main()
