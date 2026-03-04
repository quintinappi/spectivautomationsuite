' Change_Document_Standard.vbs
' Changes the active standard for the entire IDW document
' This affects all views in the drawing

Option Explicit

Dim invApp
On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")

If Err.Number <> 0 Or invApp Is Nothing Then
    MsgBox "ERROR: Inventor is not running!", vbCritical
    WScript.Quit 1
End If

Dim idwDoc
Set idwDoc = invApp.ActiveDocument

If idwDoc Is Nothing Or idwDoc.DocumentType <> 12292 Then
    MsgBox "ERROR: No IDW file is open!", vbCritical
    WScript.Quit 1
End If

WScript.Echo "=== CHANGE DOCUMENT STANDARD ==="
WScript.Echo "Drawing: " & idwDoc.DisplayName
WScript.Echo ""

' Get styles manager
Dim stylesManager
Set stylesManager = idwDoc.StylesManager

' Get current active standard
Dim currentStandard
Set currentStandard = stylesManager.ActiveStandardStyle
WScript.Echo "Current Active Standard: " & currentStandard.Name
WScript.Echo ""

' Get all available standards
Dim standardStyles
Set standardStyles = stylesManager.StandardStyles

WScript.Echo "Available Standards:"
WScript.Echo ""

Dim i
For i = 1 To standardStyles.Count
    Dim std
    Set std = standardStyles.Item(i)
    WScript.Echo i & ". " & std.Name
Next

WScript.Echo ""
WScript.Echo "Enter the NUMBER of the standard you want to use:"
WScript.Echo "(This will change the standard for the entire drawing)"
WScript.Echo ""

Dim choice
choice = WScript.StdIn.ReadLine

If Not IsNumeric(choice) Then
    WScript.Echo "Invalid choice!"
    WScript.Quit 1
End If

Dim index
index = CInt(choice)

If index < 1 Or index > standardStyles.Count Then
    WScript.Echo "Invalid choice!"
    WScript.Quit 1
End If

Dim newStandard
Set newStandard = standardStyles.Item(index)

WScript.Echo ""
WScript.Echo "Changing standard from '" & currentStandard.Name & "' to '" & newStandard.Name & "'..."

On Error Resume Next
stylesManager.ActiveStandardStyle = newStandard

If Err.Number <> 0 Then
    WScript.Echo "ERROR: " & Err.Description
    WScript.Quit 1
End If

WScript.Echo "SUCCESS! Standard changed."
WScript.Echo ""
WScript.Echo "Saving document..."

idwDoc.Save2

If Err.Number <> 0 Then
    WScript.Echo "ERROR saving: " & Err.Description
Else
    WScript.Echo "Document saved successfully!"
End If

WScript.Echo ""
WScript.Echo "Done!"
