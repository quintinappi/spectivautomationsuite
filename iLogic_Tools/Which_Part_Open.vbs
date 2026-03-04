' Which_Part_Open.vbs
Dim invApp, partDoc
Set invApp = GetObject(, "Inventor.Application")
Set partDoc = invApp.ActiveDocument
WScript.Echo "Currently open: " & partDoc.DisplayName
WScript.Echo "Full path: " & partDoc.FullFileName
