On Error Resume Next

Dim invApp
Set invApp = GetObject(, "Inventor.Application")

Dim doc
Set doc = invApp.ActiveDocument

WScript.Echo "Currently open: " & doc.DisplayName
WScript.Echo ""
WScript.Echo "Please switch to Part2 DM-UP.ipt in Inventor, then run this script again."
