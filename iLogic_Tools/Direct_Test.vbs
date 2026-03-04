' Direct Check - Very Simple Test
Dim app, doc
Set app = GetObject(, "Inventor.Application")
Set doc = app.ActiveDocument

WScript.Echo "Document: " & doc.DisplayName
WScript.Echo "Type: " & doc.DocumentType

Dim cd
Set cd = doc.ComponentDefinition
WScript.Echo "ComponentDefinition: " & TypeName(cd)

Dim occs
Set occs = cd.Occurrences
WScript.Echo "Occurrences.Count: " & occs.Count

If occs.Count > 0 Then
    Dim occ1
    Set occ1 = occs.Item(1)
    WScript.Echo "First occurrence: " & occ1.Name
    WScript.Echo "Suppressed: " & occ1.Suppressed
    
    Dim subDoc
    Set subDoc = occ1.Definition.Document
    WScript.Echo "SubDoc filename: " & subDoc.FullFileName
End If
