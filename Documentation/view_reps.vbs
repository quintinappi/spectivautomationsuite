On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")
Set sheet = invApp.ActiveDocument.ActiveSheet

WScript.Echo "View Representation Diagnostic"
WScript.Echo "============================"

For Each v In sheet.DrawingViews
    WScript.Echo "View: " & v.Name
    
    WScript.Echo "  ActiveDesignViewRepresentation: " & v.ActiveDesignViewRepresentation
    WScript.Echo "  ActivePositionalRepresentation: " & v.ActivePositionalRepresentation
    WScript.Echo "  ActiveLevelOfDetailRepresentation: " & v.ActiveLevelOfDetailRepresentation
    
    ' Check if it's "Associative"
    WScript.Echo "  DesignViewAssociative: " & v.DesignViewAssociative
    
    WScript.Echo "----------------------------"
Next
