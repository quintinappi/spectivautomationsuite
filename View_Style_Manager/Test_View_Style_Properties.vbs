' Test_View_Style_Properties.vbs
' Final comprehensive test of all possible view style properties

Option Explicit

Dim invApp
On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")

Dim idwDoc
Set idwDoc = invApp.ActiveDocument

WScript.Echo "Testing view: " & idwDoc.ActiveSheet.DrawingViews.Item(1).Name
WScript.Echo ""

Dim view
Set view = idwDoc.ActiveSheet.DrawingViews.Item(1)

' List ALL properties we can try
Dim props
props = Array("Style", "StyleName", "Standard", "ViewStyle", "DrawingStandard", _
              "StyleSource", "StyleSourceType", "OverrideStyle", "UseStandard", _
              "ApplyStandard", "BaseViewStyle", "ParentStyle")

Dim prop
For Each prop In props
    On Error Resume Next
    Err.Clear
    
    Dim val
    val = ""
    
    ' Try to get the property
    Execute("val = view." & prop)
    
    If Err.Number = 0 Then
        If IsObject(val) Then
            If Not val Is Nothing Then
                On Error Resume Next
                WScript.Echo prop & " = [Object: " & val.Name & "]"
                If Err.Number <> 0 Then
                    WScript.Echo prop & " = [Object - no Name property]"
                End If
                Err.Clear
            Else
                WScript.Echo prop & " = [Object: NULL]"
            End If
        Else
            WScript.Echo prop & " = " & val
        End If
    Else
        ' Property doesn't exist
    End If
    Err.Clear
Next

WScript.Echo ""
WScript.Echo "Now checking if we can SET ActiveStandardStyle and force update..."

' Try to force the document to update all views
On Error Resume Next
idwDoc.Update
If Err.Number = 0 Then
    WScript.Echo "idwDoc.Update() - SUCCESS"
Else
    WScript.Echo "idwDoc.Update() - " & Err.Description
End If
Err.Clear

' Try to rebuild
On Error Resume Next
idwDoc.Rebuild
If Err.Number = 0 Then
    WScript.Echo "idwDoc.Rebuild() - SUCCESS"
Else
    WScript.Echo "idwDoc.Rebuild() - " & Err.Description
End If
Err.Clear

WScript.Echo ""
WScript.Echo "=== COMPLETE ==="
