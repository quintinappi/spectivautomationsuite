' Enable Export Checkbox for Length2 - UI Automation (API doesn't support export for regular parts)
' Run this after Length2 parameter already exists

Option Explicit

Dim m_InventorApp, wShell

Sub Main()
    On Error Resume Next
    
    WScript.Echo "=== ENABLE EXPORT FOR LENGTH2 ==="
    WScript.Echo ""
    
    ' Get Inventor
    Set m_InventorApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Inventor not running"
        WScript.Quit 1
    End If
    
    ' Get active document
    Dim doc
    Set doc = m_InventorApp.ActiveDocument
    If doc Is Nothing Then
        WScript.Echo "ERROR: No active document"
        WScript.Quit 1
    End If
    
    WScript.Echo "Part: " & doc.DisplayName
    WScript.Echo ""
    
    ' Verify Length2 exists
    Dim compDef, userParams, length2Param
    Set compDef = doc.ComponentDefinition
    Set userParams = compDef.Parameters.UserParameters
    
    Err.Clear
    Set length2Param = userParams.Item("Length2")
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Length2 parameter does not exist!"
        WScript.Echo "Please run Fix_Single_Part_Length2.vbs first"
        WScript.Quit 1
    End If
    
    WScript.Echo "Length2 found: " & length2Param.Value & " " & length2Param.Units
    WScript.Echo "Equation: " & length2Param.Expression
    WScript.Echo ""
    WScript.Echo "NOTE: API properties (ExportedToSheet, Export, etc.) are NOT supported"
    WScript.Echo "      for regular part parameters - only for sheet metal parameters."
    WScript.Echo "      Using UI automation instead..."
    WScript.Echo ""
    
    ' Save first to avoid losing changes
    doc.Save
    
    ' UI automation to check export
    WScript.Echo "USER ACTION REQUIRED:"
    WScript.Echo "1. In Inventor, open Manage tab > Parameters"
    WScript.Echo "2. Find 'Length2' in the User Parameters list"
    WScript.Echo "3. Check the 'Export Parameter' checkbox for Length2"
    WScript.Echo "4. Click OK"
    WScript.Echo "5. Save the part"
    WScript.Echo ""
    
    MsgBox "MANUAL STEP REQUIRED:" & vbCrLf & vbCrLf & _
           "The export checkbox cannot be set via API for regular parts." & vbCrLf & vbCrLf & _
           "Please:" & vbCrLf & _
           "1. Open Manage tab > Parameters in Inventor" & vbCrLf & _
           "2. Find 'Length2' in User Parameters" & vbCrLf & _
           "3. Check the 'Export Parameter' checkbox" & vbCrLf & _
           "4. Click OK and Save", vbExclamation, "Manual Action Required"
    
    WScript.Echo "Waiting for user to complete manual step..."
    
End Sub

Main
