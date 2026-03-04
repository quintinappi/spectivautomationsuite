' =========================================================
' ENABLE THICKNESS EXPORT - MANUAL CHECKBOX + AUTO SAVE
' =========================================================
' Opens Parameters dialog and waits for you to check the
' Export Parameter checkbox for Thickness, then auto-saves
' =========================================================

Option Explicit

Dim m_InventorApp
Dim m_Shell

Main()

Sub Main()
    On Error Resume Next
    
    WScript.Echo "=== ENABLE THICKNESS EXPORT - MANUAL DIALOG ==="
    WScript.Echo ""
    
    Set m_InventorApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Could not connect to Inventor."
        WScript.Quit 1
    End If
    
    Dim activeDoc
    Set activeDoc = m_InventorApp.ActiveDocument
    
    If activeDoc Is Nothing Then
        WScript.Echo "ERROR: No active document."
        WScript.Quit 1
    End If
    
    If activeDoc.DocumentType <> 12290 Then ' kPartDocumentObject
        WScript.Echo "ERROR: Not a part document (.ipt)"
        WScript.Quit 1
    End If
    
    WScript.Echo "Part: " & activeDoc.DisplayName
    WScript.Echo ""
    
    ' Check if Thickness exists
    Dim params
    Set params = activeDoc.ComponentDefinition.Parameters
    
    Dim thicknessParam
    Set thicknessParam = Nothing
    
    On Error Resume Next
    Set thicknessParam = params.Item("Thickness")
    On Error Goto 0
    
    If thicknessParam Is Nothing Then
        WScript.Echo "ERROR: Thickness parameter not found"
        WScript.Quit 1
    End If
    
    WScript.Echo "Found Thickness parameter"
    WScript.Echo ""
    
    ' Initialize Shell for SendKeys
    Set m_Shell = CreateObject("WScript.Shell")
    
    ' Bring Inventor to foreground
    WScript.Echo "Step 1: Activating Inventor window..."
    m_Shell.AppActivate "Autodesk Inventor"
    WScript.Sleep 500
    
    ' Open Parameters dialog
    WScript.Echo "Step 2: Opening Manage > Parameters..."
    m_Shell.SendKeys "%m"
    WScript.Sleep 800
    
    m_Shell.SendKeys "p"
    WScript.Sleep 1500
    
    WScript.Echo ""
    WScript.Echo "INSTRUCTIONS:"
    WScript.Echo "1. In the Parameters dialog, find 'Thickness' under Sheet Metal Parameters"
    WScript.Echo "2. Check the 'Export Parameter' checkbox for Thickness"
    WScript.Echo "3. Click 'Done' to close the dialog"
    WScript.Echo ""
    WScript.Echo "The script will wait 15 seconds then save automatically..."
    WScript.Echo ""
    
    ' Wait 15 seconds for user to manually check the box and close dialog
    WScript.Sleep 15000
    
    ' Save the part
    WScript.Echo "Saving part..."
    activeDoc.Save
    
    WScript.Echo ""
    WScript.Echo "SUCCESS! Part saved."
    WScript.Echo ""
    
End Sub
