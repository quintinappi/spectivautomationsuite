' =========================================================
' ENABLE THICKNESS EXPORT - UI AUTOMATION
' =========================================================
' Opens Parameter dialog and enables Export checkbox
' =========================================================

Option Explicit

Dim m_InventorApp
Dim m_Shell

Main()

Sub Main()
    On Error Resume Next
    
    WScript.Echo "=== ENABLE THICKNESS EXPORT - UI AUTOMATION ==="
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
    
    ' Open Parameters dialog: Tools menu > Parameters
    WScript.Echo "Step 2: Opening Manage menu (Alt+M)..."
    m_Shell.SendKeys "%m"
    WScript.Sleep 800
    
    WScript.Echo "Step 3: Clicking Parameters (P)..."
    m_Shell.SendKeys "p"
    WScript.Sleep 1500
    
    ' Navigate to Thickness parameter under Sheet Metal Parameters
    WScript.Echo "Step 4: Finding Thickness in Sheet Metal Parameters..."
    ' Press Ctrl+Home to go to top of list
    m_Shell.SendKeys "^{HOME}"
    WScript.Sleep 300
    
    ' Press Right to expand the first section
    m_Shell.SendKeys "{RIGHT}"
    WScript.Sleep 300
    
    ' Press Down multiple times to navigate down to Thickness
    ' First down goes to Thickness in the expanded section
    m_Shell.SendKeys "{DOWN}"
    WScript.Sleep 200
    
    ' Now navigate right to the Export Parameter column
    WScript.Echo "Step 5: Navigating to Export Parameter checkbox..."
    ' Tab multiple times to reach Export Parameter column
    Dim tabCount
    For tabCount = 1 To 9
        m_Shell.SendKeys "{TAB}"
        WScript.Sleep 80
    Next
    
    ' Toggle the checkbox
    WScript.Echo "Step 6: Enabling Export Parameter checkbox..."
    m_Shell.SendKeys " "  ' Space to toggle checkbox
    WScript.Sleep 300
    
    ' Click Done button (Alt+D or look for OK)
    WScript.Echo "Step 7: Clicking Done..."
    m_Shell.SendKeys "%d"
    WScript.Sleep 1000
    
    ' Save the part
    WScript.Echo ""
    WScript.Echo "Saving part..."
    activeDoc.Save
    
    WScript.Echo ""
    WScript.Echo "SUCCESS! Thickness export parameter enabled."
    WScript.Echo ""
    
End Sub
