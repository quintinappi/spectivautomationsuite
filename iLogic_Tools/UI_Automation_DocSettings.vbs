' UI Automation - Document Settings Precision Toggle
' Uses SendKeys to open Document Settings and toggle precision
' Author: Quintin de Bruin © 2026

Option Explicit

Const kAssemblyDocumentObject = 12291

Dim m_InventorApp
Dim m_Shell

Sub Main()
    On Error Resume Next

    WScript.Echo "=== UI AUTOMATION - DOCUMENT SETTINGS ==="
    WScript.Echo ""

    Set m_InventorApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Inventor not running"
        Exit Sub
    End If

    Set m_Shell = CreateObject("WScript.Shell")

    If m_InventorApp.ActiveDocument Is Nothing Then
        WScript.Echo "ERROR: No active document"
        Exit Sub
    End If

    If m_InventorApp.ActiveDocument.DocumentType <> kAssemblyDocumentObject Then
        WScript.Echo "ERROR: Not an assembly"
        Exit Sub
    End If

    Dim asmDoc
    Set asmDoc = m_InventorApp.ActiveDocument
    WScript.Echo "Assembly: " & asmDoc.DisplayName
    WScript.Echo ""

    ' Collect unique plate parts
    Dim plateParts
    Set plateParts = CreateObject("Scripting.Dictionary")

    Dim occurrences
    Set occurrences = asmDoc.ComponentDefinition.Occurrences

    Dim i
    For i = 1 To occurrences.Count
        Dim occ
        Set occ = occurrences.Item(i)

        If Not occ.Suppressed Then
            Dim refDoc
            Set refDoc = occ.Definition.Document

            If LCase(Right(refDoc.FullFileName, 4)) = ".ipt" Then
                Dim partNumber
                partNumber = ""
                On Error Resume Next
                partNumber = refDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value
                Err.Clear

                If InStr(UCase(partNumber), "PL") > 0 Then
                    Dim partPath
                    partPath = refDoc.FullFileName

                    If Not plateParts.Exists(partPath) Then
                        plateParts.Add partPath, partNumber
                    End If
                End If
            End If
        End If
    Next

    WScript.Echo "Found " & plateParts.Count & " unique plate parts"
    WScript.Echo ""

    ' Process each unique plate part
    Dim updatedCount
    updatedCount = 0

    Dim partPath
    For Each partPath In plateParts.Keys
        WScript.Echo "Processing: " & plateParts.Item(partPath)

        If ProcessPlatePart(partPath) Then
            updatedCount = updatedCount + 1
            WScript.Echo "  SUCCESS"
        Else
            WScript.Echo "  FAILED"
        End If

        WScript.Echo ""

        ' Small delay between parts
        WScript.Sleep 500
    Next

    WScript.Echo "=== COMPLETE ==="
    WScript.Echo "Parts updated: " & updatedCount

End Sub

Function ProcessPlatePart(partPath)
    On Error Resume Next

    ProcessPlatePart = False

    WScript.Echo "  Opening: " & partPath

    ' Open the part
    Dim partDoc
    Set partDoc = m_InventorApp.Documents.Open(partPath, False)

    If Err.Number <> 0 Then
        WScript.Echo "  ERROR: Could not open part - " & Err.Description
        Err.Clear
        Exit Function
    End If

    ' Activate the part window
    m_InventorApp.ActiveDocument = partDoc
    WScript.Sleep 500

    ' Bring Inventor window to foreground
    m_Shell.AppActivate "Autodesk Inventor"
    WScript.Sleep 500

    WScript.Echo "  Opening Document Settings via UI automation..."

    ' Open Document Settings: Tools → Document Settings
    ' Method 1: Use keyboard shortcut (if available)
    ' Method 2: Use menu access keys

    ' Try: Alt+T (Tools menu) → D (Document Settings)
    m_Shell.SendKeys "%t"  ' Alt+T for Tools menu
    WScript.Sleep 300
    m_Shell.SendKeys "d"   ' D for Document Settings
    WScript.Sleep 1000

    ' Wait for Document Settings dialog to open
    WScript.Echo "  Waiting for Document Settings dialog..."
    WScript.Sleep 1500

    ' Navigate to Units tab (usually Ctrl+Tab or just Tab)
    ' The dialog should open on the Units tab by default
    ' If not, we need to click on it

    ' Tab to the precision field (depends on dialog layout)
    ' For now, assume it's accessible
    m_Shell.SendKeys "{TAB}"
    WScript.Sleep 200
    m_Shell.SendKeys "{TAB}"
    WScript.Sleep 200
    m_Shell.SendKeys "{TAB}"
    WScript.Sleep 200

    WScript.Echo "  Toggling precision..."

    ' Change precision: Up arrow then Down arrow
    m_Shell.SendKeys "{UP}"
    WScript.Sleep 300
    m_Shell.SendKeys "{DOWN}"
    WScript.Sleep 300

    ' Press OK to apply changes
    WScript.Echo "  Applying changes (OK button)..."
    m_Shell.SendKeys "{ENTER}"
    WScript.Sleep 1000

    ' Save the part
    WScript.Echo "  Saving part..."
    partDoc.Save
    If Err.Number <> 0 Then
        WScript.Echo "  ERROR: Could not save - " & Err.Description
        Err.Clear
        partDoc.Close True ' Close without saving
        Exit Function
    End If

    ' Close the part
    partDoc.Close False ' Skip save prompt (already saved)

    ProcessPlatePart = True

End Function

Main