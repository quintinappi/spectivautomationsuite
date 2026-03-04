' UI Automation - Single Part Test
' Tests UI automation on just the first plate part
' Author: Quintin de Bruin © 2026

Option Explicit

Const kAssemblyDocumentObject = 12291

Dim m_InventorApp
Dim m_Shell

Sub Main()
    On Error Resume Next

    WScript.Echo "=== UI AUTOMATION - SINGLE PART TEST ==="
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

    ' Find first plate part
    Dim occurrences
    Set occurrences = asmDoc.ComponentDefinition.Occurrences

    Dim foundPart
    foundPart = False

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
                    WScript.Echo "Found first plate part: " & partNumber
                    WScript.Echo "Path: " & refDoc.FullFileName
                    WScript.Echo ""

                    If ProcessPlatePart(refDoc.FullFileName, partNumber) Then
                        WScript.Echo "SUCCESS!"
                    Else
                        WScript.Echo "FAILED!"
                    End If

                    foundPart = True
                    Exit For
                End If
            End If
        End If
    Next

    If Not foundPart Then
        WScript.Echo "No plate parts found"
    End If

End Sub

Function ProcessPlatePart(partPath, partNumber)
    On Error Resume Next

    ProcessPlatePart = False

    WScript.Echo "Step 1: Opening part..."

    ' Open the part
    Dim partDoc
    Set partDoc = m_InventorApp.Documents.Open(partPath, False)

    If Err.Number <> 0 Then
        WScript.Echo "ERROR opening part: " & Err.Description
        Err.Clear
        Exit Function
    End If

    WScript.Echo "  Part opened"

    ' Make it active
    m_InventorApp.ActiveDocument.Update
    WScript.Sleep 1000

    WScript.Echo "Step 2: Activating Inventor window..."

    ' Activate Inventor window
    On Error Resume Next
    m_Shell.AppActivate "Autodesk Inventor"
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Could not activate Inventor window - " & Err.Description
        Err.Clear
        partDoc.Close True
        Exit Function
    End If

    WScript.Sleep 1000

    WScript.Echo "Step 3: Opening Document Settings (Alt+T, then D)..."

    ' Send Alt+T for Tools menu
    m_Shell.SendKeys "%t"
    WScript.Sleep 500

    ' Send D for Document Settings
    m_Shell.SendKeys "d"
    WScript.Sleep 2000

    WScript.Echo "Step 4: Navigating to precision field (Tab x5)..."

    ' Tab to precision dropdown (may need adjustment)
    Dim j
    For j = 1 To 5
        m_Shell.SendKeys "{TAB}"
        WScript.Sleep 200
    Next

    WScript.Echo "Step 5: Toggling precision (Up, then Down)..."

    ' Toggle precision
    m_Shell.SendKeys "{UP}"
    WScript.Sleep 300
    m_Shell.SendKeys "{DOWN}"
    WScript.Sleep 300

    WScript.Echo "Step 6: Clicking OK (Enter)..."

    ' Press Enter to close dialog
    m_Shell.SendKeys "{ENTER}"
    WScript.Sleep 1500

    WScript.Echo "Step 7: Saving part..."

    ' Save
    partDoc.Save
    If Err.Number <> 0 Then
        WScript.Echo "ERROR saving: " & Err.Description
        Err.Clear
        partDoc.Close True
        Exit Function
    End If

    WScript.Echo "  Part saved"

    WScript.Echo "Step 8: Closing part..."

    ' Close
    partDoc.Close False

    WScript.Echo "  Part closed"

    ProcessPlatePart = True

End Function

Main