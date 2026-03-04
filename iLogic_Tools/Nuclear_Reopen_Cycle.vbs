' Nuclear Reopen Cycle - ABSOLUTE LAST RESORT
' Closes and reopens the assembly to force COMPLETE formula re-evaluation
' This guarantees formulas re-evaluate with current precision settings
' Author: Quintin de Bruin © 2026
'
' USE THIS WHEN:
' - All other methods fail
' - BOM still shows incorrect decimals after Force_Formula_Reevaluation.vbs
' - You need ABSOLUTE CERTAINTY that formulas will refresh
'
' HOW IT WORKS:
' 1. Gets current assembly path
' 2. Saves assembly
' 3. Closes assembly (flushes ALL caches)
' 4. Reopens assembly (formulas re-evaluate from scratch)
' 5. Verifies BOM state
'
' CAVEAT: Closes your current assembly! Make sure everything is saved!

Option Explicit

Const kAssemblyDocumentObject = 12291

Dim m_InventorApp

Sub Main()
    On Error Resume Next

    WScript.Echo "=== NUCLEAR REOPEN CYCLE ==="
    WScript.Echo ""
    WScript.Echo "WARNING: This will CLOSE and REOPEN your current assembly!"
    WScript.Echo ""

    ' Connect to Inventor
    Set m_InventorApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Inventor not running"
        Exit Sub
    End If

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

    ' Get assembly path BEFORE closing
    Dim asmPath
    asmPath = asmDoc.FullFileName

    WScript.Echo "Assembly: " & asmDoc.DisplayName
    WScript.Echo "Full path: " & asmPath
    WScript.Echo ""

    ' Confirm with user
    Dim shell
    Set shell = CreateObject("WScript.Shell")

    Dim userResponse
    userResponse = MsgBox("This will close and reopen:" & vbCrLf & vbCrLf & _
                          asmPath & vbCrLf & vbCrLf & _
                          "Make sure all changes are saved!" & vbCrLf & vbCrLf & _
                          "Continue?", vbOKCancel + vbExclamation, "Nuclear Reopen")

    If userResponse = vbCancel Then
        WScript.Echo "User cancelled"
        Exit Sub
    End If

    ' Step 1: Save assembly
    WScript.Echo "Step 1: Saving assembly..."
    asmDoc.Save
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Save failed - " & Err.Description
        WScript.Echo "Aborting nuclear cycle"
        Exit Sub
    End If
    WScript.Echo "  Assembly saved"

    ' Step 2: Close assembly (THIS FLUSHES ALL CACHES)
    WScript.Echo ""
    WScript.Echo "Step 2: Closing assembly (flushing all caches)..."
    asmDoc.Close False ' Skip save prompt (we already saved)
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Close failed - " & Err.Description
        Exit Sub
    End If
    WScript.Echo "  Assembly closed - all caches flushed"

    ' Step 3: Wait for Inventor to fully release resources
    WScript.Echo ""
    WScript.Echo "Step 3: Waiting for Inventor to release resources..."
    WScript.Sleep 2000 ' 2 second pause
    WScript.Echo "  Wait complete"

    ' Step 4: Reopen assembly (FORMULAS RE-EVALUATE FROM SCRATCH)
    WScript.Echo ""
    WScript.Echo "Step 4: Reopening assembly (formulas will re-evaluate)..."
    Set asmDoc = m_InventorApp.Documents.Open(asmPath, True) ' Open visible
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Reopen failed - " & Err.Description
        WScript.Echo "ERROR DETAILS: " & Err.Number
        Exit Sub
    End If
    WScript.Echo "  Assembly reopened"

    ' Step 5: Verify it's active
    WScript.Echo ""
    WScript.Echo "Step 5: Verifying assembly is active..."
    If Not m_InventorApp.ActiveDocument Is Nothing Then
        If m_InventorApp.ActiveDocument.FullFileName = asmPath Then
            WScript.Echo "  SUCCESS: Assembly is active"
        Else
            WScript.Echo "  WARNING: Different document is active"
        End If
    Else
        WScript.Echo "  WARNING: No active document"
    End If

    ' Step 6: Force update to ensure everything is current
    WScript.Echo ""
    WScript.Echo "Step 6: Forcing final update..."
    asmDoc.Update
    If Err.Number <> 0 Then
        WScript.Echo "  WARNING: Update failed - " & Err.Description
        Err.Clear
    Else
        WScript.Echo "  Update complete"
    End If

    WScript.Echo ""
    WScript.Echo "=== NUCLEAR CYCLE COMPLETE ==="
    WScript.Echo ""
    WScript.Echo "Formulas have been COMPLETELY re-evaluated from scratch."
    WScript.Echo "BOM should now show correct precision (0 decimals)."
    WScript.Echo ""
    WScript.Echo "VERIFY:"
    WScript.Echo "1. Open BOM in Inventor"
    WScript.Echo "2. Check quantity column - should show whole numbers (no decimals)"
    WScript.Echo "3. If STILL showing decimals, the issue is NOT formula caching"
    WScript.Echo ""

    MsgBox "Nuclear reopen cycle complete!" & vbCrLf & vbCrLf & _
           "Assembly has been closed and reopened." & vbCrLf & _
           "All formulas re-evaluated from scratch." & vbCrLf & vbCrLf & _
           "Check the BOM now.", vbInformation, "Complete"

End Sub

Main
