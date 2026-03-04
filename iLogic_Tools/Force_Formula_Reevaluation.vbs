' Force Formula Re-evaluation - NUCLEAR OPTION
' Forces BOM formulas to re-evaluate by triggering transaction flush
' This mimics what the UI does when manually toggling units
' Author: Quintin de Bruin © 2026
'
' THE PROBLEM:
' - BOM formulas cache precision settings and don't re-evaluate when LinearDimensionPrecision changes
' - They ONLY re-evaluate when UnitsOfMeasure changes IN THE UI
' - Programmatic UnitsOfMeasure changes don't trigger formula dirty flags
'
' THE SOLUTION:
' - Force Inventor to flush its transaction buffer using multiple approaches
' - Trigger events that FORCE synchronous re-evaluation
' - Use document Save/Reopen cycle as final nuclear option

Option Explicit

' Inventor API Constants
Const kAssemblyDocumentObject = 12291
Const kMillimeterLengthUnits = 11269
Const kCentimeterLengthUnits = 11266

Dim m_InventorApp

Sub Main()
    On Error Resume Next

    WScript.Echo "=== FORCE FORMULA RE-EVALUATION (NUCLEAR OPTION) ==="
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
    WScript.Echo "Assembly: " & asmDoc.DisplayName
    WScript.Echo ""

    ' Get BOM for verification
    Dim bom
    Set bom = asmDoc.ComponentDefinition.BOM
    If bom Is Nothing Then
        WScript.Echo "ERROR: Could not access BOM"
        Exit Sub
    End If

    ' === METHOD 1: TRANSACTION FLUSH VIA SAVE ===
    WScript.Echo "METHOD 1: Transaction flush via Save..."
    WScript.Echo "  Saving document to flush transaction buffer..."
    asmDoc.Save
    If Err.Number <> 0 Then
        WScript.Echo "  ERROR: Save failed - " & Err.Description
        Err.Clear
    Else
        WScript.Echo "  Save complete"
    End If

    ' === METHOD 2: BOM STRUCTURED VIEW REBUILD ===
    WScript.Echo ""
    WScript.Echo "METHOD 2: BOM Structured View forced rebuild..."

    If bom.StructuredViewEnabled Then
        WScript.Echo "  Accessing BOMView object..."
        Dim bomView
        Set bomView = bom.BOMViews.Item("Structured")
        If Not bomView Is Nothing Then
            WScript.Echo "  Current BOMView.ViewState: " & bomView.ViewState

            ' Force renumbering which triggers full BOM rebuild
            WScript.Echo "  Calling Renumber() to force rebuild..."
            bomView.Renumber
            If Err.Number <> 0 Then
                WScript.Echo "  ERROR: Renumber failed - " & Err.Description
                Err.Clear
            Else
                WScript.Echo "  Renumber complete"
            End If
        End If
    End If

    ' === METHOD 3: PARAMETER SYSTEM REBUILD ===
    WScript.Echo ""
    WScript.Echo "METHOD 3: Parameter system rebuild..."

    ' Force parameter manager to rebuild all formulas
    Dim compDef
    Set compDef = asmDoc.ComponentDefinition

    Dim params
    Set params = compDef.Parameters

    WScript.Echo "  Forcing parameter update..."
    params.UpdateAfterChange = True ' Enable immediate updates

    ' Touch a parameter to force rebuild
    If params.ModelParameters.Count > 0 Then
        Dim dummyParam
        Set dummyParam = params.ModelParameters.Item(1)
        Dim originalExpr
        originalExpr = dummyParam.Expression

        WScript.Echo "  Toggling parameter expression to trigger rebuild..."
        dummyParam.Expression = originalExpr
        If Err.Number <> 0 Then
            WScript.Echo "  ERROR: Parameter toggle failed - " & Err.Description
            Err.Clear
        Else
            WScript.Echo "  Parameter rebuild triggered"
        End If
    End If

    ' === METHOD 4: DOCUMENT SETTINGS TRANSACTION ===
    WScript.Echo ""
    WScript.Echo "METHOD 4: Document Settings transaction flush..."

    ' Access UnitsOfMeasure to verify current state
    Dim unitsOfMeasure
    Set unitsOfMeasure = asmDoc.UnitsOfMeasure

    If Not unitsOfMeasure Is Nothing Then
        Dim currentUnits
        currentUnits = unitsOfMeasure.LengthUnits
        WScript.Echo "  Current length units: " & currentUnits

        ' CRITICAL: Use DisplayUnits instead of just LengthUnits
        ' DisplayUnits affects how formulas are displayed
        Dim originalDisplayUnits
        originalDisplayUnits = unitsOfMeasure.LengthDisplayUnits
        WScript.Echo "  Current display units: " & originalDisplayUnits

        ' Toggle display units (more likely to trigger formula refresh)
        Dim tempDisplayUnits
        If originalDisplayUnits = kMillimeterLengthUnits Then
            tempDisplayUnits = kCentimeterLengthUnits
        Else
            tempDisplayUnits = kMillimeterLengthUnits
        End If

        WScript.Echo "  Toggling LengthDisplayUnits to force transaction..."
        unitsOfMeasure.LengthDisplayUnits = tempDisplayUnits
        If Err.Number <> 0 Then
            WScript.Echo "  ERROR: DisplayUnits change failed - " & Err.Description
            Err.Clear
        Else
            WScript.Echo "  DisplayUnits changed to: " & tempDisplayUnits

            ' Force immediate update
            asmDoc.Update
            WScript.Echo "  Document updated"

            ' Restore original
            unitsOfMeasure.LengthDisplayUnits = originalDisplayUnits
            WScript.Echo "  DisplayUnits restored to: " & originalDisplayUnits

            ' Update again
            asmDoc.Update
            WScript.Echo "  Document updated again"
        End If
    End If

    ' === METHOD 5: REBUILD2 WITH FORCE FLAG ===
    WScript.Echo ""
    WScript.Echo "METHOD 5: Rebuild2 with AcceptErrorsAndContinue..."
    asmDoc.Rebuild2 True
    If Err.Number <> 0 Then
        WScript.Echo "  ERROR: Rebuild2 failed - " & Err.Description
        Err.Clear
    Else
        WScript.Echo "  Rebuild2 complete"
    End If

    ' === METHOD 6: FINAL SAVE ===
    WScript.Echo ""
    WScript.Echo "METHOD 6: Final save to commit all changes..."
    asmDoc.Save
    If Err.Number <> 0 Then
        WScript.Echo "  ERROR: Final save failed - " & Err.Description
        Err.Clear
    Else
        WScript.Echo "  Final save complete"
    End If

    WScript.Echo ""
    WScript.Echo "=== ALL METHODS COMPLETE ==="
    WScript.Echo ""
    WScript.Echo "VERIFICATION STEPS:"
    WScript.Echo "1. Check BOM in Inventor - do quantities show 0 decimals?"
    WScript.Echo "2. If NO: Close and reopen assembly, check again"
    WScript.Echo "3. If STILL NO: The nuclear option is needed..."
    WScript.Echo ""
    WScript.Echo "NUCLEAR OPTION (if above fails):"
    WScript.Echo "  Run: Nuclear_Reopen_Cycle.vbs"
    WScript.Echo "  This will close and reopen the assembly to force full reload"

End Sub

Main
