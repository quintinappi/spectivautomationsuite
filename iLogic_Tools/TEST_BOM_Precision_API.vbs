' TEST BOM PRECISION - API ONLY (Auto-run version)
' Tests API-only method on current assembly without confirmation

Option Explicit

Const kAssemblyDocumentObject = 12291
Const kPartDocumentObject = 12290
Const kMillimeterLengthUnits = 11269
Const kCentimeterLengthUnits = 11266

Dim m_InventorApp
Dim m_Log

' Auto-accept confirmation
Dim AUTO_MODE
AUTO_MODE = True

Sub Main()
    On Error Resume Next
    
    m_Log = "=== API-ONLY BOM PRECISION TEST ===" & vbCrLf
    m_Log = m_Log & "Started: " & Now & vbCrLf & vbCrLf
    
    Set m_InventorApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Cannot connect to Inventor"
        WScript.Quit 1
    End If
    
    ' Get assembly
    If m_InventorApp.ActiveDocument Is Nothing Then
        WScript.Echo "ERROR: No active document"
        WScript.Quit 1
    End If
    
    If m_InventorApp.ActiveDocument.DocumentType <> kAssemblyDocumentObject Then
        WScript.Echo "ERROR: Not an assembly"
        WScript.Quit 1
    End If
    
    Dim asmDoc
    Set asmDoc = m_InventorApp.ActiveDocument
    
    WScript.Echo "Assembly: " & asmDoc.DisplayName
    m_Log = m_Log & "Assembly: " & asmDoc.FullFileName & vbCrLf
    
    ' Scan for plate parts
    Dim plateParts
    Set plateParts = ScanPlateParts(asmDoc)
    
    WScript.Echo "Plate parts found: " & plateParts.Count
    
    If plateParts.Count = 0 Then
        WScript.Echo "No plate parts to process"
        WScript.Quit 0
    End If
    
    ' Process ONLY first 2 parts as test
    Dim testCount
    testCount = 2
    If plateParts.Count < testCount Then testCount = plateParts.Count
    
    WScript.Echo ""
    WScript.Echo "Testing API method on first " & testCount & " part(s)..."
    WScript.Echo ""
    
    Dim processed
    processed = 0
    
    Dim i, keys
    keys = plateParts.Keys
    
    For i = 0 To testCount - 1
        Dim partPath
        partPath = keys(i)
        
        Dim partName
        partName = plateParts(partPath)
        
        WScript.Echo "Testing: " & partName
        
        If TestAPIMethod(partPath) Then
            WScript.Echo "  API Method: SUCCESS"
            processed = processed + 1
        Else
            WScript.Echo "  API Method: FAILED"
        End If
        
        WScript.Echo ""
    Next
    
    WScript.Echo "========================================"
    WScript.Echo "TEST COMPLETE"
    WScript.Echo "Processed: " & processed & "/" & testCount
    WScript.Echo ""
    WScript.Echo "Check your BOM to see if precision updated."
    WScript.Echo "If decimals still show, API-only doesn't work"
    WScript.Echo "and we need the Add-in version with UI."
    WScript.Echo "========================================"
    
End Sub

Function ScanPlateParts(asmDoc)
    Dim result
    Set result = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    
    Dim occ
    For Each occ In asmDoc.ComponentDefinition.Occurrences
        Dim refDoc
        Set refDoc = occ.Definition.Document
        
        If Not refDoc Is Nothing Then
            If refDoc.DocumentType = kPartDocumentObject Then
                Dim partNum
                partNum = ""
                On Error Resume Next
                partNum = refDoc.PropertySets("Design Tracking Properties")("Part Number").Value
                On Error GoTo 0
                
                If InStr(UCase(partNum), "PL") > 0 Or InStr(UCase(partNum), "S355JR") > 0 Then
                    Dim fullPath
                    fullPath = refDoc.FullFileName
                    If Not result.Exists(fullPath) Then
                        result.Add fullPath, partNum
                    End If
                End If
            End If
        End If
        Err.Clear
    Next
    
    Set ScanPlateParts = result
End Function

Function TestAPIMethod(partPath)
    On Error Resume Next
    
    TestAPIMethod = False
    
    ' Open part invisible
    Dim partDoc
    Set partDoc = m_InventorApp.Documents.Open(partPath, False)
    
    If Err.Number <> 0 Or partDoc Is Nothing Then
        WScript.Echo "    ERROR: Could not open part"
        Exit Function
    End If
    
    ' Check BEFORE values
    Dim uom
    Set uom = partDoc.UnitsOfMeasure
    
    Dim beforePrec
    beforePrec = uom.LengthDisplayPrecision
    WScript.Echo "    Before: Precision = " & beforePrec
    
    ' Toggle precision via API
    uom.LengthDisplayPrecision = 3
    partDoc.Update
    
    uom.LengthDisplayPrecision = 0
    partDoc.Update
    
    ' Check AFTER values
    Dim afterPrec
    afterPrec = uom.LengthDisplayPrecision
    WScript.Echo "    After: Precision = " & afterPrec
    
    ' Force dirty flag
    On Error Resume Next
    Dim params
    Set params = partDoc.ComponentDefinition.Parameters
    Dim dummy
    Set dummy = params.UserParameters.AddByValue("_TEST_", 0, "mm")
    dummy.Value = 1
    params.UserParameters.RemoveByName("_TEST_")
    partDoc.Update
    Err.Clear
    
    ' Save
    partDoc.Save
    
    ' Close
    partDoc.Close True
    
    If Err.Number = 0 Then
        TestAPIMethod = True
    End If
End Function

Main
