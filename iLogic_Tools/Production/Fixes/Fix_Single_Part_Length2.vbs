' Fix Single Part - Add Length2 parameter linked to length dimension
' Creates Length2 user parameter referencing the longest model dimension
' Manual export checkbox enabling required

Option Explicit

' Inventor API Constants
Const kNumberParameterType = 1
Const kMillimeterLengthUnits = 11269

Dim m_InventorApp

Sub Main()
    On Error Resume Next
    
    WScript.Echo "=== FIX NON-PLATE PART - ADD LENGTH2 PARAMETER ==="
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
    
    If doc.DocumentType <> 12290 Then ' Must be a part
        WScript.Echo "ERROR: Please open FL25 part file (not assembly)"
        WScript.Quit 1
    End If
    
    WScript.Echo "Part: " & doc.DisplayName
    WScript.Echo ""
    
    ' Get component definition and parameters
    Dim compDef, params, modelParams, userParams
    Set compDef = doc.ComponentDefinition
    Set params = compDef.Parameters
    Set modelParams = params.ModelParameters
    Set userParams = params.UserParameters
    
    ' Step 1: Find the length parameter (largest value in mm)
    WScript.Echo "STEP 1: Finding length parameter..."
    Dim maxValue, maxParamName, maxParam
    maxValue = 0
    maxParamName = ""
    Set maxParam = Nothing
    
    Dim i, param
    For i = 1 To modelParams.Count
        Err.Clear
        Set param = modelParams.Item(i)
        If Err.Number = 0 Then
            Dim paramUnits
            paramUnits = LCase(Trim(param.Units))
            If paramUnits = "mm" Or paramUnits = "" Then
                If param.ModelValue > maxValue Then
                    maxValue = param.ModelValue
                    maxParamName = param.Name
                    Set maxParam = param
                End If
            End If
        End If
    Next
    
    If maxParam Is Nothing Then
        WScript.Echo "ERROR: Could not find length parameter"
        WScript.Quit 1
    End If
    
    WScript.Echo "  Found: " & maxParamName & " = " & maxParam.Value & " " & maxParam.Units
    WScript.Echo "  Model Value: " & maxValue & " (base units)"
    WScript.Echo ""
    
    ' Step 2: Create Length2 user parameter with the value from length param
    WScript.Echo "STEP 2: Creating Length2 user parameter..."
    
    Err.Clear
    Dim length2Param
    Set length2Param = userParams.Item("Length2")
    If Err.Number = 0 Then
        WScript.Echo "  Length2 already exists, deleting it first..."
        length2Param.Delete
    End If
    
    ' Create Length2 using AddByValue method
    Err.Clear
    Set length2Param = userParams.AddByValue("Length2", maxParam.Value, kMillimeterLengthUnits)
    If Err.Number <> 0 Then
        WScript.Echo "ERROR creating Length2: " & Err.Description
        WScript.Quit 1
    End If
    
    WScript.Echo "  Created Length2 = " & length2Param.Value & " " & length2Param.Units
    WScript.Echo ""
    
    ' Step 3: Set Length2's equation to reference the model parameter
    WScript.Echo "STEP 3: Linking Length2 equation to " & maxParamName & "..."
    Err.Clear
    length2Param.Expression = maxParamName
    If Err.Number <> 0 Then
        WScript.Echo "ERROR setting Length2 expression: " & Err.Description
        WScript.Quit 1
    End If
    WScript.Echo "  Length2.Expression = '" & maxParamName & "'"
    WScript.Echo ""
    
    ' Update the document
    WScript.Echo "Updating document..."
    doc.Update
    
    WScript.Echo ""
    WScript.Echo "=========================================="
    WScript.Echo "SUCCESS! Length2 parameter created and linked"
    WScript.Echo "=========================================="
    WScript.Echo ""
    WScript.Echo "Verification:"
    WScript.Echo "  " & maxParamName & " = " & maxParam.Value & " (Expression: " & maxParam.Expression & ")"
    WScript.Echo "  Length2 = " & length2Param.Value & " (Expression: " & length2Param.Expression & ")"
    WScript.Echo ""
    WScript.Echo "NOTE: To enable BOM export, manually check 'Export Parameter'"
    WScript.Echo "      checkbox for Length2 in Manage > Parameters dialog"
    WScript.Echo ""
    WScript.Echo "Save the part to persist changes."
    
End Sub

Main
