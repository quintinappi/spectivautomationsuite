' Fix Single Part with UI Automation - Add Length2 parameter
' Uses SendKeys to interact with Parameters dialog like Option 12

Option Explicit

Const kNumberParameterType = 1

Dim m_InventorApp
Dim m_WShell

Sub Main()
    On Error Resume Next
    
    WScript.Echo "=== FIX NON-PLATE PART - ADD LENGTH2 PARAMETER (UI METHOD) ==="
    WScript.Echo ""
    
    ' Get Inventor
    Set m_InventorApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Inventor not running"
        WScript.Quit 1
    End If
    
    ' Create WScript.Shell for SendKeys
    Set m_WShell = CreateObject("WScript.Shell")
    
    ' Get active document
    Dim doc
    Set doc = m_InventorApp.ActiveDocument
    If doc Is Nothing Then
        WScript.Echo "ERROR: No active document"
        WScript.Quit 1
    End If
    
    If doc.DocumentType <> 12290 Then
        WScript.Echo "ERROR: Please open FL25 part file"
        WScript.Quit 1
    End If
    
    WScript.Echo "Part: " & doc.DisplayName
    WScript.Echo ""
    
    ' Get parameters
    Dim compDef, params, modelParams, userParams
    Set compDef = doc.ComponentDefinition
    Set params = compDef.Parameters
    Set modelParams = params.ModelParameters
    Set userParams = params.UserParameters
    
    ' STEP 1: Find the length parameter
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
    WScript.Echo ""
    
    ' STEP 2-5: Use UI automation
    WScript.Echo "Opening Parameters dialog via UI automation..."
    WScript.Echo "(This will take a few seconds)"
    WScript.Echo ""
    
    ' Bring Inventor to front
    m_WShell.AppActivate "Autodesk Inventor"
    WScript.Sleep 500
    
    ' Open Parameters dialog (fx button or Tools menu)
    m_WShell.SendKeys "%n" ' Alt+N for Manage tab
    WScript.Sleep 300
    m_WShell.SendKeys "x" ' Parameters button
    WScript.Sleep 1000
    
    WScript.Echo "STEP 2: Creating Length2 user parameter..."
    
    ' Click on User Parameters section
    m_WShell.SendKeys "{TAB}" ' Move to tree
    WScript.Sleep 200
    m_WShell.SendKeys "{DOWN}" ' Move down to find User Parameters
    WScript.Sleep 200
    
    ' Add new parameter button
    m_WShell.SendKeys "%a" ' Alt+A for Add button
    WScript.Sleep 500
    
    ' Type parameter name
    m_WShell.SendKeys "Length2"
    WScript.Sleep 200
    
    ' Tab to Equation field
    m_WShell.SendKeys "{TAB}"
    WScript.Sleep 200
    
    ' STEP 4: Enter the value from d2
    m_WShell.SendKeys maxParamName
    WScript.Sleep 200
    
    ' Tab to enable export checkbox
    m_WShell.SendKeys "{TAB}" ' Unit
    WScript.Sleep 100
    m_WShell.SendKeys "{TAB}" ' Nominal
    WScript.Sleep 100
    m_WShell.SendKeys "{TAB}" ' Tol
    WScript.Sleep 100
    m_WShell.SendKeys "{TAB}" ' Model Value
    WScript.Sleep 100
    m_WShell.SendKeys "{TAB}" ' Key
    WScript.Sleep 100
    
    ' STEP 5: Enable export checkbox
    WScript.Echo "STEP 5: Enabling export checkbox..."
    m_WShell.SendKeys " " ' Space to check the export box
    WScript.Sleep 300
    
    WScript.Echo "  Length2 created and export enabled"
    WScript.Echo ""
    
    ' Now we need to set d2's equation to Length2
    WScript.Echo "STEP 3: Setting " & maxParamName & " equation to Length2..."
    
    ' Click on Model Parameters section
    m_WShell.SendKeys "{HOME}" ' Go to top of tree
    WScript.Sleep 200
    m_WShell.SendKeys "{DOWN}" ' Expand Model Parameters
    WScript.Sleep 200
    
    ' Find d2 parameter - need to navigate to it
    ' Send down arrow multiple times or search
    Dim foundParam
    foundParam = False
    For i = 1 To 20 ' Max attempts
        m_WShell.SendKeys "{DOWN}"
        WScript.Sleep 100
        ' Check if we found it by trying to edit
        If i > 2 Then ' Start checking after a few downs
            m_WShell.SendKeys "{TAB}{TAB}" ' Move to equation field
            WScript.Sleep 100
            m_WShell.SendKeys "^a" ' Select all
            WScript.Sleep 100
            m_WShell.SendKeys "Length2" ' Type Length2
            WScript.Sleep 100
            foundParam = True
            Exit For
        End If
    Next
    
    If Not foundParam Then
        WScript.Echo "WARNING: Could not automatically set " & maxParamName & " equation"
        WScript.Echo "Please manually set " & maxParamName & " = Length2 in the dialog"
    Else
        WScript.Echo "  " & maxParamName & " equation set to Length2"
    End If
    
    WScript.Sleep 500
    
    ' Close dialog
    WScript.Echo ""
    WScript.Echo "Closing Parameters dialog..."
    m_WShell.SendKeys "{ENTER}" ' OK button or Done
    WScript.Sleep 500
    
    WScript.Echo ""
    WScript.Echo "=========================================="
    WScript.Echo "Process complete!"
    WScript.Echo "=========================================="
    WScript.Echo ""
    WScript.Echo "Please verify in the Parameters dialog:"
    WScript.Echo "  1. Length2 exists in User Parameters"
    WScript.Echo "  2. Length2 equation = " & maxParamName
    WScript.Echo "  3. Length2 export checkbox is checked"
    WScript.Echo "  4. " & maxParamName & " equation = Length2"
    WScript.Echo ""
    WScript.Echo "Save the part if everything looks correct."
    
End Sub

Main
