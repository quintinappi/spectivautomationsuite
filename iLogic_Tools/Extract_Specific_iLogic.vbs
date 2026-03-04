Option Explicit

' ==============================================================================
' SPECIFIC iLOGIC EXTRACTOR - Extract iLogic from Beam_Factory and Endplate_Factory
' ==============================================================================
' This script opens specific .ipt files and extracts their iLogic rules

Const ILOGIC_ADDIN_GUID = "{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}"

Dim g_fso
Set g_fso = CreateObject("Scripting.FileSystemObject")

' Files to analyze
Dim targetFiles
targetFiles = Array( _
    "C:\Users\Quintin\Documents\Spectiv\3. Working\22. DMS plant\REF\Beam_Factory.ipt", _
    "C:\Users\Quintin\Documents\Spectiv\3. Working\22. DMS plant\REF\Endplate_Factory.ipt" _
)

' Output folder
Dim outputFolder
outputFolder = "C:\Users\Quintin\Documents\Spectiv\3. Working\22. DMS plant\REF\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\iLogic_Tools\iLogic_Export"

If Not g_fso.FolderExists(outputFolder) Then
    g_fso.CreateFolder outputFolder
End If

' Connect to Inventor
WScript.Echo "Connecting to Inventor..."
Dim invApp
On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")

If Err.Number <> 0 Or invApp Is Nothing Then
    ' Try to start Inventor
    Err.Clear
    Set invApp = CreateObject("Inventor.Application")
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Could not connect to or start Inventor!"
        WScript.Quit 1
    End If
    invApp.Visible = True
End If
Err.Clear
On Error GoTo 0

WScript.Echo "SUCCESS: Connected to Inventor"

' Get iLogic Add-In
Dim iLogicAddin
Set iLogicAddin = GetILogicAddin(invApp)

If iLogicAddin Is Nothing Then
    WScript.Echo "ERROR: Could not find iLogic Add-In!"
    WScript.Quit 1
End If

' Process each file
Dim filePath
For Each filePath In targetFiles
    If g_fso.FileExists(filePath) Then
        WScript.Echo ""
        WScript.Echo "=========================================="
        WScript.Echo "Processing: " & g_fso.GetFileName(filePath)
        WScript.Echo "=========================================="

        Call ExtractILogicFromFile(invApp, iLogicAddin, filePath, outputFolder)
    Else
        WScript.Echo "WARNING: File not found - " & filePath
    End If
Next

WScript.Echo ""
WScript.Echo "=========================================="
WScript.Echo "EXTRACTION COMPLETE!"
WScript.Echo "Output folder: " & outputFolder
WScript.Echo "=========================================="

' ==============================================================================
' FUNCTIONS
' ==============================================================================

Function GetILogicAddin(invApp)
    On Error Resume Next

    Dim addins
    Set addins = invApp.ApplicationAddIns

    Dim addin
    Set addin = addins.ItemById(ILOGIC_ADDIN_GUID)

    If Err.Number <> 0 Or addin Is Nothing Then
        ' Try to find by name
        Dim tempAddin
        For Each tempAddin In addins
            If InStr(1, tempAddin.DisplayName, "iLogic", vbTextCompare) > 0 Then
                Set addin = tempAddin
                Exit For
            End If
        Next
    End If

    If Not addin Is Nothing Then
        addin.Activate
        Set GetILogicAddin = addin
    Else
        Set GetILogicAddin = Nothing
    End If

    Err.Clear
End Function

Sub ExtractILogicFromFile(invApp, iLogicAddin, filePath, outputFolder)
    On Error Resume Next

    ' Open the document
    WScript.Echo "Opening document..."
    Dim doc
    Set doc = invApp.Documents.Open(filePath, False)

    If Err.Number <> 0 Or doc Is Nothing Then
        WScript.Echo "ERROR: Could not open document - " & Err.Description
        Err.Clear
        Exit Sub
    End If

    WScript.Echo "SUCCESS: Document opened"

    ' Get iLogic automation
    Dim iLogicAuto
    Set iLogicAuto = iLogicAddin.Automation

    If iLogicAuto Is Nothing Then
        WScript.Echo "ERROR: Could not get iLogic Automation interface"
        doc.Close True
        Exit Sub
    End If

    ' Try to get rules collection
    Dim rules
    Set rules = Nothing

    ' Method 1: Try Rules property
    Err.Clear
    Set rules = iLogicAuto.Rules(doc)

    If Err.Number <> 0 Then
        WScript.Echo "WARNING: Could not access Rules collection - " & Err.Description
        Err.Clear
    End If

    ' Check AttributeSets as fallback
    WScript.Echo "Scanning AttributeSets for iLogic rules..."
    Dim attrSets
    Set attrSets = doc.AttributeSets

    Dim ruleCount
    ruleCount = 0

    Dim attrSet
    For Each attrSet In attrSets
        Dim setName
        setName = attrSet.Name

        ' Look for iLogic-related attribute sets
        If Left(setName, 10) = "iLogicRule" Or _
           Left(setName, 6) = "iLogic" Or _
           InStr(1, setName, "Rule", vbTextCompare) > 0 Then

            WScript.Echo "  Found AttributeSet: " & setName

            ' Try to extract rule information
            Dim attr
            Dim ruleText
            ruleText = ""
            Dim ruleName
            ruleName = setName

            For Each attr In attrSet
                WScript.Echo "    - Attribute: " & attr.Name & " (Type: " & TypeName(attr.Value) & ")"

                ' Look for code/text attributes - check for iLogicRuleText specifically
                If LCase(attr.Name) = "ilogicruletext" Or _
                   LCase(attr.Name) = "text" Or _
                   LCase(attr.Name) = "code" Or _
                   LCase(attr.Name) = "rule" Or _
                   LCase(attr.Name) = "ruletext" Then
                    Err.Clear
                    ruleText = CStr(attr.Value)
                    If Err.Number = 0 And Len(ruleText) > 0 Then
                        WScript.Echo "      -> Code attribute found! (" & Len(ruleText) & " characters)"
                    End If
                    Err.Clear
                End If

                If LCase(attr.Name) = "name" Or LCase(attr.Name) = "rulename" Then
                    ruleName = CStr(attr.Value)
                End If
            Next

            ' For iLogicRule_ sets, extract the rule name from the set name
            If Left(setName, 11) = "iLogicRule_" Then
                ruleName = Mid(setName, 12)
            End If

            ' Save the rule if we found code
            If ruleText <> "" Then
                ruleCount = ruleCount + 1
                Call SaveRuleToFile(g_fso.GetBaseName(filePath), ruleName, ruleText, outputFolder)
                WScript.Echo "  EXTRACTED: " & ruleName & " (" & Len(ruleText) & " characters)"
            End If
        End If
    Next

    ' Also try to enumerate rules via Rules collection if available
    If Not rules Is Nothing Then
        WScript.Echo "Extracting via Rules collection..."

        Dim rule
        For Each rule In rules
            Err.Clear

            Dim name
            name = rule.Name

            Dim text
            text = rule.Text

            If Err.Number = 0 And text <> "" Then
                ruleCount = ruleCount + 1
                Call SaveRuleToFile(g_fso.GetBaseName(filePath), name, text, outputFolder)
                WScript.Echo "  EXTRACTED: " & name & " (" & Len(text) & " characters)"
            End If
        Next
    End If

    WScript.Echo ""
    WScript.Echo "TOTAL RULES EXTRACTED: " & ruleCount

    ' Extract parameters/properties as well
    WScript.Echo ""
    WScript.Echo "Extracting parameters..."
    Call ExtractParameters(doc, g_fso.GetBaseName(filePath), outputFolder)

    ' Close document
    doc.Close True
    Err.Clear
End Sub

Sub SaveRuleToFile(baseName, ruleName, ruleText, outputFolder)
    On Error Resume Next

    Dim safeRuleName
    safeRuleName = CleanFileName(ruleName)

    Dim filePath
    filePath = outputFolder & "\" & baseName & "_Rule_" & safeRuleName & ".vb"

    Dim outFile
    Set outFile = g_fso.CreateTextFile(filePath, True, False)

    outFile.WriteLine "' =============================================="
    outFile.WriteLine "' iLogic Rule Export"
    outFile.WriteLine "' =============================================="
    outFile.WriteLine "' Source File: " & baseName
    outFile.WriteLine "' Rule Name: " & ruleName
    outFile.WriteLine "' Exported: " & Now()
    outFile.WriteLine "' =============================================="
    outFile.WriteLine ""
    outFile.WriteLine ruleText
    outFile.Close

    Err.Clear
End Sub

Sub ExtractParameters(doc, baseName, outputFolder)
    On Error Resume Next

    Dim outPath
    outPath = outputFolder & "\" & baseName & "_Parameters.txt"

    Dim outFile
    Set outFile = g_fso.CreateTextFile(outPath, True, False)

    outFile.WriteLine "=============================================="
    outFile.WriteLine "PARAMETERS EXTRACTION"
    outFile.WriteLine "=============================================="
    outFile.WriteLine "Source File: " & baseName
    outFile.WriteLine "Exported: " & Now()
    outFile.WriteLine "=============================================="
    outFile.WriteLine ""

    ' User parameters
    outFile.WriteLine "USER PARAMETERS:"
    outFile.WriteLine "-------------------------------------------"

    Dim params
    Set params = doc.ComponentDefinition.Parameters.UserParameters

    Dim param
    For Each param In params
        Err.Clear
        outFile.WriteLine "  " & param.Name & " = " & param.Expression & " [" & param.Units & "]"
        If param.Comment <> "" Then
            outFile.WriteLine "    Comment: " & param.Comment
        End If
    Next

    outFile.WriteLine ""
    outFile.WriteLine "MODEL PARAMETERS:"
    outFile.WriteLine "-------------------------------------------"

    Set params = doc.ComponentDefinition.Parameters.ModelParameters

    For Each param In params
        Err.Clear
        outFile.WriteLine "  " & param.Name & " = " & param.Expression
    Next

    ' iProperties
    outFile.WriteLine ""
    outFile.WriteLine "iPROPERTIES:"
    outFile.WriteLine "-------------------------------------------"

    Dim propSets
    Set propSets = doc.PropertySets

    Dim propSet
    For Each propSet In propSets
        outFile.WriteLine ""
        outFile.WriteLine "[" & propSet.DisplayName & "]"

        Dim prop
        For Each prop In propSet
            Err.Clear
            Dim propVal
            propVal = prop.Value
            If Err.Number = 0 Then
                outFile.WriteLine "  " & prop.DisplayName & " = " & propVal
            End If
        Next
    Next

    outFile.Close

    WScript.Echo "  Parameters saved to: " & g_fso.GetFileName(outPath)

    Err.Clear
End Sub

Function CleanFileName(name)
    Dim result
    result = name

    result = Replace(result, "\", "_")
    result = Replace(result, "/", "_")
    result = Replace(result, ":", "_")
    result = Replace(result, "*", "_")
    result = Replace(result, "?", "_")
    result = Replace(result, """", "_")
    result = Replace(result, "<", "_")
    result = Replace(result, ">", "_")
    result = Replace(result, "|", "_")
    result = Replace(result, " ", "_")

    CleanFileName = result
End Function
