' Change_Dimension_Style.vbs - DETAILING WORKFLOW STEP 12: Change Dimension Style
' DETAILING WORKFLOW - STEP 12: Change dimension style
'
' Lists available dimension styles in the active drawing, allows user to pick one,
' then replaces all dimension styles in the document with the selected style.

Option Explicit

On Error Resume Next

Dim invApp
Set invApp = GetObject(, "Inventor.Application")
If invApp Is Nothing Then
    MsgBox "ERROR: Cannot connect to Inventor. Make sure Inventor is running.", vbCritical, "Change Dimension Style"
    WScript.Quit 1
End If

' Find an open drawing document
Dim idwDoc
Set idwDoc = Nothing
If Not invApp.ActiveDocument Is Nothing Then
    If invApp.ActiveDocument.DocumentType = 12292 Then ' kDrawingDocumentObject
        Set idwDoc = invApp.ActiveDocument
    End If
End If

If idwDoc Is Nothing Then
    ' Search all open docs
    Dim docs, i
    Set docs = invApp.Documents
    For i = 1 To docs.Count
        If docs.Item(i).DocumentType = 12292 Then
            Set idwDoc = docs.Item(i)
            Exit For
        End If
    Next
End If

If idwDoc Is Nothing Then
    MsgBox "No drawing document found. Open the IDW you want to modify and try again.", vbExclamation, "Change Dimension Style"
    WScript.Quit 1
End If

' Access the drawing's styles manager
Dim dsm
Set dsm = idwDoc.StylesManager
If Err.Number <> 0 Or dsm Is Nothing Then
    MsgBox "ERROR: Cannot get StylesManager for the drawing.", vbCritical, "Change Dimension Style"
    WScript.Quit 1
End If

Dim dStyles
Set dStyles = dsm.DimensionStyles
If dStyles.Count = 0 Then
    MsgBox "No dimension styles found in this drawing.", vbInformation, "Change Dimension Style"
    WScript.Quit 0
End If

' Build list for user
Dim listText
listText = "Available dimension styles in " & idwDoc.DisplayName & ":" & vbCrLf & vbCrLf
For i = 1 To dStyles.Count
    listText = listText & CStr(i) & ": " & dStyles.Item(i).Name & vbCrLf
Next

' Show list and ask for selection
Dim prompt, input
prompt = listText & vbCrLf & "Enter the index number (1.." & CStr(dStyles.Count) & ") of the style to apply to all dimensions, or enter the exact style name:" & vbCrLf & "(Cancel to abort)"
input = InputBox(prompt, "Select Dimension Style")
If input = "" Then
    MsgBox "No selection made. Aborting.", vbInformation, "Change Dimension Style"
    WScript.Quit 0
End If

' Resolve selection to a DimensionStyle object
Dim targetStyle
Set targetStyle = Nothing
If IsNumeric(input) Then
    Dim idx
    idx = CInt(input)
    If idx >= 1 And idx <= dStyles.Count Then
        Set targetStyle = dStyles.Item(idx)
    End If
End If

If targetStyle Is Nothing Then
    ' Try to match by name (case-insensitive)
    Dim sName
    sName = Trim(input)
    For i = 1 To dStyles.Count
        If UCase(dStyles.Item(i).Name) = UCase(sName) Then
            Set targetStyle = dStyles.Item(i)
            Exit For
        End If
    Next
End If

If targetStyle Is Nothing Then
    MsgBox "Could not find a dimension style matching your input. Aborting.", vbExclamation, "Change Dimension Style"
    WScript.Quit 1
End If

' Confirm action
Dim confirmMsg
confirmMsg = "Change ALL dimensions in " & idwDoc.DisplayName & " to use the style: '" & targetStyle.Name & "'?" & vbCrLf & vbCrLf & "This will replace references to other dimension styles in the document." & vbCrLf & vbCrLf & "Purge replaced styles from the document? (Yes will remove the old styles)"
Dim res
res = MsgBox(confirmMsg, vbYesNoCancel + vbQuestion, "Confirm Replace Dimension Styles")
If res = vbCancel Then
    MsgBox "Aborted by user.", vbInformation, "Change Dimension Style"
    WScript.Quit 0
End If

Dim purge
purge = False
If res = vbYes Then purge = True

' If target is the only style, nothing to do
If dStyles.Count = 1 Then
    MsgBox "Only one dimension style exists in the document — nothing to replace.", vbInformation, "Change Dimension Style"
    WScript.Quit 0
End If

' Build collection of styles to replace (all except the target)
Dim objColl
Set objColl = invApp.TransientObjects.CreateObjectCollection()
Dim replacedNames
replacedNames = ""
For i = 1 To dStyles.Count
    Dim st
    Set st = dStyles.Item(i)
    If st.InternalName <> targetStyle.InternalName Then
        objColl.Add st
        replacedNames = replacedNames & st.Name & ", "
    End If
Next
If Len(replacedNames) > 2 Then replacedNames = Left(replacedNames, Len(replacedNames) - 2)

' If nothing to replace, exit
If objColl.Count = 0 Then
    MsgBox "No other dimension styles to replace; nothing to do.", vbInformation, "Change Dimension Style"
    WScript.Quit 0
End If

' Perform replacement with error handling and fallback
On Error Resume Next
dsm.ReplaceStyles objColl, targetStyle, purge
If Err.Number <> 0 Then
    Dim initialErr
    initialErr = Err.Description
    Err.Clear

    ' Fallback: try a local copy of the target style
    Dim tmpName, tmpStyle
    tmpName = targetStyle.Name & "_LOCAL_TMP_" & Replace(Replace(CStr(Now), "/", "-"), ":", "-")

    On Error Resume Next
    Set tmpStyle = targetStyle.Copy(tmpName)
    If Err.Number = 0 And Not tmpStyle Is Nothing Then
        Err.Clear
        dsm.ReplaceStyles objColl, tmpStyle, purge
        If Err.Number = 0 Then
            ' Success with tmpStyle
            MsgBox "Successfully replaced all dimension styles with: " & tmpStyle.Name & vbCrLf & vbCrLf & "Replaced styles: " & replacedNames, vbInformation, "Change Dimension Style"
            WScript.Quit 0
        Else
            ' Still failing - attempt per-style replacement to identify bad styles
            Err.Clear
            Dim failedList
            failedList = ""
            Dim succeededList
            succeededList = ""

            For i = 1 To objColl.Count
                Dim singleColl, curStyle
                Set singleColl = invApp.TransientObjects.CreateObjectCollection()
                Set curStyle = objColl.Item(i)
                singleColl.Add curStyle

                On Error Resume Next
                dsm.ReplaceStyles singleColl, tmpStyle, purge
                If Err.Number <> 0 Then
                    failedList = failedList & curStyle.Name & " (" & Err.Description & "), "
                    Err.Clear
                Else
                    succeededList = succeededList & curStyle.Name & ", "
                End If
            Next

            If Len(failedList) > 2 Then failedList = Left(failedList, Len(failedList) - 2)
            If Len(succeededList) > 2 Then succeededList = Left(succeededList, Len(succeededList) - 2)

            If failedList <> "" Then
                MsgBox "Partial success:" & vbCrLf & vbCrLf & "Replaced: " & succeededList & vbCrLf & vbCrLf & "Failed: " & failedList, vbExclamation, "Change Dimension Style"
            Else
                MsgBox "Successfully replaced all dimension styles with: " & tmpStyle.Name & vbCrLf & vbCrLf & "Replaced styles: " & succeededList, vbInformation, "Change Dimension Style"
            End If
        End If
    Else
        MsgBox "ERROR: ReplaceStyles failed and could not create local copy of target style" & vbCrLf & vbCrLf & "Error: " & initialErr, vbCritical, "Change Dimension Style"
        Err.Clear
        WScript.Quit 1
    End If
Else
    ' Success on first try
    MsgBox "Successfully replaced all dimension styles with: " & targetStyle.Name & vbCrLf & vbCrLf & "Replaced styles: " & replacedNames, vbInformation, "Change Dimension Style"
End If

WScript.Quit 0
