' TEST_Convert_With_Arguments.vbs  
' Try to pass face as argument to convert command
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, selectSet, cmdMgr

WScript.Echo "=== CONVERT WITH ARGUMENTS ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
Set partDoc = invApp.ActiveDocument
Set compDef = partDoc.ComponentDefinition
Set selectSet = partDoc.SelectSet
Set cmdMgr = invApp.CommandManager

WScript.Echo "Part: " & partDoc.DisplayName
WScript.Echo "ComponentDef Type: " & TypeName(compDef)
WScript.Echo ""

' Step 1: Revert if needed
If TypeName(compDef) = "SheetMetalComponentDefinition" Then
    WScript.Echo "=== REVERTING TO STANDARD ==="
    
    If compDef.HasFlatPattern Then
        compDef.FlatPattern.Delete
        partDoc.Update
    End If
    
    Dim revertCmd
    Set revertCmd = cmdMgr.ControlDefinitions.Item("PartConvertToStandardPartCmd")
    
    If revertCmd.Enabled Then
        revertCmd.Execute
        WScript.Sleep 1000
        partDoc.Update
        Set compDef = partDoc.ComponentDefinition
        WScript.Echo "Reverted. New Type: " & TypeName(compDef)
    End If
    WScript.Echo ""
End If

' Step 2: Find largest face
WScript.Echo "=== FINDING LARGEST FACE ==="

Dim body, faces, face, largestFace, largestArea
largestArea = 0

Set body = compDef.SurfaceBodies.Item(1)
Set faces = body.Faces

For Each face In faces
    Dim area
    area = face.Evaluator.Area * 100
    If Err.Number = 0 And area > largestArea Then
        largestArea = area
        Set largestFace = face
    End If
    Err.Clear
Next

WScript.Echo "Largest face: " & FormatNumber(largestArea, 0) & " mm²"
WScript.Echo ""

' Step 3: Try different methods to pass face to command
WScript.Echo "=== TRYING COMMAND WITH ARGUMENTS ==="

Dim convertCmd
Set convertCmd = cmdMgr.ControlDefinitions.Item("PartConvertToSheetMetalCmd")

WScript.Echo "Command: " & convertCmd.InternalName
WScript.Echo "Enabled: " & convertCmd.Enabled
WScript.Echo ""

' Method 1: NameValueMap
WScript.Echo "Method 1: NameValueMap..."

Dim nvm
Set nvm = invApp.TransientObjects.CreateNameValueMap

' Try different parameter names
nvm.Add "BaseFace", largestFace
If Err.Number <> 0 Then
    WScript.Echo "  Add BaseFace: " & Err.Description
    Err.Clear
Else
    WScript.Echo "  Added BaseFace to NameValueMap"
End If

' Try to execute with NameValueMap
WScript.Echo ""
WScript.Echo "Trying convertCmd.Execute2(nvm)..."
convertCmd.Execute2 nvm
If Err.Number <> 0 Then
    WScript.Echo "  Execute2: " & Err.Description
    Err.Clear
End If

WScript.Sleep 500

' Method 2: Try ExecuteWithArguments
WScript.Echo ""
WScript.Echo "Method 2: ExecuteWithArguments..."

Dim args
Set args = invApp.TransientObjects.CreateObjectCollection
args.Add largestFace

' Look for ExecuteWithArguments method
' Not standard - just checking

' Method 3: Check CommandTypeDef
WScript.Echo ""  
WScript.Echo "Method 3: Check command type..."
WScript.Echo "  Type: " & TypeName(convertCmd)

' Check if it's a ControlDefinition or something else
Dim ctrlType
ctrlType = convertCmd.ControlType
If Err.Number = 0 Then
    WScript.Echo "  ControlType: " & ctrlType
Else
    Err.Clear
End If

' Method 4: Try through SelectEvents
WScript.Echo ""
WScript.Echo "Method 4: SelectEvents approach..."

' This would require InteractionEvents which need callback handlers
' Not easily doable in VBScript

' Method 5: Try SheetMetalFeatures.Add directly
WScript.Echo ""
WScript.Echo "Method 5: Direct API - SheetMetalFeatures..."

' Since part is now standard, need PartComponentDefinition
Dim smFeats
Set smFeats = compDef.Features.SheetMetalFeatures
If Err.Number <> 0 Then
    WScript.Echo "  SheetMetalFeatures: " & Err.Description
    Err.Clear
Else
    WScript.Echo "  SheetMetalFeatures Type: " & TypeName(smFeats)
End If

' Try ConvertToSheetMetalFeatures on standard part
Set smFeats = compDef.Features.ConvertToSheetMetalFeatures
If Err.Number <> 0 Then
    WScript.Echo "  ConvertToSheetMetalFeatures: " & Err.Description
    Err.Clear  
Else
    WScript.Echo "  ConvertToSheetMetalFeatures Type: " & TypeName(smFeats)
    WScript.Echo "  Count: " & smFeats.Count
    
    ' Try CreateDefinition
    Dim cstDef
    Set cstDef = smFeats.CreateDefinition(largestFace)
    If Err.Number <> 0 Then
        WScript.Echo "  CreateDefinition(face): " & Err.Description
        Err.Clear
        
        Set cstDef = smFeats.CreateDefinition()
        If Err.Number <> 0 Then
            WScript.Echo "  CreateDefinition(): " & Err.Description
            Err.Clear
        Else
            WScript.Echo "  CreateDefinition() Type: " & TypeName(cstDef)
            
            ' Try to set BaseFace
            If TypeName(cstDef) <> "Empty" And Not cstDef Is Nothing Then
                WScript.Echo "  Setting BaseFace..."
                Set cstDef.BaseFace = largestFace
                If Err.Number <> 0 Then
                    WScript.Echo "    Error: " & Err.Description
                    Err.Clear
                Else
                    WScript.Echo "    BaseFace SET!"
                    
                    ' Try to Add
                    WScript.Echo "  Adding feature..."
                    Dim newFeat
                    Set newFeat = smFeats.Add(cstDef)
                    If Err.Number <> 0 Then
                        WScript.Echo "    Error: " & Err.Description
                        Err.Clear
                    Else
                        WScript.Echo "    Feature added! Type: " & TypeName(newFeat)
                    End If
                End If
            End If
        End If
    Else
        WScript.Echo "  CreateDefinition(face) Type: " & TypeName(cstDef)
    End If
End If

WScript.Echo ""
WScript.Echo "=== FINAL STATE ==="
partDoc.Update
Set compDef = partDoc.ComponentDefinition
WScript.Echo "ComponentDef Type: " & TypeName(compDef)
WScript.Echo "SubType: " & partDoc.SubType

WScript.Echo ""
WScript.Echo "=== DONE ==="
