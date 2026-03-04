' TEST_Direct_API_Convert.vbs
' Convert to sheet metal using direct API, not command dialog
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, cmdMgr

WScript.Echo "=== DIRECT API SHEET METAL CONVERSION ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
Set partDoc = invApp.ActiveDocument
Set compDef = partDoc.ComponentDefinition
Set cmdMgr = invApp.CommandManager

WScript.Echo "Part: " & partDoc.DisplayName

' Step 1: Check if derived and break link
WScript.Echo ""
WScript.Echo "=== STEP 1: CHECK FOR DERIVED LINK ==="

Dim feat, isDerived
isDerived = False

For Each feat In compDef.Features
    If InStr(feat.Name, "::") > 0 Then
        isDerived = True
        WScript.Echo "Found derived reference: " & feat.Name
        Exit For
    End If
    Err.Clear
Next

If isDerived Then
    WScript.Echo "Breaking derived link..."
    
    Dim breakCmd
    Set breakCmd = cmdMgr.ControlDefinitions.Item("PartBreakLinkDerivedPartCtxCmd")
    
    If Not breakCmd Is Nothing And breakCmd.Enabled Then
        breakCmd.Execute
        WScript.Sleep 500
        partDoc.Update
        WScript.Echo "Link broken!"
    Else
        WScript.Echo "Break link command not available"
    End If
Else
    WScript.Echo "Not a derived part"
End If

' Refresh compDef
Set compDef = partDoc.ComponentDefinition

' Step 2: Check current type
WScript.Echo ""
WScript.Echo "=== STEP 2: CHECK PART TYPE ==="

Dim subType
subType = ""
subType = compDef.SubType
WScript.Echo "Current SubType: " & subType

Dim isSheetMetal
isSheetMetal = (subType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
WScript.Echo "Is Sheet Metal: " & isSheetMetal

' Step 3: Convert to sheet metal via SubType
WScript.Echo ""
WScript.Echo "=== STEP 3: CONVERT TO SHEET METAL ==="

If Not isSheetMetal Then
    WScript.Echo "Setting SubType to Sheet Metal..."
    
    compDef.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
    
    If Err.Number = 0 Then
        WScript.Echo "SubType changed!"
        partDoc.Update
        
        ' Verify
        Set compDef = partDoc.ComponentDefinition
        WScript.Echo "New SubType: " & compDef.SubType
    Else
        WScript.Echo "Failed to change SubType: " & Err.Description
        Err.Clear
    End If
End If

' Step 4: Set thickness
WScript.Echo ""
WScript.Echo "=== STEP 4: SET THICKNESS ==="

' Get the smallest dimension as thickness
Dim rBox, dimX, dimY, dimZ, thickness
Set rBox = compDef.RangeBox

dimX = Abs(rBox.MaxPoint.X - rBox.MinPoint.X)
dimY = Abs(rBox.MaxPoint.Y - rBox.MinPoint.Y)
dimZ = Abs(rBox.MaxPoint.Z - rBox.MinPoint.Z)

' Find smallest
thickness = dimX
If dimY < thickness Then thickness = dimY
If dimZ < thickness Then thickness = dimZ

WScript.Echo "Detected thickness: " & FormatNumber(thickness * 10, 1) & " mm"

' Set thickness if sheet metal
Dim smDef
Set smDef = compDef

On Error Resume Next
smDef.Thickness.Value = thickness
If Err.Number = 0 Then
    WScript.Echo "Thickness set!"
Else
    WScript.Echo "Could not set thickness: " & Err.Description
    Err.Clear
End If

partDoc.Update

' Step 5: Create flat pattern
WScript.Echo ""
WScript.Echo "=== STEP 5: CREATE FLAT PATTERN ==="

WScript.Echo "HasFlatPattern: " & smDef.HasFlatPattern

If Not smDef.HasFlatPattern Then
    WScript.Echo "Creating flat pattern..."
    
    smDef.Unfold
    
    If Err.Number = 0 Then
        WScript.Echo "Flat pattern created!"
    Else
        WScript.Echo "Unfold failed: " & Err.Description
        Err.Clear
    End If
    
    partDoc.Update
End If

' Step 6: Check result
WScript.Echo ""
WScript.Echo "=== FINAL RESULT ==="

Set compDef = partDoc.ComponentDefinition
WScript.Echo "SubType: " & compDef.SubType
WScript.Echo "HasFlatPattern: " & compDef.HasFlatPattern

If compDef.HasFlatPattern Then
    Dim fp
    Set fp = compDef.FlatPattern
    WScript.Echo ""
    WScript.Echo "Flat pattern dimensions:"
    WScript.Echo "  Length: " & FormatNumber(fp.Length * 10, 1) & " mm"
    WScript.Echo "  Width: " & FormatNumber(fp.Width * 10, 1) & " mm"
    
    If fp.Length * 10 > 50 And fp.Width * 10 > 50 Then
        WScript.Echo ""
        WScript.Echo "*** SUCCESS! CORRECT ORIENTATION! ***"
    Else
        WScript.Echo ""
        WScript.Echo "Edge view - wrong orientation"
    End If
End If

WScript.Echo ""
WScript.Echo "=== DONE ==="
