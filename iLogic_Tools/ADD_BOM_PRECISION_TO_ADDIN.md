# Add BOM Precision Tool to Inventor Add-In

## Overview
This guide shows how to integrate the new `AssemblyBOMPrecision` tool into your existing Inventor Add-in.

## Step 1: Copy the New File

Copy `AssemblyBOMPrecision.vb` to your Add-in project folder alongside the other .vb files.

## Step 2: Update StandardAddInServer.vb

Add the new button by modifying your `StandardAddInServer.vb`:

### 2.1 Add the button variable (after line 16):

```vb
Private m_cloneButton As ButtonDefinition
Private m_fixDecimalsButton As ButtonDefinition
Private m_bomPrecisionButton As ButtonDefinition  ' <-- ADD THIS
```

### 2.2 Add the button creation in AddToUserInterface() (after line 86):

```vb
' BOM Precision button - Assembly-level tool
m_bomPrecisionButton = m_inventorApplication.CommandManager.ControlDefinitions.AddButtonDefinition(
    "Update BOM Precision", 
    "BOMPrecisionButton", 
    CommandTypesEnum.kQueryOnlyCmdType,
    m_addInSite.Parent.ClientId, 
    "Update BOM precision for all plate parts in assembly",
    "Update BOM Precision", 
    , , 
    ButtonDisplayEnum.kDisplayTextInLearningMode)

AddHandler m_bomPrecisionButton.OnExecute, AddressOf BOMPrecisionButton_OnExecute
panel.CommandControls.AddButton(m_bomPrecisionButton)
```

### 2.3 Add the cleanup in RemoveFromUserInterface() (after line 103):

```vb
If m_bomPrecisionButton IsNot Nothing Then
    m_bomPrecisionButton.Delete()
    m_bomPrecisionButton = Nothing
End If
```

### 2.4 Add the handler method (after line 141):

```vb
Private Sub BOMPrecisionButton_OnExecute()
    Try
        Dim precisionTool As New AssemblyBOMPrecision(m_inventorApplication)
        precisionTool.Execute()
    Catch ex As Exception
        MsgBox("Error updating BOM precision: " & ex.Message, MsgBoxStyle.Critical, "Error")
    End Try
End Sub
```

## Step 3: Build and Deploy

1. Build the project in Visual Studio
2. Deploy using your existing DEPLOY_NOW.bat
3. Restart Inventor

## Result

You'll see a new button in the Assembly ribbon:
- **Name**: "Update BOM Precision"
- **Location**: Same panel as "Clone Assembly" and "Fix Plate Decimals"
- **Function**: Scans assembly for plate parts and updates precision

## Differences from Existing Tools

| Tool | Scope | Method | When to Use |
|------|-------|--------|-------------|
| Fix Plate Decimals | Single part | API only | Individual part fixes |
| Update BOM Precision | Entire assembly | API + Document Settings | Batch update all plates |

## Technical Notes

### Why This Works Better

The `AssemblyBOMPrecision` tool:
1. Shows a progress dialog (user can cancel)
2. Processes parts in background
3. Uses the Document Settings command (more reliable)
4. Has retry logic built-in
5. Properly finalizes the assembly

### Command IDs Tried

The tool attempts these command IDs to open Document Settings:
- `PartDocumentSettingsCmd`
- `AppDocumentSettingsCmd`
- `PartSettingsCmd`

If none work, it falls back to Alt+D keyboard shortcut.

## Troubleshooting

### Button doesn't appear
- Check that the Add-in loaded: Tools > Add-ins > Check if enabled
- Rebuild and redeploy
- Check Inventor's Add-in Manager for errors

### Tool runs but BOM doesn't update
- Try manually refreshing the BOM view in Inventor
- Close and reopen the assembly
- The precision WAS updated, Inventor just hasn't refreshed the display

### Parts fail to process
- Check if parts are read-only
- Ensure you have write permissions
- Some parts may need manual intervention

## Alternative: Manual Integration

If you prefer not to modify the Add-in, you can still use the VBScript version:

```batch
iLogic_Tools\Launch_BOM_Precision_Robust.bat
```

The Add-in version is just more user-friendly with the progress dialog.
