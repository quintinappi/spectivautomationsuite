# Assembly Cloner Add-In for Autodesk Inventor

This add-in allows you to clone Inventor assemblies and automatically patch iLogic rules to update references to the new part names.

## Prerequisites

- Autodesk Inventor 2026
- .NET Framework 4.8.1
- Visual Studio 2022 (for building)

## Building the Add-In

1. Open `AssemblyClonerAddIn.sln` in Visual Studio.
2. Ensure the references to Inventor DLLs are correct (update paths if Inventor is installed elsewhere).
3. Build the solution in Release mode.

## Deploying the Add-In

1. Copy `AssemblyClonerAddIn.dll` and `AssemblyClonerAddIn.addin` to the Inventor AddIns folder:
   - Default: `C:\Users\[username]\AppData\Roaming\Autodesk\Inventor 2026\Addins\`
2. Restart Inventor.
3. The add-in should appear in the Assembly tab > Manage panel.

## Using the Add-In

1. Open an assembly in Inventor.
2. Click the "Clone Assembly & Patch iLogic" button.
3. Enter a new name for the assembly.
4. The add-in will save the assembly and parts with the new names and update iLogic rules.

## Notes

- This is a basic implementation. Enhance the patching logic as needed for your specific iLogic code patterns.
- Test thoroughly before using in production.