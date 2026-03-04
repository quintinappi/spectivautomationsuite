===============================================================================
AssemblyScannerPatcher - Experimental Tool
Created: 2026-01-15
Location: Experiments/2026-01-15_AssemblyScannerPatcher/
===============================================================================

OVERVIEW
--------
Unified tool for scanning Inventor assemblies, patching iLogic rules after
renaming, and comparing before/after states to verify changes.

PURPOSE
---------
This tool addresses the critical gap where:
1. Assembly renaming/cloning modifies file names
2. iLogic rules still reference old component names
3. derived parts with external base components break
4. No verification that changes were applied correctly

This tool provides:
- PRE-WORK SCAN: Baseline analysis of assembly state
- WORK EXECUTION: Run Assembly_Renamer or Assembly_Cloner
- POST-WORK PATCH: Auto-patch iLogic rules if needed
- VERIFICATION: Compare before/after states

COMPONENTS
---------
  * AssemblyScannerPatcher.exe - Main console application
  * iLogicPatcher.dll - iLogic patching engine
  * Build_All.cmd - Master build script
  * Build_Scanner.cmd - Build scanner only
  * Build_iLogicPatcher.cmd - Build DLL only
  * README.txt - This file

REQUIREMENTS
-----------
1. .NET Framework 4.0 or higher (for vbc.exe compiler)
2. Autodesk Inventor 2025 (or compatible version)
3. iLogic Add-in installed in Inventor
4. vbc.exe in system PATH:
   - C:\Windows\Microsoft.NET\Framework\v4.0.30319
   - or: C:\Windows\Microsoft.NET\Framework64\v4.0.30319

BUILD INSTRUCTIONS
-----------------
1. Open command prompt
2. Navigate to this folder
3. Run Build_All.cmd

   C:\> cd "C:\Users\...\Experiments\2026-01-15_AssemblyScannerPatcher"
   C:\...\2026-01-15_AssemblyScannerPatcher> Build_All.cmd

This will compile both iLogicPatcher.dll and AssemblyScannerPatcher.exe

TROUBLESHOOTING BUILD ISSUES
----------------------------
If vbc.exe not found:
  - Add to PATH: set PATH=%PATH%;C:\Windows\Microsoft.NET\Framework64\v4.0.30319

If Inventor Interop not found:
  - Edit Build_iLogicPatcher.cmd
  - Change AUTODESK_INVENTOR_PATH to your Inventor installation

USAGE
-----

SCAN MODE (Before work)
-----------------------
Extracts iLogic rules, derived parts info, and assembly structure

  AssemblyScannerPatcher.exe SCAN "C:\path\assembly.iam" "C:\output\before_scan"

Output folder will contain:
  * Scan_Log.txt - Text log of scan operations
  * Scan_Summary.html - Visual summary
  * iLogic_Rules/ - Extracted iLogic rules for each document
  * Derived_Parts/Derived_Parts_Report.txt - Audit of derived parts
  * Assembly_Structure.txt - Component hierarchy

PATCH MODE (After renaming)
----------------------------
Applies mapping file to update iLogic rules with new component names

  AssemblyScannerPatcher.exe PATCH "C:\path\assembly.iam"

Requirements:
  * Assembly must be OPEN in Inventor before running
  * STEP_1_MAPPING.txt must exist in assembly's folder
  * iLogicPatcher.dll must be in same folder as EXE

COMPARE MODE (After patching)
------------------------------
Generates HTML comparison report between before/after scans

  AssemblyScannerPatcher.exe COMPARE "C:\output\before_scan" "C:\output\after_scan" "C:\output\report"

Report includes:
  * File structure changes
  * iLogic rule differences
  * Derived parts changes
  * Assembly structure differences

INTEGRATION WITH ASSEMBLY_RENAMER.VBS
------------------------------------
Add these lines to Assembly_Renamer.vbs:

  ' STEP 0: Baseline scan (BEFORE renaming)
  Dim baseScanFolder
  baseScanFolder = fso.GetParentFolderName(activeDoc.FullFileName) & "\Before_Scan"
  LogMessage "STEP 0: Running baseline scan..."
  CreateObject("WScript.Shell").Run " """ & scriptDir & "\..\Experiments\2026-01-15_AssemblyScannerPatcher\AssemblyScannerPatcher.exe"" SCAN """ & activeDoc.FullFileName & """ """ & baseScanFolder & """", 1, False

  ' ... existing renaming steps ...

  ' STEP 8.6: Patch iLogic and verify (AFTER renaming)
  Dim afterScanFolder
  afterScanFolder = fso.GetParentFolderName(activeDoc.FullFileName) & "\After_Scan"
  LogMessage "STEP 8.6: Running post-work scan and patching..."
  CreateObject("WScript.Shell").Run " """ & scriptDir & "\..\Experiments\2026-01-15_AssemblyScannerPatcher\AssemblyScannerPatcher.exe"" PATCH """ & activeDoc.FullFileName & """", 1, True
  CreateObject("WScript.Shell").Run " """ & scriptDir & "\..\Experiments\2026-01-15_AssemblyScannerPatcher\AssemblyScannerPatcher.exe"" SCAN """ & activeDoc.FullFileName & """ """ & afterScanFolder & """", 1, False

  ' Wait for scan to complete
  LogMessage "STEP 8.7: Generating comparison report..."
  WScript.Sleep 5000
  CreateObject("WScript.Shell").Run " """ & scriptDir & "\..\Experiments\2026-01-15_AssemblyScannerPatcher\AssemblyScannerPatcher.exe"" COMPARE """ & baseScanFolder & """ """ & afterScanFolder & """ """ & fso.GetParentFolderName(activeDoc.FullFileName) & "\Comparison_Report""", 1, False

INTEGRATION WITH ASSEMBLY_CLONER.VBS
------------------------------------
Same integration pattern as Assembly_Renamer.vbs

TESTING PROCEDURE
-----------------
1. Select test assembly with:
   - Both derived parts AND iLogic rules
   - Complex component hierarchy
   - Multiple subassemblies

2. Open assembly in Inventor

3. Run BEFORE scan:
   AssemblyScannerPatcher.exe SCAN "C:\test\staircase.iam" "C:\temp\staircase_before"

4. Run Assembly_Renamer to rename components
   (Use SpectivLauncher or run directly)

5. Run AFTER scan:
   AssemblyScannerPatcher.exe SCAN "C:\test\staircase.iam" "C:\temp\staircase_after"

6. Run patch:
   AssemblyScannerPatcher.exe PATCH "C:\test\staircase.iam"

7. Run FINAL scan after patching:
   AssemblyScannerPatcher.exe SCAN "C:\test\staircase.iam" "C:\temp\staircase_final"

8. Generate comparison:
   AssemblyScannerPatcher.exe COMPARE "C:\temp\staircase_before" "C:\temp\staircase_final" "C:\temp\staircase_report"

9. Review comparison report to verify:
   - All iLogic rules updated correctly
   - No broken component references
   - Derived parts properly fixed

EXPECTED BEHAVIOR
----------------
SCAN MODE:
  ✓ Connection to running Inventor instance
  ✓ Extract all iLogic rules from assembly and references
  ✓ Detect all derived parts (external vs local)
  ✓ Capture assembly component hierarchy
  ✓ Save all results to output folder

PATCH MODE:
  ✓ Read STEP_1_MAPPING.txt from assembly folder
  ✓ Parse file names and strip .ipt/.iam extensions
  ✓ Connect to iLogic add-in
  ✓ Update all iLogic rules with new component names
  ✓ Handle occurrence numbers (:1 through :50)
  ✓ Log all replacements made

COMPARE MODE:
  ✓ Generate HTML comparison report
  ✓ Highlight changed sections
  ✓ Show side-by-side differences
  ✓ Confirm all iLogic rules updated

KNOWN LIMITATIONS
----------------
1. PATCH mode requires assembly to be OPEN in Inventor
   - Cannot use Apprentice API (no iLogic access)
   - Must manually open before running

2. SCAN mode requires running Inventor instance
   - Cannot scan closed assemblies
   - Same Apprentice limitation

3. Derived parts scanning is basic
   - Complex derived part hierarchies not fully analyzed
   - Only scans immediate base components

4. iLogic patching only handles quoted string patterns
   - May miss unquoted component references
   - Occurrence number matching limited to :1-:50

FUTURE ENHANCEMENTS
-----------------
[ ] Apprentice mode for scanning closed assemblies
[ ] Comprehensive derived part chain analysis
[ - ] More sophisticated iLogic pattern matching
[ ] Direct VB.NET API for Assembly_Renamer integration
[ ] GUI application for easier testing
[ ] Integration with SpectivLauncher UI

ROLLBACK PLAN
------------
If testing fails:

1. Delete experimental folder contents
2. Restore from original vbs scripts
3. Manual iLogic patching using Fix_Derived_Parts.vbs
4. Manual comparison using file diff tools

MOVING TO PRODUCTION
--------------------
Prerequisites for experimental tool to be promoted:
1. ✓ Successful test on assembly with derived parts AND iLogic
2. ✓ Comparison report shows all iLogic rules updated correctly
3. ✓ No broken references after patching
4. ✓ Integration with Assembly_Renamer.vbs works
5. ✓ Integration with Assembly_Cloner.vbs works
6. ✓ User acceptance testing complete

Promotion steps:
1. Copy AssemblyScannerPatcher.exe and iLogicPatcher.dll to \Part_Renaming\
2. Update Launch_UI.ps1 to include new buttons
3. Update Build_Launcher.bat to auto-compile scanner
4. Move source code to Part_Renaming\ subfolder
5. Update documentation in standard folder
6. Remove from Experiments folder

CONTACT
-------
For issues or questions, refer to the main development team.

===============================================================================
