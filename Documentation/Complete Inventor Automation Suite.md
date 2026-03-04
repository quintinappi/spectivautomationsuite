### **Complete Inventor Automation Suite Workflow**

### **Pre-requisites : Must have Autodesk Apprentice Installed**



###### **PART RENAMING WORKFLOW (Steps 1-10)**

###### **Purpose: Rename parts in assemblies with sequential numbering and update all references.**



**Step 1: Ensure prefix is ready (scan registry)**

Run Registry\_Manager.vbs to familiarize with current registry for your prefix *( if already renamed )*

Run Smart\_Prefix\_Scanner.vbs to scan existing assemblies and update registry counters *( if already renamed )*

This prevents duplicate numbering when adding new assemblies to existing projects *( if already renamed )*



**Step 2: Open assembly in Inventor**

Open your target assembly (.iam file) in Autodesk Inventor

Ensure all sub-assemblies are loaded, and main assemblies is loaded fully.

Open BOM and ensure descriptions field are updated as : ( as scripts detect these keywords in order to group naming of parts )

&nbsp;	PL 6mm S355JR ( for plates )

&nbsp;	UB203x133x32 ( for beams )

&nbsp;	UC152x152x32 ( for columns )

&nbsp;	LPL 12mm VRN400 ( for liners )

&nbsp;	P ( pipes )

&nbsp;	A70x70x6 ( for angles )

Drag Stock Number next to Description and enter this formula FOR ALL PLATES =PL<thickness> S355JR - <sheet metal width> x <sheet metal length> , to auto calculate plate description for parts list.



**Step 3: Run renamer from start** 

Execute Assembly\_Renamer.vbs (main renaming tool)

Always save the mapping text file in the source root ( if asked for any tools )

***IMPORTANT:*** Monitor Inventor closely for popup dialogs during execution



**Step 4: Monitor Inventor for popups**

Watch for any dialog boxes that appear during the renaming process

Click OK or provide required input when prompted



**Step 5: Verify groupings**

Review the automatic component groupings displayed by the renamer

Ensure similar parts are grouped correctly, if at any stage there is a discrepancy, close the terminal immediately and inspect your BOM to find the discrepancy so grouping can be correct.



**Step 6: Define naming schemes**

For each component group, the script auto defines naming schemes (e.g., "NCRH01-000-PL{N}") , but you may change it manually ( NOT {N} AS THIS IS FORMULATED ), but strongly not advised. 

Expected description format: "PL 6mm S355JR" or similar stock specifications

Naming format: PREFIX-GROUP{N}.ipt (e.g., NCRH01-000-PL173.ipt)



**Step 7: Monitor for popups during copying \& rename process**

Watch for popup dialogs during the heritage-based copying process

Respond to any prompts that appear



**Step 8: Update assembly references**

The renamer automatically updates all assembly references

Original files are kept for safety

Important: The cleanup can be run after this step to remove any unused parts in assembly ( have assembly open to scan for this part ), this moves all unused parts into a folder keeping folder structure.



**Step 9: Update IDW files**

The renamer should automatically detects and updates all IDW drawings in the directory

Mapping file is saved for external updates



**Step 10: Run idw fixer and iLogic patcher next ( if step 9 fails )**

Execute Emergency\_IDW\_Fixer.vbs to fix any missed IDW references ( this is the reason for site\_1\_mapping.txt , when asked point it to this )

Execute iLogic\_Patcher.vbs to update iLogic rules with new component named ( for any assemblies that contain derived parts, for complex assemblies this can be done safely , manually )

Proceed to DETAILING WORKFLOW (STEP 2)



###### **DETAILING WORKFLOW (Steps 1-15)**

###### **Purpose: Prepare drawings for production with proper formatting and export.**



**Step 1: Open assembly in Inventor**

Ensure your renamed assembly is open in Inventor

All IDW files should be updated from PART RENAMING step



**Step 2: Run sheet metal converter ( IF THERE ARE ANY PLATES IN YOUR ASSEMBLY , at this point your description should be correct and all plates should be named -PL{N} )**

Execute Sheet\_Metal\_Converter.vbs

Converts PL parts to sheet metal in assemblies

Prompts for flat face selection and applies defaults

Important: During this process it is imperative to monitor Inventor, all plate parts will be opened automatically, and you are required to click the face you want flat patterened. 

&nbsp;		If any parts flatten in wrong direction, you must fix this manually by right-clicking on flat pattern -> edit flat pattern definition and select the right edge.

Open Parameters for each plate and tick the Export checkbox of Thickness.



**Step 3: Add Length2 property** 

This step ensures the length parameter for each non-plate part with length is exported for use in parts list formula

Execute Fix\_All\_Non\_Plate\_Parts.vbs

Adds Length2 parameter to all non-plate parts missing Length

For every part done in this step , it is imperative to re-open, go to Parameters and tick the Export checkbox on Length2 user defined property for use in parts list formula.



**Step 4: Enable export for non-plate length properties**

Execute Length\_Parameter\_Exporter.vbs

Enables export for Length parameter on non-plate parts ( this is also for use in parts list formulas )



**Step 5: Refresh BOM Precision**

This step runs through the assembly and opens and updates the precision of all plates to display correctly in BOM ( once started DO NOT TOUCH keyboard or mouse until finished ) as this is a UI automation.



#### ***After Detailing of Assemblies and Parts***



**Step 5: Create GA PART LIST \& DETAIL OF PARTS PART LIST ( use  BEAM PARTS LIST template for all sheets )**					

**Open your idw and manually place parts lists on all GA sheets, for individual parts, run below script, it will ask you which page's assembly you want to check.**

After checking the assembly parts, it will look at the page you are currently viewing and place a parts list only containing all the parts you have detailed on that page.

Execute Create\_Sheet\_Parts\_List.vbs

Creates parts list containing only components visible on current sheet



**Step 6: Change balloon style ( if needed )**

Execute Change\_Balloon\_Style.vbs

Updates balloon styles in drawings to match standards



**Step 7: Change dimension style ( if needed )**

Execute Change\_Dimension\_Style.vbs

Updates dimension styles in drawings to match standards



**Step 8: Clean up unused parts ( as mentioned above )**

Execute Unused\_Part\_Finder.vbs

Moves unused IPT files to backup folder for cleanup



**Step 9: Update titles to Pentalin style**

Execute Title\_Updater.vbs

Updates base view titles with exact format requirements:

PARTS: NCHR01-000-PL1 / SCALE 1:5

ASSEMBLIES: NCHR01-000-BA1 / 7-OFF REQ'D / SCALE 1:20





**Step 10: Find Missing Detailed Parts**

Run this script after everything is done and ready for export, this script will tell you if you missed any parts for detailing



**Step 11: Export IDW sheets to PDF**

Execute Export\_IDW\_Sheets\_to\_PDF.vbs

Exports each sheet to separate PDF with correct numbering

All Colors As Black enabled





**Steps 13-17: Additional detailing steps**

Complete any remaining drawing annotations and formatting

Verify all dimensions, notes, and callouts are correct

Final review of drawing package

Archive completed drawings

Update project documentation



##### **Additional Workflow Tools: Assembly Cloner, Prefix Cloner, and Part Cloner**



###### **ASSEMBLY CLONER WORKFLOW**

###### **Purpose: Create a complete isolated copy of an entire assembly hierarchy with all sub-assemblies, parts, and drawings.**



**Step 1: Open source assembly in Inventor**

Open the assembly you want to clone in Autodesk Inventor

Ensure all sub-assemblies are loaded and accessible



**Step 2: Run Assembly Cloner**

Execute Assembly\_Cloner.vbs

IMPORTANT: Monitor Inventor for popup dialogs during execution



**Step 3: Select destination folder**

Choose a new folder location for the cloned assembly ( do NOT move it out of this location until you have run all idw fixers and happy with fixes, then you may move it )

The tool will create a fully isolated copy with no cross-references to originals



**Step 4: Choose cloning options**

Decide whether to rename parts during cloning:

**Without renaming:** Exact copy of all files

**With renaming:** Apply heritage-based naming schemes to all parts



**Step 5: If renaming selected - configure naming schemes**

Review automatic component groupings

Define naming schemes for each group (e.g., "NCRH02-000-PL{N}")

Expected description format: "PL 6mm S355JR" or similar



**Step 6: Monitor for popups during copying**

Watch for any dialog boxes during the file copying and reference updating process

The tool recursively copies:

Main assembly (.iam)

All sub-assemblies

All parts (.ipt)

All drawings (.idw)



**Step 7: Verify cloned assembly**

Check that all references are updated to local copies

Open the cloned assembly to ensure it loads correctly

Review the generated STEP\_1\_MAPPING.txt for reference tracking

Run the idw fixer and point it to this mapping file to ensure idw's references are updated before moving out of this folder. 



###### **PREFIX CLONER WORKFLOW**

###### **Purpose: Clone an assembly but change only the prefix portion of filenames while keeping part suffixes intact.**

Use Case: For copying an entire model to another section of the plant using the same model , any modifications afterwards will only affect said model



**Step 1: Open source assembly in Inventor**

Open the assembly you want to clone in Autodesk Inventor

Ensure all files are accessible and fully loaded



**Step 2: Run Prefix Cloner**

Execute Prefix\_Cloner.vbs

The tool will automatically scan files to detect the common prefix



**Step 3: Confirm old prefix**

**Important:** Review the detected prefix (e.g., "N1SCR04-780-")

Confirm or manually enter the old prefix to replace



**Step 4: Enter new prefix**

**Important:** Specify the new prefix (e.g., "N2SCR04-780-")

The tool keeps all part suffixes unchanged



**Step 5: Select destination folder**

Choose a new folder location for the cloned assembly

Example transformation:

N1SCR04-780-B1.ipt → N2SCR04-780-B1.ipt

N1SCR04-780-CH5.ipt → N2SCR04-780-CH5.ipt

Once again, DO NOT MOVE the model out of this folder until you have run the idw fixer script and confirmed all references are updated. 



**Step 6: Monitor for popups during copying**

Watch for dialog boxes during file copying

The tool copies:

Main assembly (.iam)

All sub-assemblies

All parts (.ipt)

All drawings (.idw)



**Step 7: Verify prefix changes**

Check that all references are updated with new prefix

Open cloned assembly to ensure it loads correctly

Review STEP\_1\_MAPPING.txt for traceability

Once again, DO NOT MOVE the model out of this folder until you have run the idw fixer script and confirmed all references are updated.





###### **PART CLONER WORKFLOW**

###### **Purpose: Create a copy of a single part file to a new location.**



**Step 1: Open source part in Inventor**

Open the individual part (.ipt) you want to clone in Autodesk Inventor



**Step 2: Run Part Cloner**

Execute Part\_Cloner.vbs

The tool detects the currently open part automatically



**Step 3: Select destination folder**

Choose a new folder location for the cloned part

The tool will copy the .ipt file to the selected location



**Step 4: Choose renaming option**

Decide whether to rename the part during cloning:

Keep original name: Exact copy

Rename: Specify new filename



**Step 5: Review part properties (optional)**

The tool displays iProperties of the part for verification

Useful for confirming part details before cloning



**Step 6: Verify cloned part**

Check that the part file was copied successfully

Open the cloned part to ensure it loads correctly

Common Notes for All Cloners

Backup first: Always backup source files before cloning

Monitor Inventor: Watch for popup dialogs during execution

Isolated copies: All cloners create fully isolated copies with no cross-references

Reference updates: Assembly and IDW references are automatically updated ( RUN IDW FIXER IF ANY REFERENCES ARE NOT UPDATED, DO NOT PANIC, IDW FIXER WILL FIX THIS AS LONG AS YOU HAVE NOT MOVED THE FOLDER OR DELETED THE MAPPING )

Logging: All operations are logged for troubleshooting

Mapping files: STEP\_1\_MAPPING.txt is generated for reference tracking



**When to Use Each Tool**

**Assembly Cloner:** When you need a complete copy of an entire project or sub-assembly

**Prefix Cloner:** When changing project prefixes (e.g., moving from Plant 1 to Plant 2) while keeping part numbering

**Part Cloner:** When you need to duplicate individual parts for modification



###### **Important Notes**

Always backup your files before running any automation scripts

Monitor Inventor closely during execution for popup dialogs

Run workflows in sequence - PART RENAMING first, then DETAILING

Check logs - All scripts generate log files for troubleshooting

Registry management is critical for maintaining sequential numbering

Mapping files are created during renaming and used by subsequent tools

