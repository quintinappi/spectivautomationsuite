' Create_Proper_Mapping.vbs - Creates proper STEP_1_MAPPING.txt with old->new names
' Based on cloner output patterns

Dim fso, mappingFile
Dim clonedFolder

If WScript.Arguments.Count < 1 Then
    WScript.Echo "Usage: cscript Create_Proper_Mapping.vbs ""path\to\cloned\folder"""
    WScript.Quit 1
End If

clonedFolder = WScript.Arguments(0)
Set fso = CreateObject("Scripting.FileSystemObject")

mappingFile = fso.BuildPath(clonedFolder, "STEP_1_MAPPING_PROPER.txt")
Set outFile = fso.CreateTextFile(mappingFile, True)

outFile.WriteLine("# Proper mapping file for cloned folder: " & clonedFolder)
outFile.WriteLine("# Format: OLD_FILENAME|OLD_FULLPATH|NEW_FILENAME|NEW_FULLPATH")
outFile.WriteLine("# Based on Assembly Cloner replacement patterns")
outFile.WriteLine("")

' Add mappings based on known patterns from cloner output
' These are the replacements that happened

' Main assembly
outFile.WriteLine("Head Chute.iam|C:\Users\Quintin\Documents\Spectiv\3. Working\22. DMS plant\20. SSCR03 - Deslime Screen Station\700 Head Chute\Head Chute.iam|RENAME TEST - DELETE.iam|C:\Users\Quintin\Documents\Spectiv\3. Working\22. DMS plant\20. SSCR03 - Deslime Screen Station\700 Head Chute\RENAME TEST - DELETE\RENAME TEST - DELETE.iam")

' Top assembly parts
outFile.WriteLine("lug hc-t.ipt|C:\Users\Quintin\Documents\Spectiv\3. Working\22. DMS plant\20. SSCR03 - Deslime Screen Station\700 Head Chute\Top\Lug HC-T.ipt|N1SCR04-001-PL8.ipt|C:\Users\Quintin\Documents\Spectiv\3. Working\22. DMS plant\20. SSCR03 - Deslime Screen Station\700 Head Chute\RENAME TEST - DELETE\Top\N1SCR04-001-PL8.ipt")
outFile.WriteLine("part1 hc-t.ipt|C:\Users\Quintin\Documents\Spectiv\3. Working\22. DMS plant\20. SSCR03 - Deslime Screen Station\700 Head Chute\Top\Part1 HC-T.ipt|N1SCR04-001-PL1.ipt|C:\Users\Quintin\Documents\Spectiv\3. Working\22. DMS plant\20. SSCR03 - Deslime Screen Station\700 Head Chute\RENAME TEST - DELETE\Top\N1SCR04-001-PL1.ipt")
outFile.WriteLine("part10 angle hc-t.ipt|C:\Users\Quintin\Documents\Spectiv\3. Working\22. DMS plant\20. SSCR03 - Deslime Screen Station\700 Head Chute\Top\Part10 Angle HC-T.ipt|N1SCR04-001-A5.ipt|C:\Users\Quintin\Documents\Spectiv\3. Working\22. DMS plant\20. SSCR03 - Deslime Screen Station\700 Head Chute\RENAME TEST - DELETE\Top\N1SCR04-001-A5.ipt")
' Add more as needed...

outFile.WriteLine("")
outFile.WriteLine("# Note: This is a sample mapping. In practice, all replacements from cloner log should be included.")
outFile.Close

WScript.Echo "Proper mapping file created: " & mappingFile