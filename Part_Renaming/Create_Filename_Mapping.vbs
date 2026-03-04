' Create_Filename_Mapping.vbs - Creates STEP_1_MAPPING.txt with filename-based mappings
' Based on cloner log

Dim fso, mappingFile
Dim clonedFolder, sourceFolder

clonedFolder = "C:\Users\Quintin\Documents\Spectiv\3. Working\22. DMS plant\20. SSCR03 - Deslime Screen Station\700 Head Chute\RENAME TEST - DELETE"
sourceFolder = "C:\Users\Quintin\Documents\Spectiv\3. Working\22. DMS plant\20. SSCR03 - Deslime Screen Station\700 Head Chute"

Set fso = CreateObject("Scripting.FileSystemObject")
mappingFile = fso.BuildPath(clonedFolder, "STEP_1_MAPPING_FILENAME.txt")

Set outFile = fso.CreateTextFile(mappingFile, True)
outFile.WriteLine("# Filename-based mapping file for cloned folder: " & clonedFolder)
outFile.WriteLine("# Format: OLD_FILENAME|NEW_FILENAME")
outFile.WriteLine("# For filename-only matching")
outFile.WriteLine("")

' Main assembly
outFile.WriteLine("Head Chute.iam|RENAME TEST - DELETE.iam")

' IAM replacements (sub-assemblies)
outFile.WriteLine("bottom.iam|Bottom.iam")
outFile.WriteLine("ncvr06-701-bolted connection1.iam|NCVR06-701-Bolted Connection1.iam")
outFile.WriteLine("ncvr06-701-bolted connection2.iam|NCVR06-701-Bolted Connection2.iam")
outFile.WriteLine("ncvr06-701-bolted connection3.iam|NCVR06-701-Bolted Connection3.iam")
outFile.WriteLine("launder.iam|Launder.iam")
outFile.WriteLine("lid-1.iam|Lid-1.iam")
outFile.WriteLine("lid-2.iam|Lid-2.iam")
outFile.WriteLine("middle.iam|Middle.iam")
outFile.WriteLine("support beam-1.iam|Support Beam-1.iam")
outFile.WriteLine("top.iam|Top.iam")

' IPT replacements - Top
outFile.WriteLine("lug hc-t.ipt|N1SCR04-001-PL8.ipt")
outFile.WriteLine("part1 hc-t.ipt|N1SCR04-001-PL1.ipt")
outFile.WriteLine("part10 angle hc-t.ipt|N1SCR04-001-A5.ipt")
outFile.WriteLine("part11 (part10 angle hc-t mir).ipt|N1SCR04-001-A6.ipt")
outFile.WriteLine("part12 angle hc-t.ipt|N1SCR04-001-A7.ipt")
outFile.WriteLine("part13 (part12 angle hc-t mir).ipt|N1SCR04-001-A8.ipt")
outFile.WriteLine("part14 angle hc-t.ipt|N1SCR04-001-A9.ipt")
outFile.WriteLine("part15 (part14 angle hc-t mir).ipt|N1SCR04-001-A10.ipt")
outFile.WriteLine("part16 angle hc-t.ipt|N1SCR04-001-A11.ipt")
outFile.WriteLine("part17 hc-t.ipt|N1SCR04-001-PL6.ipt")
outFile.WriteLine("part18 corner gusset hc-t.ipt|N1SCR04-001-PL7.ipt")
outFile.WriteLine("part2 hc-t.ipt|N1SCR04-001-PL2.ipt")
outFile.WriteLine("part3 hc-t.ipt|N1SCR04-001-PL3.ipt")
outFile.WriteLine("part4 hc-t.ipt|N1SCR04-001-PL4.ipt")
outFile.WriteLine("part5 hc-t.ipt|N1SCR04-001-PL5.ipt")
outFile.WriteLine("part6 angle hc-t.ipt|N1SCR04-001-A1.ipt")
outFile.WriteLine("part7 angle hc-t.ipt|N1SCR04-001-A2.ipt")
outFile.WriteLine("part8 angle hc-t.ipt|N1SCR04-001-A3.ipt")
outFile.WriteLine("part9 (part8 angle hc-t mir).ipt|N1SCR04-001-A4.ipt")

' Middle
outFile.WriteLine("lug hc-m.ipt|N1SCR04-001-PL19.ipt")
outFile.WriteLine("part1 hc-m.ipt|N1SCR04-001-PL9.ipt")
outFile.WriteLine("part10 angle hc-m.ipt|N1SCR04-001-A13.ipt")
outFile.WriteLine("part11 (part10 angle hc-m mir).ipt|N1SCR04-001-A14.ipt")
outFile.WriteLine("part12 angle hc-m.ipt|N1SCR04-001-A15.ipt")
outFile.WriteLine("part13 angle hc-m.ipt|N1SCR04-001-A16.ipt")
outFile.WriteLine("part14 angle hc-m.ipt|N1SCR04-001-A17.ipt")
outFile.WriteLine("part15 hc-m.ipt|N1SCR04-001-PL17.ipt")
outFile.WriteLine("part16 corner gusset hc-m.ipt|N1SCR04-001-PL18.ipt")
outFile.WriteLine("part17 stiffener hc-m.ipt|N1SCR04-001-FL1.ipt")
outFile.WriteLine("part18 stiffener hc-m.ipt|N1SCR04-001-FL2.ipt")
outFile.WriteLine("part19 stiffener hc-m.ipt|N1SCR04-001-FL3.ipt")
outFile.WriteLine("part2 hc-m.ipt|N1SCR04-001-PL10.ipt")
outFile.WriteLine("part20 stiffener hc-m.ipt|N1SCR04-001-FL4.ipt")
outFile.WriteLine("part3 hc-m.ipt|N1SCR04-001-PL11.ipt")
outFile.WriteLine("part4 hc-m.ipt|N1SCR04-001-PL12.ipt")
outFile.WriteLine("part5 hc-m.ipt|N1SCR04-001-PL13.ipt")
outFile.WriteLine("part6 hc-m.ipt|N1SCR04-001-PL14.ipt")
outFile.WriteLine("part7 hc-m.ipt|N1SCR04-001-PL15.ipt")
outFile.WriteLine("part8 hc-m.ipt|N1SCR04-001-PL16.ipt")
outFile.WriteLine("part9 angle hc-m.ipt|N1SCR04-001-A12.ipt")

' Launder
outFile.WriteLine("part1 hc-l.ipt|N1SCR04-001-PL20.ipt")
outFile.WriteLine("part10 stiffener hc-l.ipt|N1SCR04-001-FL7.ipt")
outFile.WriteLine("part11 stiffener hc-l.ipt|N1SCR04-001-FL8.ipt")
outFile.WriteLine("part12 stiffener hc-l.ipt|N1SCR04-001-FL9.ipt")
outFile.WriteLine("part13 stiffener hc-l.ipt|N1SCR04-001-FL10.ipt")
outFile.WriteLine("part14 stiffener hc-l.ipt|N1SCR04-001-FL11.ipt")
outFile.WriteLine("part2 hc-l.ipt|N1SCR04-001-PL21.ipt")
outFile.WriteLine("part3 hc-l.ipt|N1SCR04-001-PL22.ipt")
outFile.WriteLine("part4 angle hc-l.ipt|N1SCR04-001-A18.ipt")
outFile.WriteLine("part5 angle hc-l.ipt|N1SCR04-001-A19.ipt")
outFile.WriteLine("part6 pipe hc-l.ipt|N1SCR04-001-P1.ipt")
outFile.WriteLine("part7 flange hc-l.ipt|N1SCR04-001-FLG1.ipt")
outFile.WriteLine("part8 stiffener hc-l.ipt|N1SCR04-001-FL6.ipt")
outFile.WriteLine("part9 stiffener hc-l.ipt|N1SCR04-001-FL5.ipt")

' Bottom
outFile.WriteLine("lug hc-b.ipt|N1SCR04-001-PL29.ipt")
outFile.WriteLine("part1 hc-b.ipt|N1SCR04-001-PL23.ipt")
outFile.WriteLine("part10 stiffener hc-b.ipt|N1SCR04-001-FL12.ipt")
outFile.WriteLine("part11 stiffener hc-b.ipt|N1SCR04-001-FL13.ipt")
outFile.WriteLine("part12 stiffener hc-b.ipt|N1SCR04-001-FL14.ipt")
outFile.WriteLine("part13 stiffener hc-b.ipt|N1SCR04-001-FL15.ipt")
outFile.WriteLine("part14 stiffener hc-b.ipt|N1SCR04-001-FL16.ipt")
outFile.WriteLine("part2 hc-b.ipt|N1SCR04-001-PL24.ipt")
outFile.WriteLine("part3 hc-b.ipt|N1SCR04-001-PL25.ipt")
outFile.WriteLine("part4 hc-b.ipt|N1SCR04-001-PL26.ipt")
outFile.WriteLine("part5 hc-b.ipt|N1SCR04-001-PL27.ipt")
outFile.WriteLine("part6 hc-b.ipt|N1SCR04-001-PL28.ipt")
outFile.WriteLine("part7 angle hc-b.ipt|N1SCR04-001-A20.ipt")
outFile.WriteLine("part8 angle hc-b.ipt|N1SCR04-001-A21.ipt")
outFile.WriteLine("part9 angle hc-b.ipt|N1SCR04-001-A22.ipt")
outFile.WriteLine("part15.ipt|N1SCR04-001-PL30.ipt")
outFile.WriteLine("part16 corner gusset.ipt|N1SCR04-001-PL31.ipt")

' Lid-1
outFile.WriteLine("part1 hc-l1.ipt|N1SCR04-001-PL32.ipt")
outFile.WriteLine("part2 pipe hc-l1.ipt|N1SCR04-001-P2.ipt")
outFile.WriteLine("part3 flange hc-l1.ipt|N1SCR04-001-FLG2.ipt")

' Lid-2
outFile.WriteLine("part1 hc-l2.ipt|N1SCR04-001-PL33.ipt")
outFile.WriteLine("part2 stiffener hc-l2.ipt|N1SCR04-001-FL17.ipt")
outFile.WriteLine("part3 stiffener hc-l2.ipt|N1SCR04-001-FL18.ipt")
outFile.WriteLine("part4 handle hc-l2.ipt|N1SCR04-001-R1.ipt")

' Support Beam-1
outFile.WriteLine("part1 beam.ipt|N1SCR04-001-B1.ipt")
outFile.WriteLine("part2 end plate.ipt|N1SCR04-001-PL34.ipt")
outFile.WriteLine("part3 end plate.ipt|N1SCR04-001-PL35.ipt")

outFile.WriteLine("")
outFile.WriteLine("# Total mappings: 100+")
outFile.Close

WScript.Echo "Filename mapping file created: " & mappingFile