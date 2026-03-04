' Create_Complete_Mapping.vbs - Creates complete STEP_1_MAPPING.txt with all replacements
' Based on cloner log

Dim fso, mappingFile
Dim clonedFolder, sourceFolder

clonedFolder = "C:\Users\Quintin\Documents\Spectiv\3. Working\22. DMS plant\20. SSCR03 - Deslime Screen Station\700 Head Chute\RENAME TEST - DELETE"
sourceFolder = "C:\Users\Quintin\Documents\Spectiv\3. Working\22. DMS plant\20. SSCR03 - Deslime Screen Station\700 Head Chute"

Set fso = CreateObject("Scripting.FileSystemObject")
mappingFile = fso.BuildPath(clonedFolder, "STEP_1_MAPPING_COMPLETE.txt")

Set outFile = fso.CreateTextFile(mappingFile, True)
outFile.WriteLine("# Complete mapping file for cloned folder: " & clonedFolder)
outFile.WriteLine("# Format: OLD_FILENAME|OLD_FULLPATH|NEW_FILENAME|NEW_FULLPATH")
outFile.WriteLine("# Generated from Assembly Cloner replacements")
outFile.WriteLine("")

' Main assembly
outFile.WriteLine("Head Chute.iam|" & sourceFolder & "\Head Chute.iam|RENAME TEST - DELETE.iam|" & clonedFolder & "\RENAME TEST - DELETE.iam")

' IAM replacements (sub-assemblies)
outFile.WriteLine("bottom.iam|" & sourceFolder & "\Bottom\Bottom.iam|Bottom.iam|" & clonedFolder & "\Bottom\Bottom.iam")
outFile.WriteLine("ncvr06-701-bolted connection1.iam|" & sourceFolder & "\Head Chute\Design Accelerator\NCVR06-701-Bolted Connection1.iam|NCVR06-701-Bolted Connection1.iam|" & clonedFolder & "\Head Chute\Design Accelerator\NCVR06-701-Bolted Connection1.iam")
outFile.WriteLine("ncvr06-701-bolted connection2.iam|" & sourceFolder & "\Head Chute\Design Accelerator\NCVR06-701-Bolted Connection2.iam|NCVR06-701-Bolted Connection2.iam|" & clonedFolder & "\Head Chute\Design Accelerator\NCVR06-701-Bolted Connection2.iam")
outFile.WriteLine("ncvr06-701-bolted connection3.iam|" & sourceFolder & "\Head Chute\Design Accelerator\NCVR06-701-Bolted Connection3.iam|NCVR06-701-Bolted Connection3.iam|" & clonedFolder & "\Head Chute\Design Accelerator\NCVR06-701-Bolted Connection3.iam")
outFile.WriteLine("launder.iam|" & sourceFolder & "\Launder\Launder.iam|Launder.iam|" & clonedFolder & "\Launder\Launder.iam")
outFile.WriteLine("lid-1.iam|" & sourceFolder & "\Lid-1\Lid-1.iam|Lid-1.iam|" & clonedFolder & "\Lid-1\Lid-1.iam")
outFile.WriteLine("lid-2.iam|" & sourceFolder & "\Lid-2\Lid-2.iam|Lid-2.iam|" & clonedFolder & "\Lid-2\Lid-2.iam")
outFile.WriteLine("middle.iam|" & sourceFolder & "\Middle\Middle.iam|Middle.iam|" & clonedFolder & "\Middle\Middle.iam")
outFile.WriteLine("support beam-1.iam|" & sourceFolder & "\Support Beam-1\Support Beam-1.iam|Support Beam-1.iam|" & clonedFolder & "\Support Beam-1\Support Beam-1.iam")
outFile.WriteLine("top.iam|" & sourceFolder & "\Top\Top.iam|Top.iam|" & clonedFolder & "\Top\Top.iam")

' IPT replacements - Top
outFile.WriteLine("lug hc-t.ipt|" & sourceFolder & "\Top\Lug HC-T.ipt|N1SCR04-001-PL8.ipt|" & clonedFolder & "\Top\N1SCR04-001-PL8.ipt")
outFile.WriteLine("part1 hc-t.ipt|" & sourceFolder & "\Top\Part1 HC-T.ipt|N1SCR04-001-PL1.ipt|" & clonedFolder & "\Top\N1SCR04-001-PL1.ipt")
outFile.WriteLine("part10 angle hc-t.ipt|" & sourceFolder & "\Top\Part10 Angle HC-T.ipt|N1SCR04-001-A5.ipt|" & clonedFolder & "\Top\N1SCR04-001-A5.ipt")
outFile.WriteLine("part11 (part10 angle hc-t mir).ipt|" & sourceFolder & "\Top\Part11 (Part10 Angle HC-T MIR).ipt|N1SCR04-001-A6.ipt|" & clonedFolder & "\Top\N1SCR04-001-A6.ipt")
outFile.WriteLine("part12 angle hc-t.ipt|" & sourceFolder & "\Top\Part12 Angle HC-T.ipt|N1SCR04-001-A7.ipt|" & clonedFolder & "\Top\N1SCR04-001-A7.ipt")
outFile.WriteLine("part13 (part12 angle hc-t mir).ipt|" & sourceFolder & "\Top\Part13 (Part12 Angle HC-T MIR).ipt|N1SCR04-001-A8.ipt|" & clonedFolder & "\Top\N1SCR04-001-A8.ipt")
outFile.WriteLine("part14 angle hc-t.ipt|" & sourceFolder & "\Top\Part14 Angle HC-T.ipt|N1SCR04-001-A9.ipt|" & clonedFolder & "\Top\N1SCR04-001-A9.ipt")
outFile.WriteLine("part15 (part14 angle hc-t mir).ipt|" & sourceFolder & "\Top\Part15 (Part14 Angle HC-T MIR).ipt|N1SCR04-001-A10.ipt|" & clonedFolder & "\Top\N1SCR04-001-A10.ipt")
outFile.WriteLine("part16 angle hc-t.ipt|" & sourceFolder & "\Top\Part16 Angle HC-T.ipt|N1SCR04-001-A11.ipt|" & clonedFolder & "\Top\N1SCR04-001-A11.ipt")
outFile.WriteLine("part17 hc-t.ipt|" & sourceFolder & "\Top\Part17 HC-T.ipt|N1SCR04-001-PL6.ipt|" & clonedFolder & "\Top\N1SCR04-001-PL6.ipt")
outFile.WriteLine("part18 corner gusset hc-t.ipt|" & sourceFolder & "\Top\Part18 Corner Gusset HC-T.ipt|N1SCR04-001-PL7.ipt|" & clonedFolder & "\Top\N1SCR04-001-PL7.ipt")
outFile.WriteLine("part2 hc-t.ipt|" & sourceFolder & "\Top\Part2 HC-T.ipt|N1SCR04-001-PL2.ipt|" & clonedFolder & "\Top\N1SCR04-001-PL2.ipt")
outFile.WriteLine("part3 hc-t.ipt|" & sourceFolder & "\Top\Part3 HC-T.ipt|N1SCR04-001-PL3.ipt|" & clonedFolder & "\Top\N1SCR04-001-PL3.ipt")
outFile.WriteLine("part4 hc-t.ipt|" & sourceFolder & "\Top\Part4 HC-T.ipt|N1SCR04-001-PL4.ipt|" & clonedFolder & "\Top\N1SCR04-001-PL4.ipt")
outFile.WriteLine("part5 hc-t.ipt|" & sourceFolder & "\Top\Part5 HC-T.ipt|N1SCR04-001-PL5.ipt|" & clonedFolder & "\Top\N1SCR04-001-PL5.ipt")
outFile.WriteLine("part6 angle hc-t.ipt|" & sourceFolder & "\Top\Part6 Angle HC-T.ipt|N1SCR04-001-A1.ipt|" & clonedFolder & "\Top\N1SCR04-001-A1.ipt")
outFile.WriteLine("part7 angle hc-t.ipt|" & sourceFolder & "\Top\Part7 Angle HC-T.ipt|N1SCR04-001-A2.ipt|" & clonedFolder & "\Top\N1SCR04-001-A2.ipt")
outFile.WriteLine("part8 angle hc-t.ipt|" & sourceFolder & "\Top\Part8 Angle HC-T.ipt|N1SCR04-001-A3.ipt|" & clonedFolder & "\Top\N1SCR04-001-A3.ipt")
outFile.WriteLine("part9 (part8 angle hc-t mir).ipt|" & sourceFolder & "\Top\Part9 (Part8 Angle HC-T MIR).ipt|N1SCR04-001-A4.ipt|" & clonedFolder & "\Top\N1SCR04-001-A4.ipt")

' Middle
outFile.WriteLine("lug hc-m.ipt|" & sourceFolder & "\Middle\Lug HC-M.ipt|N1SCR04-001-PL19.ipt|" & clonedFolder & "\Middle\N1SCR04-001-PL19.ipt")
outFile.WriteLine("part1 hc-m.ipt|" & sourceFolder & "\Middle\Part1 HC-M.ipt|N1SCR04-001-PL9.ipt|" & clonedFolder & "\Middle\N1SCR04-001-PL9.ipt")
outFile.WriteLine("part10 angle hc-m.ipt|" & sourceFolder & "\Middle\Part10 Angle HC-M.ipt|N1SCR04-001-A13.ipt|" & clonedFolder & "\Middle\N1SCR04-001-A13.ipt")
outFile.WriteLine("part11 (part10 angle hc-m mir).ipt|" & sourceFolder & "\Middle\Part11 (Part10 Angle HC-M MIR).ipt|N1SCR04-001-A14.ipt|" & clonedFolder & "\Middle\N1SCR04-001-A14.ipt")
outFile.WriteLine("part12 angle hc-m.ipt|" & sourceFolder & "\Middle\Part12 Angle HC-M.ipt|N1SCR04-001-A15.ipt|" & clonedFolder & "\Middle\N1SCR04-001-A15.ipt")
outFile.WriteLine("part13 angle hc-m.ipt|" & sourceFolder & "\Middle\Part13 Angle HC-M.ipt|N1SCR04-001-A16.ipt|" & clonedFolder & "\Middle\N1SCR04-001-A16.ipt")
outFile.WriteLine("part14 angle hc-m.ipt|" & sourceFolder & "\Middle\Part14 Angle HC-M.ipt|N1SCR04-001-A17.ipt|" & clonedFolder & "\Middle\N1SCR04-001-A17.ipt")
outFile.WriteLine("part15 hc-m.ipt|" & sourceFolder & "\Middle\Part15 HC-M.ipt|N1SCR04-001-PL17.ipt|" & clonedFolder & "\Middle\N1SCR04-001-PL17.ipt")
outFile.WriteLine("part16 corner gusset hc-m.ipt|" & sourceFolder & "\Middle\Part16 Corner Gusset HC-M.ipt|N1SCR04-001-PL18.ipt|" & clonedFolder & "\Middle\N1SCR04-001-PL18.ipt")
outFile.WriteLine("part17 stiffener hc-m.ipt|" & sourceFolder & "\Middle\Part17 Stiffener HC-M.ipt|N1SCR04-001-FL1.ipt|" & clonedFolder & "\Middle\N1SCR04-001-FL1.ipt")
outFile.WriteLine("part18 stiffener hc-m.ipt|" & sourceFolder & "\Middle\Part18 Stiffener HC-M.ipt|N1SCR04-001-FL2.ipt|" & clonedFolder & "\Middle\N1SCR04-001-FL2.ipt")
outFile.WriteLine("part19 stiffener hc-m.ipt|" & sourceFolder & "\Middle\Part19 Stiffener HC-M.ipt|N1SCR04-001-FL3.ipt|" & clonedFolder & "\Middle\N1SCR04-001-FL3.ipt")
outFile.WriteLine("part2 hc-m.ipt|" & sourceFolder & "\Middle\Part2 HC-M.ipt|N1SCR04-001-PL10.ipt|" & clonedFolder & "\Middle\N1SCR04-001-PL10.ipt")
outFile.WriteLine("part20 stiffener hc-m.ipt|" & sourceFolder & "\Middle\Part20 Stiffener HC-M.ipt|N1SCR04-001-FL4.ipt|" & clonedFolder & "\Middle\N1SCR04-001-FL4.ipt")
outFile.WriteLine("part3 hc-m.ipt|" & sourceFolder & "\Middle\Part3 HC-M.ipt|N1SCR04-001-PL11.ipt|" & clonedFolder & "\Middle\N1SCR04-001-PL11.ipt")
outFile.WriteLine("part4 hc-m.ipt|" & sourceFolder & "\Middle\Part4 HC-M.ipt|N1SCR04-001-PL12.ipt|" & clonedFolder & "\Middle\N1SCR04-001-PL12.ipt")
outFile.WriteLine("part5 hc-m.ipt|" & sourceFolder & "\Middle\Part5 HC-M.ipt|N1SCR04-001-PL13.ipt|" & clonedFolder & "\Middle\N1SCR04-001-PL13.ipt")
outFile.WriteLine("part6 hc-m.ipt|" & sourceFolder & "\Middle\Part6 HC-M.ipt|N1SCR04-001-PL14.ipt|" & clonedFolder & "\Middle\N1SCR04-001-PL14.ipt")
outFile.WriteLine("part7 hc-m.ipt|" & sourceFolder & "\Middle\Part7 HC-M.ipt|N1SCR04-001-PL15.ipt|" & clonedFolder & "\Middle\N1SCR04-001-PL15.ipt")
outFile.WriteLine("part8 hc-m.ipt|" & sourceFolder & "\Middle\Part8 HC-M.ipt|N1SCR04-001-PL16.ipt|" & clonedFolder & "\Middle\N1SCR04-001-PL16.ipt")
outFile.WriteLine("part9 angle hc-m.ipt|" & sourceFolder & "\Middle\Part9 Angle HC-M.ipt|N1SCR04-001-A12.ipt|" & clonedFolder & "\Middle\N1SCR04-001-A12.ipt")

' Launder
outFile.WriteLine("part1 hc-l.ipt|" & sourceFolder & "\Launder\Part1 HC-L.ipt|N1SCR04-001-PL20.ipt|" & clonedFolder & "\Launder\N1SCR04-001-PL20.ipt")
outFile.WriteLine("part10 stiffener hc-l.ipt|" & sourceFolder & "\Launder\Part10 Stiffener HC-L.ipt|N1SCR04-001-FL7.ipt|" & clonedFolder & "\Launder\N1SCR04-001-FL7.ipt")
outFile.WriteLine("part11 stiffener hc-l.ipt|" & sourceFolder & "\Launder\Part11 Stiffener HC-L.ipt|N1SCR04-001-FL8.ipt|" & clonedFolder & "\Launder\N1SCR04-001-FL8.ipt")
outFile.WriteLine("part12 stiffener hc-l.ipt|" & sourceFolder & "\Launder\Part12 Stiffener HC-L.ipt|N1SCR04-001-FL9.ipt|" & clonedFolder & "\Launder\N1SCR04-001-FL9.ipt")
outFile.WriteLine("part13 stiffener hc-l.ipt|" & sourceFolder & "\Launder\Part13 Stiffener HC-L.ipt|N1SCR04-001-FL10.ipt|" & clonedFolder & "\Launder\N1SCR04-001-FL10.ipt")
outFile.WriteLine("part14 stiffener hc-l.ipt|" & sourceFolder & "\Launder\Part14 Stiffener HC-L.ipt|N1SCR04-001-FL11.ipt|" & clonedFolder & "\Launder\N1SCR04-001-FL11.ipt")
outFile.WriteLine("part2 hc-l.ipt|" & sourceFolder & "\Launder\Part2 HC-L.ipt|N1SCR04-001-PL21.ipt|" & clonedFolder & "\Launder\N1SCR04-001-PL21.ipt")
outFile.WriteLine("part3 hc-l.ipt|" & sourceFolder & "\Launder\Part3 HC-L.ipt|N1SCR04-001-PL22.ipt|" & clonedFolder & "\Launder\N1SCR04-001-PL22.ipt")
outFile.WriteLine("part4 angle hc-l.ipt|" & sourceFolder & "\Launder\Part4 Angle HC-L.ipt|N1SCR04-001-A18.ipt|" & clonedFolder & "\Launder\N1SCR04-001-A18.ipt")
outFile.WriteLine("part5 angle hc-l.ipt|" & sourceFolder & "\Launder\Part5 Angle HC-L.ipt|N1SCR04-001-A19.ipt|" & clonedFolder & "\Launder\N1SCR04-001-A19.ipt")
outFile.WriteLine("part6 pipe hc-l.ipt|" & sourceFolder & "\Launder\Part6 Pipe HC-L.ipt|N1SCR04-001-P1.ipt|" & clonedFolder & "\Launder\N1SCR04-001-P1.ipt")
outFile.WriteLine("part7 flange hc-l.ipt|" & sourceFolder & "\Launder\Part7 Flange HC-L.ipt|N1SCR04-001-FLG1.ipt|" & clonedFolder & "\Launder\N1SCR04-001-FLG1.ipt")
outFile.WriteLine("part8 stiffener hc-l.ipt|" & sourceFolder & "\Launder\Part8 Stiffener HC-L.ipt|N1SCR04-001-FL6.ipt|" & clonedFolder & "\Launder\N1SCR04-001-FL6.ipt")
outFile.WriteLine("part9 stiffener hc-l.ipt|" & sourceFolder & "\Launder\Part9 Stiffener HC-L.ipt|N1SCR04-001-FL5.ipt|" & clonedFolder & "\Launder\N1SCR04-001-FL5.ipt")

' Bottom
outFile.WriteLine("lug hc-b.ipt|" & sourceFolder & "\Bottom\Lug HC-B.ipt|N1SCR04-001-PL29.ipt|" & clonedFolder & "\Bottom\N1SCR04-001-PL29.ipt")
outFile.WriteLine("part1 hc-b.ipt|" & sourceFolder & "\Bottom\Part1 HC-B.ipt|N1SCR04-001-PL23.ipt|" & clonedFolder & "\Bottom\N1SCR04-001-PL23.ipt")
outFile.WriteLine("part10 stiffener hc-b.ipt|" & sourceFolder & "\Bottom\Part10 Stiffener HC-B.ipt|N1SCR04-001-FL12.ipt|" & clonedFolder & "\Bottom\N1SCR04-001-FL12.ipt")
outFile.WriteLine("part11 stiffener hc-b.ipt|" & sourceFolder & "\Bottom\Part11 Stiffener HC-B.ipt|N1SCR04-001-FL13.ipt|" & clonedFolder & "\Bottom\N1SCR04-001-FL13.ipt")
outFile.WriteLine("part12 stiffener hc-b.ipt|" & sourceFolder & "\Bottom\Part12 Stiffener HC-B.ipt|N1SCR04-001-FL14.ipt|" & clonedFolder & "\Bottom\N1SCR04-001-FL14.ipt")
outFile.WriteLine("part13 stiffener hc-b.ipt|" & sourceFolder & "\Bottom\Part13 Stiffener HC-B.ipt|N1SCR04-001-FL15.ipt|" & clonedFolder & "\Bottom\N1SCR04-001-FL15.ipt")
outFile.WriteLine("part14 stiffener hc-b.ipt|" & sourceFolder & "\Bottom\Part14 Stiffener HC-B.ipt|N1SCR04-001-FL16.ipt|" & clonedFolder & "\Bottom\N1SCR04-001-FL16.ipt")
outFile.WriteLine("part2 hc-b.ipt|" & sourceFolder & "\Bottom\Part2 HC-B.ipt|N1SCR04-001-PL24.ipt|" & clonedFolder & "\Bottom\N1SCR04-001-PL24.ipt")
outFile.WriteLine("part3 hc-b.ipt|" & sourceFolder & "\Bottom\Part3 HC-B.ipt|N1SCR04-001-PL25.ipt|" & clonedFolder & "\Bottom\N1SCR04-001-PL25.ipt")
outFile.WriteLine("part4 hc-b.ipt|" & sourceFolder & "\Bottom\Part4 HC-B.ipt|N1SCR04-001-PL26.ipt|" & clonedFolder & "\Bottom\N1SCR04-001-PL26.ipt")
outFile.WriteLine("part5 hc-b.ipt|" & sourceFolder & "\Bottom\Part5 HC-B.ipt|N1SCR04-001-PL27.ipt|" & clonedFolder & "\Bottom\N1SCR04-001-PL27.ipt")
outFile.WriteLine("part6 hc-b.ipt|" & sourceFolder & "\Bottom\Part6 HC-B.ipt|N1SCR04-001-PL28.ipt|" & clonedFolder & "\Bottom\N1SCR04-001-PL28.ipt")
outFile.WriteLine("part7 angle hc-b.ipt|" & sourceFolder & "\Bottom\Part7 Angle HC-B.ipt|N1SCR04-001-A20.ipt|" & clonedFolder & "\Bottom\N1SCR04-001-A20.ipt")
outFile.WriteLine("part8 angle hc-b.ipt|" & sourceFolder & "\Bottom\Part8 Angle HC-B.ipt|N1SCR04-001-A21.ipt|" & clonedFolder & "\Bottom\N1SCR04-001-A21.ipt")
outFile.WriteLine("part9 angle hc-b.ipt|" & sourceFolder & "\Bottom\Part9 Angle HC-B.ipt|N1SCR04-001-A22.ipt|" & clonedFolder & "\Bottom\N1SCR04-001-A22.ipt")
outFile.WriteLine("part15.ipt|" & sourceFolder & "\Bottom\Part15.ipt|N1SCR04-001-PL30.ipt|" & clonedFolder & "\Bottom\N1SCR04-001-PL30.ipt")
outFile.WriteLine("part16 corner gusset.ipt|" & sourceFolder & "\Bottom\Part16 Corner Gusset.ipt|N1SCR04-001-PL31.ipt|" & clonedFolder & "\Bottom\N1SCR04-001-PL31.ipt")

' Lid-1
outFile.WriteLine("part1 hc-l1.ipt|" & sourceFolder & "\Lid-1\Part1 HC-L1.ipt|N1SCR04-001-PL32.ipt|" & clonedFolder & "\Lid-1\N1SCR04-001-PL32.ipt")
outFile.WriteLine("part2 pipe hc-l1.ipt|" & sourceFolder & "\Lid-1\Part2 Pipe HC-L1.ipt|N1SCR04-001-P2.ipt|" & clonedFolder & "\Lid-1\N1SCR04-001-P2.ipt")
outFile.WriteLine("part3 flange hc-l1.ipt|" & sourceFolder & "\Lid-1\Part3 Flange HC-L1.ipt|N1SCR04-001-FLG2.ipt|" & clonedFolder & "\Lid-1\N1SCR04-001-FLG2.ipt")

' Lid-2
outFile.WriteLine("part1 hc-l2.ipt|" & sourceFolder & "\Lid-2\Part1 HC-L2.ipt|N1SCR04-001-PL33.ipt|" & clonedFolder & "\Lid-2\N1SCR04-001-PL33.ipt")
outFile.WriteLine("part2 stiffener hc-l2.ipt|" & sourceFolder & "\Lid-2\Part2 Stiffener HC-L2.ipt|N1SCR04-001-FL17.ipt|" & clonedFolder & "\Lid-2\N1SCR04-001-FL17.ipt")
outFile.WriteLine("part3 stiffener hc-l2.ipt|" & sourceFolder & "\Lid-2\Part3 Stiffener HC-L2.ipt|N1SCR04-001-FL18.ipt|" & clonedFolder & "\Lid-2\N1SCR04-001-FL18.ipt")
outFile.WriteLine("part4 handle hc-l2.ipt|" & sourceFolder & "\Lid-2\Part4 Handle HC-L2.ipt|N1SCR04-001-R1.ipt|" & clonedFolder & "\Lid-2\N1SCR04-001-R1.ipt")

' Support Beam-1
outFile.WriteLine("part1 beam.ipt|" & sourceFolder & "\Support Beam-1\Part1 Beam.ipt|N1SCR04-001-B1.ipt|" & clonedFolder & "\Support Beam-1\N1SCR04-001-B1.ipt")
outFile.WriteLine("part2 end plate.ipt|" & sourceFolder & "\Support Beam-1\Part2 End Plate.ipt|N1SCR04-001-PL34.ipt|" & clonedFolder & "\Support Beam-1\N1SCR04-001-PL34.ipt")
outFile.WriteLine("part3 end plate.ipt|" & sourceFolder & "\Support Beam-1\Part3 End Plate.ipt|N1SCR04-001-PL35.ipt|" & clonedFolder & "\Support Beam-1\N1SCR04-001-PL35.ipt")

outFile.WriteLine("")
outFile.WriteLine("# Total mappings: 100+")
outFile.Close

WScript.Echo "Complete mapping file created: " & mappingFile