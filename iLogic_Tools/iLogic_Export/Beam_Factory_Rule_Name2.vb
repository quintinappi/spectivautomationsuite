' ==============================================
' iLogic Rule Export
' ==============================================
' Source File: Beam_Factory
' Rule Name: Name2
' Exported: 12/15/2025 7:16:34 AM
' ==============================================

iProperties.Value("Custom", "Size_IBeam") = Size_IBeam
iProperties.Value("Custom", "Size_HBeam") = Size_HBeam

MultiValue.SetList("Beam_Type", "I-Beam", "H-Beam", "IPE")

MultiValue.SetList("Size_IPE", "IPE100", "IPE120", "IPE140", "IPE160", "IPE180", "IPE200")

MultiValue.SetList("Size_IBeam", "203 x 133 x 25", "203 x 133 x 30", "254 x 146 x 31", "254 x 146 x 37", "254 x 146 x 43",
"305 x 102 x 25", "305 x 102 x 28", "305 x 102 x 33", "305 x 165 x 40", "305 x 165 x 46", "305 x 165 x 54",
"356 x 171 x 45", "356 x 171 x 51", "356 x 171 x 57", "356 x 171 x 67", "406 x 140 x 39", "406 x 140 x 46",
"457 x 191 x 67", "457 x 191 x 74", "457 x 191 x 82", "457 x 191 x 89", "457 x 191 x 98",
"533 x 210 x 82", "533 x 210 x 92", "533 x 210 x 101", "533 x 210 x 109", "533 x 210 x 122")

MultiValue.SetList("Size_HBeam", "152 x 152 x 23", "152 x 152 x 30", "152 x 152 x 37",
"203 x 203 x 46", "203 x 203 x 52", "203 x 203 x 60", "203 x 203 x 71", "203 x 203 x 86",
"254 x 254 x 73", "254 x 254 x 89", "254 x 254 x 107", "254 x 254 x 132", "254 x 254 x 167",
"305 x 305 x 97", "305 x 305 x 118", "305 x 305 x 137", "305 x 305 x 158")

If 		Beam_Type = "I-Beam" Then
		I_Beam = True
		H_Beam = False
		IPE = False
		
iProperties.Value("Project", "Description")	 = "=<Size_IBeam> PFI"
		
Select Case Size_IBeam
	
	Case "203 x 133 x 25"
		kg_m = 25
		h_ = 203.2
		b_ = 133.2
		tw = 5.7
		tf = 7.8
		r1 = 7.6

	Case "203 x 133 x 30"
		kg_m = 30
		h_ = 206.8
		b_ = 133.9
		tw = 6.4
		tf = 9.6
		r1 = 7.6
		
	Case "254 x 146 x 31"
		kg_m = 31
		h_ = 251.4
		b_ = 146.1
		tw = 6
		tf = 8.6
		r1 = 7.6

	Case "254 x 146 x 37"
		kg_m = 37
		h_ = 256
		b_ = 146.4
		tw = 6.3
		tf = 10.9
		r1 = 7.6
		
	Case "254 x 146 x 43"
		kg_m = 43
		h_ = 259.6
		b_ = 147.3
		tw = 7.2
		tf = 12.7
		r1 = 7.6		

	Case "305 x 102 x 25"
		kg_m = 25
		h_ = 305.1
		b_ = 101.6
		tw = 5.8
		tf = 6.7
		r1 = 7.6
		
	Case "305 x 102 x 28"
		kg_m = 28
		h_ = 308.9
		b_ = 101.9
		tw = 6.1
		tf = 8.9
		r1 = 7.6
		
	Case "305 x 102 x 33"
		kg_m = 33
		h_ = 312.7
		b_ = 102.4
		tw = 6.6
		tf = 10.8
		r1 = 7.6
		
	Case "305 x 165 x 40"
		kg_m = 40
		h_ = 303.8
		b_ = 165.1
		tw = 6.1
		tf = 10.2
		r1 = 8.9
		
	Case "305 x 165 x 46"
		kg_m = 46
		h_ = 307.1
		b_ = 165.7
		tw = 6.7
		tf = 11.8
		r1 = 8.9
		
	Case "305 x 165 x 54"
		kg_m = 54
		h_ = 310.9
		b_ = 166.8
		tw = 7.7
		tf = 13.7
		r1 = 8.9
		
	Case "356 x 171 x 45"
		kg_m = 45
		h_ = 352
		b_ = 171
		tw = 6.9
		tf = 9.7
		r1 =  10.2
		
	Case "356 x 171 x 51"
		kg_m = 51
		h_ = 355.6
		b_ = 171.5
		tw = 7.3
		tf = 11.5
		r1 = 10.2
		
	Case "356 x 171 x 57"
		kg_m = 57
		h_ = 358.6
		b_ = 172.1
		tw = 8
		tf = 13
		r1 = 10.2
		
	Case "356 x 171 x 67"
		kg_m = 67
		h_ = 364
		b_ = 173.2
		tw = 9.1
		tf = 15.7
		r1 = 10.2
		
	Case "406 x 140 x 39"
		kg_m = 39
		h_ = 397.3
		b_ = 141.8
		tw = 6.3
		tf = 8.6
		r1 = 10.2
		
	Case "406 x 140 x 46"
		kg_m = 46
		h_ = 402.3
		b_ = 142.4
		tw = 6.9
		tf = 11.2
		r1 = 10.2
		
	Case "406 x 178 x 54"
		kg_m = 54
		h_ = 402.6
		b_ = 177.6
		tw = 7.6
		tf = 10.9
		r1 = 10.2
		
	Case "406 x 178 x 60"
		kg_m = 60
		h_ = 406.4
		b_ = 177.8
		tw = 7.8
		tf = 12.8
		r1 = 10.2
		
	Case "406 x 178 x 67"
		kg_m = 67
		h_ = 409.4
		b_ = 178.8
		tw = 8.8
		tf = 14.3
		r1 = 10.2
		
	Case "406 x 178 x 74"
		kg_m = 74
		h_ = 412.8
		b_ = 179.7
		tw = 9.7
		tf = 16
		r1 = 10.2
		
	Case "457 x 191 x 67"
		kg_m = 67
		h_ = 453.6
		b_ = 189.9
		tw = 8.5
		tf = 12.7
		r1 = 10.2
		
	Case "457 x 191 x 74"
		kg_m = 74
		h_ = 457.2
		b_ = 190.5
		tw = 9.1
		tf = 14.5
		r1 = 10.2
		
	Case "457 x 191 x 82"
		kg_m = 82
		h_ = 460.2
		b_ = 191.3
		tw = 9.9
		tf = 16
		r1 = 10.2
		
	Case "457 x 191 x 89"
		kg_m = 89
		h_ = 463.6
		b_ = 192
		tw = 10.6
		tf = 17.7
		r1 = 10.2
		
	Case "457 x 191 x 98"
		kg_m = 98
		h_ = 467.6
		b_ = 192.8
		tw = 11.4
		tf = 19.6
		r1 = 10.2
		
	Case "533 x 210 x 82"
		kg_m = 82
		h_ = 528.3
		b_ = 208.7
		tw = 9.6
		tf = 13.2
		r1 = 12.7
		
	Case "533 x 210 x 92"
		kg_m = 92
		h_ = 533.1
		b_ = 209.3
		tw = 10.2
		tf = 15.6
		r1 = 12.7
		
	Case "533 x 210 x 101"
		kg_m = 101
		h_ = 536.7
		b_ = 210.1
		tw = 10.9
		tf = 17.4
		r1 = 12.7
		
	Case "533 x 210 x 109"
		kg_m = 109
		h_ = 539.5
		b_ = 210.7
		tw = 11.6
		tf = 18.8
		r1 = 12.7
		
	Case "533 x 210 x 122"
		kg_m = 122
		h_ = 544.6
		b_ = 211.9
		tw = 12.8
		tf = 21.3
		r1 = 12.7
		
End Select			
		
Else If Beam_Type = "H-Beam" Then
		I_Beam = False
		H_Beam = True
		IPE = False
		
iProperties.Value("Project", "Description")	 = "=<Size_HBeam> PFH"				
		
Select Case Size_HBeam
	
	Case "152 x 152 x 23"
		kg_m = 23
		h_ = 152.4
		b_ = 152.4
		tw = 6.1
		tf = 6.8
		r1 = 7.6
		
	Case "152 x 152 x 30"
		kg_m = 30
		h_ = 157.5
		b_ = 152.9
		tw = 6.6
		tf = 9.4
		r1 = 7.6
		
	Case "152 x 152 x 37"
		kg_m = 37
		h_ = 161.8
		b_ = 154.4
		tw = 8.1
		tf = 11.5
		r1 = 7.6
		
	Case "203 x 203 x 46"
		kg_m = 46
		h_ = 203.2
		b_ = 203.2
		tw = 7.3
		tf = 11
		r1 = 10.2
		
	Case "203 x 203 x 52"
		kg_m = 52
		h_ = 206.2
		b_ = 203.9
		tw = 8
		tf = 12.5
		r1 = 10.2
		
	Case "203 x 203 x 60"
		kg_m = 60
		h_ = 209.6
		b_ = 205.2
		tw = 9.3
		tf = 14.2
		r1 = 10.2
		
	Case "203 x 203 x 71"
		kg_m = 71
		h_ = 215.9
		b_ = 206.2
		tw = 10.3
		tf = 17.3
		r1 = 10.2
		
	Case "203 x 203 x 86"
		kg_m = 86
		h_ = 222.3
		b_ = 208.8
		tw = 13
		tf = 20.5
		r1 = 10.2
		
	Case "254 x 254 x 73"
		kg_m = 73
		h_ = 254.2
		b_ = 254
		tw = 8.6
		tf = 14.2
		r1 = 12.7
		
	Case "254 x 254 x 89"
		kg_m = 89
		h_ = 260.4
		b_ = 255.9
		tw = 10.5
		tf = 17.3
		r1 = 12.7
		
	Case "254 x 254 x 107"
		kg_m = 107
		h_ = 266.7
		b_ = 258.3
		tw = 13
		tf = 20.5
		r1 = 12.7
		
	Case "254 x 254 x 132"
		kg_m = 132
		h_ = 276.4
		b_ = 261
		tw = 15.6
		tf = 25.1
		r1 = 12.7
		
	Case "254 x 254 x 167"
		kg_m = 167
		h_ = 289.1
		b_ = 264.5
		tw = 19.2
		tf = 31.7
		r1 = 12.7
		
	Case "305 x 305 x 97"
		kg_m = 97
		h_ = 307.8
		b_ = 304.8
		tw = 9.9
		tf = 15.4
		r1 = 15.2
		
	Case "305 x 305 x 118"
		kg_m = 118
		h_ = 314.5
		b_ = 306.8
		tw = 11.9
		tf = 18.7
		r1 = 15.2
		
	Case "305 x 305 x 137"
		kg_m = 137
		h_ = 320.5
		b_ = 308.7
		tw = 13.8
		tf = 21.7
		r1 = 15.2
		
	Case "305 x 305 x 158"
		kg_m = 158
		h_ = 327.2
		b_ = 310.6
		tw = 15.7
		tf = 25
		r1 = 15.2	
		
End Select		
		
Else If	Beam_Type = "IPE" Then
		I_Beam = False
		H_Beam = False
		IPE = True
		
iProperties.Value("Project", "Description")	 = Size_IPE	
		
Select Case Size_IPE
	
	Case "IPE100"
		h_ = 100
		b_ = 55
		tw = 4.1
		tf = 5.7
		r1 = 7
		
	Case "IPE120"
		h_ = 120
		b_ = 64
		tw = 4.4
		tf = 6.3
		r1 = 7
		
	Case "IPE140"
		h_ = 140
		b_ = 73
		tw = 4.7
		tf = 6.9
		r1 = 7
		
	Case "IPE160"
		h_ = 160
		b_ = 82
		tw = 5
		tf = 7.4
		r1 =9
		
	Case "IPE180"
		h_ = 180
		b_ = 91
		tw = 5.3
		tf = 8
		r1 = 9

	Case "IPE200"
		h_ = 200
		b_ = 100
		tw = 5.6
		tf = 8.5
		r1 = 12
	
End Select

End If	

'MultiValue.UpdateAfterChange = True

iLogicVb.UpdateWhenDone = True

