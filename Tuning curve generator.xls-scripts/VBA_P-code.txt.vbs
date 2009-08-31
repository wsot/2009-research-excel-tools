' Processing file: Tuning curve generator.xls
' ===============================================================================
' Module streams:
' _VBA_PROJECT_CUR/VBA/ThisWorkbook - 1210 bytes
' Line #0:
' 	Option  (Explicit)
' Line #1:
' _VBA_PROJECT_CUR/VBA/Sheet1 - 1150 bytes
' _VBA_PROJECT_CUR/VBA/ImportFrom - 5492 bytes
' Line #0:
' Line #1:
' 	FuncDefn (Private Sub Cancel_Click())
' Line #2:
' 	LitVarSpecial (False)
' 	St doImport 
' Line #3:
' 	Ld id_FFFF 
' 	ArgsCall Unload 0x0001 
' 	QuoteRem 0x0015 0x0015 "Unloads the UserForm."
' Line #4:
' 	EndSub 
' Line #5:
' Line #6:
' 	FuncDefn (Private Sub ImportButton_Click())
' Line #7:
' 	Ld BlockSelect1 
' 	MemLd ActiveBlock 
' 	LitStr 0x0000 ""
' 	Ne 
' 	IfBlock 
' Line #8:
' 	Ld BlockSelect1 
' 	MemLd UseServer 
' 	St theServer 
' Line #9:
' 	Ld BlockSelect1 
' 	MemLd UseTank 
' 	St theTank 
' Line #10:
' 	Ld BlockSelect1 
' 	MemLd ActiveBlock 
' 	St theBlock 
' Line #11:
' 	LitVarSpecial (True)
' 	St doImport 
' Line #12:
' 	Ld id_FFFF 
' 	ArgsCall Unload 0x0001 
' Line #13:
' 	ElseBlock 
' Line #14:
' 	LitStr 0x001F "Please select a block to import"
' 	Paren 
' 	ArgsCall MsgBox 0x0001 
' Line #15:
' 	EndIfBlock 
' Line #16:
' 	EndSub 
' Line #17:
' Line #18:
' 	FuncDefn (Private Sub TankSelect1_TankChanged(ActTank As String, ActServer As String))
' Line #19:
' 	Ld ActServer 
' 	Ld BlockSelect1 
' 	MemSt UseServer 
' Line #20:
' 	Ld ActTank 
' 	Ld BlockSelect1 
' 	MemSt UseTank 
' Line #21:
' 	Ld BlockSelect1 
' 	ArgsMemCall (Call) Refresh 0x0000 
' Line #22:
' 	EndSub 
' Line #23:
' Line #24:
' 	FuncDefn (Private Sub UserForm_Activate())
' Line #25:
' 	Ld theServer 
' 	LitStr 0x0000 ""
' 	Ne 
' 	IfBlock 
' Line #26:
' 	Ld theServer 
' 	Ld TankSelect1 
' 	MemSt UseServer 
' Line #27:
' 	Ld theServer 
' 	Ld BlockSelect1 
' 	MemSt UseServer 
' Line #28:
' 	Ld theTank 
' 	LitStr 0x0000 ""
' 	Ne 
' 	IfBlock 
' Line #29:
' 	Ld theTank 
' 	Ld TankSelect1 
' 	MemSt ActiveTank 
' Line #30:
' 	Ld theTank 
' 	Ld BlockSelect1 
' 	MemSt UseTank 
' Line #31:
' 	Ld theBlock 
' 	LitStr 0x0000 ""
' 	Ne 
' 	IfBlock 
' Line #32:
' 	Ld theBlock 
' 	Ld BlockSelect1 
' 	MemSt ActiveBlock 
' Line #33:
' 	EndIfBlock 
' Line #34:
' 	Ld BlockSelect1 
' 	ArgsMemCall Refresh 0x0000 
' Line #35:
' 	EndIfBlock 
' Line #36:
' 	Ld TankSelect1 
' 	ArgsMemCall Refresh 0x0000 
' Line #37:
' 	EndIfBlock 
' Line #38:
' 	EndSub 
' _VBA_PROJECT_CUR/VBA/Module1 - 17761 bytes
' Line #0:
' 	Option  (Explicit)
' Line #1:
' 	Dim (Global) 
' 	VarDefn doImport
' Line #2:
' 	Dim (Global) 
' 	VarDefn theServer
' 	VarDefn theTank
' 	VarDefn theBlock
' Line #3:
' Line #4:
' 	FuncDefn (Sub buildTuningCurves())
' Line #5:
' 	Ld ImportFrom 
' 	ArgsMemCall Show 0x0000 
' Line #6:
' Line #7:
' 	Ld doImport 
' 	IfBlock 
' Line #8:
' 	ArgsCall (Call) processImport 0x0000 
' Line #9:
' 	EndIfBlock 
' Line #10:
' 	EndSub 
' Line #11:
' Line #12:
' 	FuncDefn (Sub processImport())
' Line #13:
' 	Dim 
' 	VarDefn lBinWidth (As Double)
' Line #14:
' 	LitStr 0x0002 "B1"
' 	LitStr 0x0008 "Settings"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	St lBinWidth 
' Line #15:
' 	Dim 
' 	VarDefn theWorksheets (As Variant)
' Line #16:
' 	QuoteRem 0x0004 0x007E "Dim chanHistTmp(32) As Long 'used as a temporary store to build a histogram across multiple 'Swep's before outputting the data"
' Line #17:
' 	Dim 
' 	OptionBase 
' 	LitDI2 0x001F 
' 	VarDefn _B_var_arrHistTmp (As Long)
' Line #18:
' Line #19:
' 	Dim 
' 	VarDefn i (As Long)
' Line #20:
' 	Dim 
' 	VarDefn j (As Long)
' Line #21:
' 	Dim 
' 	VarDefn k (As Long)
' Line #22:
' 	Dim 
' 	VarDefn l (As Long)
' Line #23:
' Line #24:
' 	ArgsLd buildWorksheetArray 0x0000 
' 	St theWorksheets 
' Line #25:
' Line #26:
' 	Dim 
' 	VarDefn objttx
' Line #27:
' 	SetStmt 
' 	LitStr 0x0007 "TTank.X"
' 	ArgsLd CreateObject 0x0001 
' 	Set objttx 
' Line #28:
' Line #29:
' 	Ld theServer 
' 	LitStr 0x0002 "Me"
' 	Ld objttx 
' 	ArgsMemLd ConnectServer 0x0002 
' 	LitDI2 0x0001 
' 	Coerce (Lng) 
' 	Ne 
' 	IfBlock 
' Line #30:
' 	LitStr 0x0015 "Connecting to server "
' 	Ld theServer 
' 	Concat 
' 	LitStr 0x0008 " failed."
' 	Concat 
' 	Paren 
' 	ArgsCall MsgBox 0x0001 
' Line #31:
' 	ExitSub 
' Line #32:
' 	EndIfBlock 
' Line #33:
' Line #34:
' 	Ld theTank 
' 	LitStr 0x0001 "R"
' 	Ld objttx 
' 	ArgsMemLd OpenTank 0x0002 
' 	LitDI2 0x0001 
' 	Coerce (Lng) 
' 	Ne 
' 	IfBlock 
' Line #35:
' 	LitStr 0x0013 "Connecting to tank "
' 	Ld theTank 
' 	Concat 
' 	LitStr 0x000B " on server "
' 	Concat 
' 	Ld theServer 
' 	Concat 
' 	LitStr 0x0009 " failed ."
' 	Concat 
' 	Paren 
' 	ArgsCall MsgBox 0x0001 
' Line #36:
' 	Ld objttx 
' 	ArgsMemCall (Call) ReleaseServer 0x0000 
' Line #37:
' 	ExitSub 
' Line #38:
' 	EndIfBlock 
' Line #39:
' Line #40:
' 	Ld theBlock 
' 	Ld objttx 
' 	ArgsMemLd SelectBlock 0x0001 
' 	LitDI2 0x0001 
' 	Coerce (Lng) 
' 	Ne 
' 	IfBlock 
' Line #41:
' 	LitStr 0x0014 "Connecting to block "
' 	Ld theBlock 
' 	Concat 
' 	LitStr 0x0009 " in tank "
' 	Concat 
' 	Ld theTank 
' 	Concat 
' 	LitStr 0x000B " on server "
' 	Concat 
' 	Ld theServer 
' 	Concat 
' 	LitStr 0x0008 " failed."
' 	Concat 
' 	Paren 
' 	ArgsCall MsgBox 0x0001 
' Line #42:
' 	Ld objttx 
' 	ArgsMemCall (Call) CloseTank 0x0000 
' Line #43:
' 	Ld objttx 
' 	ArgsMemCall (Call) ReleaseServer 0x0000 
' Line #44:
' 	ExitSub 
' Line #45:
' 	EndIfBlock 
' Line #46:
' Line #47:
' 	Ld objttx 
' 	ArgsMemCall (Call) CreateEpocIndexing 0x0000 
' Line #48:
' Line #49:
' 	Dim 
' 	VarDefn freqList (As Dictionary)
' Line #50:
' 	Dim 
' 	VarDefn ampList (As Dictionary)
' Line #51:
' Line #52:
' 	SetStmt 
' 	New id_FFFF
' 	Set freqList 
' Line #53:
' 	SetStmt 
' 	New id_FFFF
' 	Set ampList 
' Line #54:
' Line #55:
' 	Dim 
' 	VarDefn dblStartTime (As Double)
' Line #56:
' 	Dim 
' 	VarDefn dblEndTime (As Double)
' Line #57:
' 	Dim 
' 	VarDefn varReturn (As Variant)
' Line #58:
' 	Dim 
' 	VarDefn varAmp (As Variant)
' Line #59:
' Line #60:
' Line #61:
' 	Do 
' Line #62:
' 	LitDI2 0x01F4 
' 	LitStr 0x0004 "Frq1"
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	Ld dblStartTime 
' 	LitR8 0x0000 0x0000 0x0000 0x0000 
' 	LitStr 0x0003 "ALL"
' 	Ld objttx 
' 	ArgsMemLd ReadEventsV 0x0007 
' 	St i 
' Line #63:
' 	Ld i 
' 	LitDI2 0x0000 
' 	Eq 
' 	IfBlock 
' Line #64:
' 	ExitDo 
' Line #65:
' 	EndIfBlock 
' Line #66:
' Line #67:
' 	LitDI2 0x0000 
' 	Ld i 
' 	LitDI2 0x0000 
' 	Ld objttx 
' 	ArgsMemLd ParseEvInfoV 0x0003 
' 	St varReturn 
' Line #68:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld i 
' 	LitDI2 0x0001 
' 	Sub 
' 	Paren 
' 	For 
' Line #69:
' 	LitDI2 0x0006 
' 	Ld j 
' 	ArgsLd varReturn 0x0002 
' 	Ld freqList 
' 	ArgsMemLd Exists 0x0001 
' 	Not 
' 	IfBlock 
' Line #70:
' 	LitDI2 0x0006 
' 	Ld j 
' 	ArgsLd varReturn 0x0002 
' 	LitStr 0x0000 ""
' 	Ld freqList 
' 	ArgsMemCall (Call) Add 0x0002 
' Line #71:
' 	EndIfBlock 
' Line #72:
' 	LitDI2 0x0005 
' 	Ld j 
' 	ArgsLd varReturn 0x0002 
' 	LitDI2 0x0001 
' 	LitDI4 0x86A0 0x0001 
' 	Div 
' 	Paren 
' 	Add 
' 	St dblStartTime 
' Line #73:
' 	LitStr 0x0004 "Lev1"
' 	LitDI2 0x0005 
' 	Ld j 
' 	ArgsLd varReturn 0x0002 
' 	LitDI2 0x0000 
' 	Ld objttx 
' 	ArgsMemLd QryEpocAtV 0x0003 
' 	St varAmp 
' Line #74:
' 	Ld varAmp 
' 	Ld ampList 
' 	ArgsMemLd Exists 0x0001 
' 	Not 
' 	IfBlock 
' Line #75:
' 	Ld varAmp 
' 	LitStr 0x0000 ""
' 	Ld ampList 
' 	ArgsMemCall (Call) Add 0x0002 
' Line #76:
' 	EndIfBlock 
' Line #77:
' Line #78:
' 	StartForVariable 
' 	Next 
' Line #79:
' Line #80:
' 	Ld i 
' 	LitDI2 0x01F4 
' 	Lt 
' 	IfBlock 
' Line #81:
' 	ExitDo 
' Line #82:
' 	EndIfBlock 
' Line #83:
' 	Loop 
' Line #84:
' Line #85:
' 	LitDI2 0x0000 
' 	St i 
' Line #86:
' 	LitDI2 0x0000 
' 	St j 
' Line #87:
' Line #88:
' 	QuoteRem 0x0000 0x0016 "    Dim freqAmpArray()"
' Line #89:
' 	QuoteRem 0x0000 0x003F "    Dim freqAmpArray(freqList.Count - 1, ampList.Count - 1, 32)"
' Line #90:
' 	Dim 
' 	VarDefn iFreqIndex (As Integer)
' Line #91:
' 	Dim 
' 	VarDefn iAmpIndex (As Integer)
' Line #92:
' Line #93:
' 	Dim 
' 	VarDefn vFreqKeys (As Variant)
' Line #94:
' 	Dim 
' 	VarDefn vAmpKeys (As Variant)
' Line #95:
' Line #96:
' Line #97:
' 	Ld freqList 
' 	MemLd Keys 
' 	St vFreqKeys 
' Line #98:
' 	Ld ampList 
' 	MemLd Keys 
' 	St vAmpKeys 
' Line #99:
' Line #100:
' 	Dim 
' 	VarDefn varChanData (As Variant)
' Line #101:
' 	Dim 
' 	VarDefn IsEmpty (As Double)
' Line #102:
' Line #103:
' 	StartForVariable 
' 	Ld iFreqIndex 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld vFreqKeys 
' 	FnUBound 0x0000 
' 	For 
' Line #104:
' 	StartForVariable 
' 	Ld iAmpIndex 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld vAmpKeys 
' 	FnUBound 0x0000 
' 	For 
' Line #105:
' 	LitStr 0x0007 "Frq1 = "
' 	Ld iFreqIndex 
' 	ArgsLd vFreqKeys 0x0001 
' 	Coerce (Str) 
' 	Concat 
' 	LitStr 0x000C " and Lev1 = "
' 	Concat 
' 	Ld iAmpIndex 
' 	ArgsLd vAmpKeys 0x0001 
' 	Coerce (Str) 
' 	Concat 
' 	Ld objttx 
' 	ArgsMemCall (Call) SetFilterWithDescEx 0x0001 
' Line #106:
' 	LitStr 0x0004 "Swep"
' 	LitDI2 0x0000 
' 	Ld objttx 
' 	ArgsMemLd GetEpocsExV 0x0002 
' 	St varReturn 
' Line #107:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld varReturn 
' 	LitDI2 0x0002 
' 	FnUBound 0x0001 
' 	For 
' Line #108:
' 	LitDI2 0x0002 
' 	Ld i 
' 	ArgsLd varReturn 0x0002 
' 	St dblStartTime 
' Line #109:
' 	Ld dblStartTime 
' 	Ld lBinWidth 
' 	Add 
' 	St dblEndTime 
' Line #110:
' 	QuoteRem 0x0010 0x001C "dblEndTime = varReturn(3, i)"
' Line #111:
' 	Ld dblStartTime 
' 	St IsEmpty 
' Line #112:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0001 
' 	LitDI2 0x0020 
' 	For 
' Line #113:
' 	Do 
' Line #114:
' 	LitDI2 0x01F4 
' 	LitStr 0x0004 "CSPK"
' 	Ld j 
' 	LitDI2 0x0000 
' 	Ld dblStartTime 
' 	Ld dblEndTime 
' 	LitStr 0x0009 "JUSTTIMES"
' 	Ld objttx 
' 	ArgsMemLd ReadEventsV 0x0007 
' 	St k 
' Line #115:
' 	Ld k 
' 	LitDI2 0x0000 
' 	Eq 
' 	IfBlock 
' Line #116:
' 	ExitDo 
' Line #117:
' 	EndIfBlock 
' Line #118:
' Line #119:
' 	Ld j 
' 	LitDI2 0x0001 
' 	Sub 
' 	ArgsLd _B_var_arrHistTmp 0x0001 
' 	Coerce (Lng) 
' 	Ld k 
' 	Coerce (Lng) 
' 	Add 
' 	Ld j 
' 	LitDI2 0x0001 
' 	Sub 
' 	ArgsSt _B_var_arrHistTmp 0x0001 
' Line #120:
' Line #121:
' 	QuoteRem 0x0000 0x0042 "                        varChanData = objttx.ParseEvInfoV(0, k, 6)"
' Line #122:
' 	QuoteRem 0x0000 0x002C "                        For l = 0 To (k - 1)"
' Line #123:
' 	QuoteRem 0x0000 0x0059 "                            Worksheets.Item("Settings").Cells(iAmpIndex + 3, 1).Value = j"
' Line #124:
' 	QuoteRem 0x0000 0x0066 "                            Worksheets.Item("Settings").Cells(iAmpIndex + 3, 2).Value = varChanData(0)"
' Line #125:
' 	QuoteRem 0x0000 0x001C "                        Next"
' Line #126:
' Line #127:
' 	Ld k 
' 	LitDI2 0x01F4 
' 	Lt 
' 	IfBlock 
' Line #128:
' 	ExitDo 
' Line #129:
' 	ElseBlock 
' Line #130:
' 	Ld k 
' 	LitDI2 0x0001 
' 	Sub 
' 	LitDI2 0x0001 
' 	LitDI2 0x0006 
' 	Ld objttx 
' 	ArgsMemLd ParseEvInfoV 0x0003 
' 	St varChanData 
' Line #131:
' 	LitDI2 0x0000 
' 	ArgsLd varChanData 0x0001 
' 	LitDI2 0x0001 
' 	LitDI4 0x86A0 0x0001 
' 	Div 
' 	Paren 
' 	Add 
' 	St dblStartTime 
' Line #132:
' 	EndIfBlock 
' Line #133:
' 	Loop 
' Line #134:
' 	Ld IsEmpty 
' 	St dblStartTime 
' Line #135:
' 	StartForVariable 
' 	Next 
' Line #136:
' 	StartForVariable 
' 	Next 
' Line #137:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0001 
' 	LitDI2 0x0020 
' 	For 
' Line #138:
' 	Ld j 
' 	LitDI2 0x0001 
' 	Sub 
' 	ArgsLd _B_var_arrHistTmp 0x0001 
' 	Ld vAmpKeys 
' 	FnUBound 0x0000 
' 	LitDI2 0x0003 
' 	Add 
' 	Paren 
' 	Ld iAmpIndex 
' 	Sub 
' 	Ld iFreqIndex 
' 	LitDI2 0x0002 
' 	Add 
' 	Ld j 
' 	LitDI2 0x0001 
' 	Sub 
' 	ArgsLd theWorksheets 0x0001 
' 	ArgsMemLd Cells 0x0002 
' 	MemSt Value 
' Line #139:
' 	LitDI2 0x0000 
' 	Ld j 
' 	LitDI2 0x0001 
' 	Sub 
' 	ArgsSt _B_var_arrHistTmp 0x0001 
' Line #140:
' 	StartForVariable 
' 	Next 
' Line #141:
' 	StartForVariable 
' 	Next 
' Line #142:
' 	StartForVariable 
' 	Next 
' Line #143:
' Line #144:
' 	Ld theWorksheets 
' 	Ld vFreqKeys 
' 	Ld vAmpKeys 
' 	ArgsCall (Call) SubwriteAxes 0x0003 
' Line #145:
' Line #146:
' 	Ld objttx 
' 	ArgsMemCall (Call) CloseTank 0x0000 
' Line #147:
' 	Ld objttx 
' 	ArgsMemCall (Call) ReleaseServer 0x0000 
' Line #148:
' 	EndSub 
' Line #149:
' Line #150:
' 	FuncDefn (Function buildWorksheetArray() As Variant)
' Line #151:
' 	Dim 
' 	OptionBase 
' 	LitDI2 0x001F 
' 	VarDefn theWorksheets
' Line #152:
' 	Dim 
' 	VarDefn strWsname (As String)
' Line #153:
' 	Dim 
' 	VarDefn intWSNum (As Long)
' Line #154:
' Line #155:
' 	Dim 
' 	VarDefn i (As Integer)
' Line #156:
' Line #157:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0001 
' 	Ld Worksheets 
' 	MemLd Count 
' 	For 
' Line #158:
' 	Ld i 
' 	Ld Worksheets 
' 	ArgsMemLd Item 0x0001 
' 	MemLd Name 
' 	St strWsname 
' Line #159:
' 	Ld strWsname 
' 	LitDI2 0x0004 
' 	ArgsLd Left 0x0002 
' 	LitStr 0x0004 "Site"
' 	Eq 
' 	IfBlock 
' Line #160:
' 	Ld strWsname 
' 	Ld strWsname 
' 	FnLen 
' 	LitDI2 0x0004 
' 	Sub 
' 	ArgsLd Right 0x0002 
' 	ArgsLd IsNumeric 0x0001 
' 	IfBlock 
' Line #161:
' 	Ld strWsname 
' 	Ld strWsname 
' 	FnLen 
' 	LitDI2 0x0004 
' 	Sub 
' 	ArgsLd Right 0x0002 
' 	Coerce (Int) 
' 	St intWSNum 
' Line #162:
' 	Ld intWSNum 
' 	LitDI2 0x0021 
' 	Lt 
' 	Ld intWSNum 
' 	LitDI2 0x0000 
' 	Gt 
' 	And 
' 	IfBlock 
' Line #163:
' 	SetStmt 
' 	Ld i 
' 	Ld Worksheets 
' 	ArgsMemLd Item 0x0001 
' 	Ld intWSNum 
' 	LitDI2 0x0001 
' 	Sub 
' 	ArgsSet theWorksheets 0x0001 
' Line #164:
' 	EndIfBlock 
' Line #165:
' 	EndIfBlock 
' Line #166:
' 	EndIfBlock 
' Line #167:
' 	StartForVariable 
' 	Next 
' Line #168:
' Line #169:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	LitDI2 0x001F 
' 	For 
' Line #170:
' 	Ld i 
' 	ArgsLd theWorksheets 0x0001 
' 	ArgsLd Sheet70 0x0001 
' 	IfBlock 
' Line #171:
' 	Ld i 
' 	LitDI2 0x0000 
' 	Gt 
' 	IfBlock 
' Line #172:
' 	SetStmt 
' 	ParamOmitted 
' 	Ld i 
' 	LitDI2 0x0001 
' 	Sub 
' 	ArgsLd theWorksheets 0x0001 
' 	LitDI2 0x0001 
' 	Ld xlWorksheet 
' 	Ld Worksheets 
' 	ArgsMemLd Add 0x0004 
' 	Ld i 
' 	ArgsSet theWorksheets 0x0001 
' Line #173:
' 	ElseBlock 
' Line #174:
' 	SetStmt 
' 	ParamOmitted 
' 	Ld Worksheets 
' 	MemLd Count 
' 	Ld Worksheets 
' 	ArgsMemLd Item 0x0001 
' 	LitDI2 0x0001 
' 	Ld xlWorksheet 
' 	Ld Worksheets 
' 	ArgsMemLd Add 0x0004 
' 	Ld i 
' 	ArgsSet theWorksheets 0x0001 
' Line #175:
' 	EndIfBlock 
' Line #176:
' 	LitStr 0x0004 "Site"
' 	Ld i 
' 	LitDI2 0x0001 
' 	Add 
' 	Coerce (Str) 
' 	Concat 
' 	Ld i 
' 	ArgsLd theWorksheets 0x0001 
' 	MemSt Name 
' Line #177:
' 	EndIfBlock 
' Line #178:
' 	StartForVariable 
' 	Next 
' Line #179:
' 	Ld theWorksheets 
' 	St buildWorksheetArray 
' Line #180:
' 	EndFunc 
' Line #181:
' Line #182:
' 	FuncDefn (Sub SubwriteAxes(theWorksheets As Variant, rowLabels As Variant, deleteWorksheets As Variant))
' Line #183:
' 	Dim 
' 	VarDefn i (As Long)
' Line #184:
' 	Dim 
' 	VarDefn j (As Long)
' Line #185:
' Line #186:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld theWorksheets 
' 	FnUBound 0x0000 
' 	For 
' Line #187:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld deleteWorksheets 
' 	FnUBound 0x0000 
' 	For 
' Line #188:
' 	Ld j 
' 	ArgsLd deleteWorksheets 0x0001 
' 	Ld deleteWorksheets 
' 	FnUBound 0x0000 
' 	LitDI2 0x0003 
' 	Add 
' 	Paren 
' 	Ld j 
' 	Sub 
' 	LitDI2 0x0001 
' 	Ld i 
' 	ArgsLd theWorksheets 0x0001 
' 	ArgsMemLd Cells 0x0002 
' 	MemSt Value 
' Line #189:
' 	StartForVariable 
' 	Next 
' Line #190:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld rowLabels 
' 	FnUBound 0x0000 
' 	For 
' Line #191:
' 	Ld j 
' 	ArgsLd rowLabels 0x0001 
' 	LitDI2 0x0002 
' 	Ld j 
' 	LitDI2 0x0002 
' 	Add 
' 	Ld i 
' 	ArgsLd theWorksheets 0x0001 
' 	ArgsMemLd Cells 0x0002 
' 	MemSt Value 
' Line #192:
' 	StartForVariable 
' 	Next 
' Line #193:
' 	StartForVariable 
' 	Next 
' Line #194:
' Line #195:
' 	EndSub 
' Line #196:
' Line #197:
' 	FuncDefn (Sub Delete())
' Line #198:
' 	Dim 
' 	VarDefn strWsname (As String)
' Line #199:
' 	Dim 
' 	VarDefn intWSNum (As Long)
' Line #200:
' Line #201:
' 	Dim 
' 	VarDefn i (As Integer)
' Line #202:
' Line #203:
' 	Ld Worksheets 
' 	MemLd Count 
' 	St i 
' Line #204:
' Line #205:
' 	Do 
' Line #206:
' 	Ld i 
' 	LitDI2 0x0000 
' 	Eq 
' 	IfBlock 
' Line #207:
' 	ExitDo 
' Line #208:
' 	EndIfBlock 
' Line #209:
' 	Ld i 
' 	Ld Worksheets 
' 	ArgsMemLd Item 0x0001 
' 	MemLd Name 
' 	St strWsname 
' Line #210:
' 	Ld strWsname 
' 	LitDI2 0x0004 
' 	ArgsLd Left 0x0002 
' 	LitStr 0x0004 "Site"
' 	Eq 
' 	IfBlock 
' Line #211:
' 	Ld strWsname 
' 	Ld strWsname 
' 	FnLen 
' 	LitDI2 0x0004 
' 	Sub 
' 	ArgsLd Right 0x0002 
' 	ArgsLd IsNumeric 0x0001 
' 	IfBlock 
' Line #212:
' 	Ld strWsname 
' 	Ld strWsname 
' 	FnLen 
' 	LitDI2 0x0004 
' 	Sub 
' 	ArgsLd Right 0x0002 
' 	Coerce (Int) 
' 	St intWSNum 
' Line #213:
' 	Ld intWSNum 
' 	LitDI2 0x0021 
' 	Lt 
' 	Ld intWSNum 
' 	LitDI2 0x0000 
' 	Gt 
' 	And 
' 	IfBlock 
' Line #214:
' 	Ld i 
' 	Ld Worksheets 
' 	ArgsMemLd Item 0x0001 
' 	ArgsMemCall UserForm1 0x0000 
' Line #215:
' 	EndIfBlock 
' Line #216:
' 	EndIfBlock 
' Line #217:
' 	EndIfBlock 
' Line #218:
' 	Ld i 
' 	LitDI2 0x0001 
' 	Sub 
' 	St i 
' Line #219:
' 	Loop 
' Line #220:
' 	EndSub 
