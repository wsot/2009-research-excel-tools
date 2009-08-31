' Processing file: Tuning curve generator - sigmaplot integration - backup working 20090815.xls
' ===============================================================================
' Module streams:
' _VBA_PROJECT_CUR/VBA/ThisWorkbook - 1210 bytes
' Line #0:
' 	Option  (Explicit)
' Line #1:
' _VBA_PROJECT_CUR/VBA/Sheet1 - 1150 bytes
' _VBA_PROJECT_CUR/VBA/ImportFrom - 5491 bytes
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
' _VBA_PROJECT_CUR/VBA/Module1 - 31044 bytes
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
' 	LitVarSpecial (False)
' 	ArgsCall (Call) processImport 0x0001 
' Line #9:
' 	EndIfBlock 
' Line #10:
' 	EndSub 
' Line #11:
' Line #12:
' 	FuncDefn (Sub importIntoSigmaplot())
' Line #13:
' 	Ld ImportFrom 
' 	ArgsMemCall Show 0x0000 
' Line #14:
' Line #15:
' 	Ld doImport 
' 	IfBlock 
' Line #16:
' 	LitVarSpecial (True)
' 	ArgsCall (Call) processImport 0x0001 
' Line #17:
' 	EndIfBlock 
' Line #18:
' 	EndSub 
' Line #19:
' Line #20:
' 	FuncDefn (Sub processImport(spNB As Boolean))
' Line #21:
' 	Dim 
' 	VarDefn lBinWidth (As Double)
' Line #22:
' 	LitStr 0x0002 "B1"
' 	LitStr 0x0008 "Settings"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	St lBinWidth 
' Line #23:
' Line #24:
' 	Dim 
' 	VarDefn lMaxHistHeigh (As Double)
' Line #25:
' 	LitDI2 0x0000 
' 	St lMaxHistHeigh 
' Line #26:
' Line #27:
' 	Dim 
' 	VarDefn theWorksheets (As Variant)
' Line #28:
' 	QuoteRem 0x0004 0x007E "Dim chanHistTmp(32) As Long 'used as a temporary store to build a histogram across multiple 'Swep's before outputting the data"
' Line #29:
' 	Dim 
' 	OptionBase 
' 	LitDI2 0x001F 
' 	VarDefn _B_var_arrHistTmp (As Long)
' Line #30:
' Line #31:
' 	Dim (Const) 
' 	LitDI2 0x0001 
' 	VarDefn _B_var_Const
' Line #32:
' 	Dim (Const) 
' 	LitDI2 0x0000 
' 	VarDefn _B_var_iColOffset
' Line #33:
' Line #34:
' 	Dim 
' 	VarDefn i (As Long)
' Line #35:
' 	Dim 
' 	VarDefn j (As Long)
' Line #36:
' 	Dim 
' 	VarDefn k (As Long)
' Line #37:
' 	Dim 
' 	VarDefn l (As Long)
' Line #38:
' Line #39:
' 	ArgsLd buildWorksheetArray 0x0000 
' 	St theWorksheets 
' Line #40:
' Line #41:
' 	Dim 
' 	VarDefn objttx
' Line #42:
' 	SetStmt 
' 	LitStr 0x0007 "TTank.X"
' 	ArgsLd CreateObject 0x0001 
' 	Set objttx 
' Line #43:
' Line #44:
' 	Ld theServer 
' 	LitStr 0x0002 "Me"
' 	Ld objttx 
' 	ArgsMemLd ConnectServer 0x0002 
' 	LitDI2 0x0001 
' 	Coerce (Lng) 
' 	Ne 
' 	IfBlock 
' Line #45:
' 	LitStr 0x0015 "Connecting to server "
' 	Ld theServer 
' 	Concat 
' 	LitStr 0x0008 " failed."
' 	Concat 
' 	Paren 
' 	ArgsCall MsgBox 0x0001 
' Line #46:
' 	ExitSub 
' Line #47:
' 	EndIfBlock 
' Line #48:
' Line #49:
' 	Ld theTank 
' 	LitStr 0x0001 "R"
' 	Ld objttx 
' 	ArgsMemLd OpenTank 0x0002 
' 	LitDI2 0x0001 
' 	Coerce (Lng) 
' 	Ne 
' 	IfBlock 
' Line #50:
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
' Line #51:
' 	Ld objttx 
' 	ArgsMemCall (Call) ReleaseServer 0x0000 
' Line #52:
' 	ExitSub 
' Line #53:
' 	EndIfBlock 
' Line #54:
' Line #55:
' 	Ld theBlock 
' 	Ld objttx 
' 	ArgsMemLd SelectBlock 0x0001 
' 	LitDI2 0x0001 
' 	Coerce (Lng) 
' 	Ne 
' 	IfBlock 
' Line #56:
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
' Line #57:
' 	Ld objttx 
' 	ArgsMemCall (Call) CloseTank 0x0000 
' Line #58:
' 	Ld objttx 
' 	ArgsMemCall (Call) ReleaseServer 0x0000 
' Line #59:
' 	ExitSub 
' Line #60:
' 	EndIfBlock 
' Line #61:
' Line #62:
' 	Ld objttx 
' 	ArgsMemCall (Call) CreateEpocIndexing 0x0000 
' Line #63:
' Line #64:
' 	Dim 
' 	VarDefn freqList (As Dictionary)
' Line #65:
' 	Dim 
' 	VarDefn ampList (As Dictionary)
' Line #66:
' Line #67:
' 	SetStmt 
' 	New id_FFFF
' 	Set freqList 
' Line #68:
' 	SetStmt 
' 	New id_FFFF
' 	Set ampList 
' Line #69:
' Line #70:
' 	Dim 
' 	VarDefn dblStartTime (As Double)
' Line #71:
' 	Dim 
' 	VarDefn dblEndTime (As Double)
' Line #72:
' 	Dim 
' 	VarDefn varReturn (As Variant)
' Line #73:
' 	Dim 
' 	VarDefn varAmp (As Variant)
' Line #74:
' Line #75:
' Line #76:
' 	Do 
' Line #77:
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
' Line #78:
' 	Ld i 
' 	LitDI2 0x0000 
' 	Eq 
' 	IfBlock 
' Line #79:
' 	ExitDo 
' Line #80:
' 	EndIfBlock 
' Line #81:
' Line #82:
' 	LitDI2 0x0000 
' 	Ld i 
' 	LitDI2 0x0000 
' 	Ld objttx 
' 	ArgsMemLd ParseEvInfoV 0x0003 
' 	St varReturn 
' Line #83:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld i 
' 	LitDI2 0x0001 
' 	Sub 
' 	Paren 
' 	For 
' Line #84:
' 	LitDI2 0x0006 
' 	Ld j 
' 	ArgsLd varReturn 0x0002 
' 	Ld freqList 
' 	ArgsMemLd Exists 0x0001 
' 	Not 
' 	IfBlock 
' Line #85:
' 	LitDI2 0x0006 
' 	Ld j 
' 	ArgsLd varReturn 0x0002 
' 	LitStr 0x0000 ""
' 	Ld freqList 
' 	ArgsMemCall (Call) Add 0x0002 
' Line #86:
' 	EndIfBlock 
' Line #87:
' 	LitDI2 0x0005 
' 	Ld j 
' 	ArgsLd varReturn 0x0002 
' 	LitDI2 0x0001 
' 	LitDI4 0x86A0 0x0001 
' 	Div 
' 	Paren 
' 	Add 
' 	St dblStartTime 
' Line #88:
' 	LitStr 0x0004 "Lev1"
' 	LitDI2 0x0005 
' 	Ld j 
' 	ArgsLd varReturn 0x0002 
' 	LitDI2 0x0000 
' 	Ld objttx 
' 	ArgsMemLd QryEpocAtV 0x0003 
' 	St varAmp 
' Line #89:
' 	Ld varAmp 
' 	Ld ampList 
' 	ArgsMemLd Exists 0x0001 
' 	Not 
' 	IfBlock 
' Line #90:
' 	Ld varAmp 
' 	LitStr 0x0000 ""
' 	Ld ampList 
' 	ArgsMemCall (Call) Add 0x0002 
' Line #91:
' 	EndIfBlock 
' Line #92:
' Line #93:
' 	StartForVariable 
' 	Next 
' Line #94:
' Line #95:
' 	Ld i 
' 	LitDI2 0x01F4 
' 	Lt 
' 	IfBlock 
' Line #96:
' 	ExitDo 
' Line #97:
' 	EndIfBlock 
' Line #98:
' 	Loop 
' Line #99:
' Line #100:
' 	LitDI2 0x0000 
' 	St i 
' Line #101:
' 	LitDI2 0x0000 
' 	St j 
' Line #102:
' Line #103:
' 	QuoteRem 0x0000 0x0016 "    Dim freqAmpArray()"
' Line #104:
' 	QuoteRem 0x0000 0x003F "    Dim freqAmpArray(freqList.Count - 1, ampList.Count - 1, 32)"
' Line #105:
' 	Dim 
' 	VarDefn iFreqIndex (As Integer)
' Line #106:
' 	Dim 
' 	VarDefn iAmpIndex (As Integer)
' Line #107:
' Line #108:
' 	Dim 
' 	VarDefn vFreqKeys (As Variant)
' Line #109:
' 	Dim 
' 	VarDefn vAmpKeys (As Variant)
' Line #110:
' Line #111:
' Line #112:
' 	Ld freqList 
' 	MemLd Keys 
' 	St vFreqKeys 
' Line #113:
' 	Ld ampList 
' 	MemLd Keys 
' 	St vAmpKeys 
' Line #114:
' Line #115:
' 	Dim 
' 	VarDefn varChanData (As Variant)
' Line #116:
' 	Dim 
' 	VarDefn IsEmpty (As Double)
' Line #117:
' Line #118:
' 	StartForVariable 
' 	Ld iFreqIndex 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld vFreqKeys 
' 	FnUBound 0x0000 
' 	For 
' Line #119:
' 	StartForVariable 
' 	Ld iAmpIndex 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld vAmpKeys 
' 	FnUBound 0x0000 
' 	For 
' Line #120:
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
' Line #121:
' 	LitStr 0x0004 "Swep"
' 	LitDI2 0x0000 
' 	Ld objttx 
' 	ArgsMemLd GetEpocsExV 0x0002 
' 	St varReturn 
' Line #122:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld varReturn 
' 	LitDI2 0x0002 
' 	FnUBound 0x0001 
' 	For 
' Line #123:
' 	LitDI2 0x0002 
' 	Ld i 
' 	ArgsLd varReturn 0x0002 
' 	St dblStartTime 
' Line #124:
' 	Ld dblStartTime 
' 	Ld lBinWidth 
' 	Add 
' 	St dblEndTime 
' Line #125:
' 	QuoteRem 0x0010 0x001C "dblEndTime = varReturn(3, i)"
' Line #126:
' 	Ld dblStartTime 
' 	St IsEmpty 
' Line #127:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0001 
' 	LitDI2 0x0020 
' 	For 
' Line #128:
' 	Do 
' Line #129:
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
' Line #130:
' 	Ld k 
' 	LitDI2 0x0000 
' 	Eq 
' 	IfBlock 
' Line #131:
' 	ExitDo 
' Line #132:
' 	EndIfBlock 
' Line #133:
' Line #134:
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
' Line #135:
' Line #136:
' 	QuoteRem 0x0000 0x0042 "                        varChanData = objttx.ParseEvInfoV(0, k, 6)"
' Line #137:
' 	QuoteRem 0x0000 0x002C "                        For l = 0 To (k - 1)"
' Line #138:
' 	QuoteRem 0x0000 0x0059 "                            Worksheets.Item("Settings").Cells(iAmpIndex + 3, 1).Value = j"
' Line #139:
' 	QuoteRem 0x0000 0x0066 "                            Worksheets.Item("Settings").Cells(iAmpIndex + 3, 2).Value = varChanData(0)"
' Line #140:
' 	QuoteRem 0x0000 0x001C "                        Next"
' Line #141:
' Line #142:
' 	Ld k 
' 	LitDI2 0x01F4 
' 	Lt 
' 	IfBlock 
' Line #143:
' 	ExitDo 
' Line #144:
' 	ElseBlock 
' Line #145:
' 	Ld k 
' 	LitDI2 0x0001 
' 	Sub 
' 	LitDI2 0x0001 
' 	LitDI2 0x0006 
' 	Ld objttx 
' 	ArgsMemLd ParseEvInfoV 0x0003 
' 	St varChanData 
' Line #146:
' 	LitDI2 0x0000 
' 	ArgsLd varChanData 0x0001 
' 	LitDI2 0x0001 
' 	LitDI4 0x86A0 0x0001 
' 	Div 
' 	Paren 
' 	Add 
' 	St dblStartTime 
' Line #147:
' 	EndIfBlock 
' Line #148:
' 	Loop 
' Line #149:
' 	Ld IsEmpty 
' 	St dblStartTime 
' Line #150:
' 	StartForVariable 
' 	Next 
' Line #151:
' 	StartForVariable 
' 	Next 
' Line #152:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0001 
' 	LitDI2 0x0020 
' 	For 
' Line #153:
' 	Ld j 
' 	LitDI2 0x0001 
' 	Sub 
' 	ArgsLd _B_var_arrHistTmp 0x0001 
' 	Ld vAmpKeys 
' 	FnUBound 0x0000 
' 	Ld _B_var_Const 
' 	Add 
' 	LitDI2 0x0002 
' 	Add 
' 	Paren 
' 	Ld iAmpIndex 
' 	Sub 
' 	Ld iFreqIndex 
' 	Ld _B_var_iColOffset 
' 	Add 
' 	LitDI2 0x0002 
' 	Add 
' 	Ld j 
' 	LitDI2 0x0001 
' 	Sub 
' 	ArgsLd theWorksheets 0x0001 
' 	ArgsMemLd Cells 0x0002 
' 	MemSt Value 
' Line #154:
' 	Ld j 
' 	LitDI2 0x0001 
' 	Sub 
' 	ArgsLd _B_var_arrHistTmp 0x0001 
' 	Ld lMaxHistHeigh 
' 	Gt 
' 	IfBlock 
' Line #155:
' 	Ld j 
' 	LitDI2 0x0001 
' 	Sub 
' 	ArgsLd _B_var_arrHistTmp 0x0001 
' 	St lMaxHistHeigh 
' Line #156:
' 	EndIfBlock 
' Line #157:
' Line #158:
' 	LitDI2 0x0000 
' 	Ld j 
' 	LitDI2 0x0001 
' 	Sub 
' 	ArgsSt _B_var_arrHistTmp 0x0001 
' Line #159:
' 	StartForVariable 
' 	Next 
' Line #160:
' 	StartForVariable 
' 	Next 
' Line #161:
' 	StartForVariable 
' 	Next 
' Line #162:
' Line #163:
' 	Ld theWorksheets 
' 	Ld vFreqKeys 
' 	Ld vAmpKeys 
' 	Ld _B_var_iColOffset 
' 	Ld _B_var_Const 
' 	ArgsCall (Call) SubwriteAxes 0x0005 
' Line #164:
' Line #165:
' 	Ld objttx 
' 	ArgsMemCall (Call) CloseTank 0x0000 
' Line #166:
' 	Ld objttx 
' 	ArgsMemCall (Call) ReleaseServer 0x0000 
' Line #167:
' Line #168:
' 	Ld spNB 
' 	IfBlock 
' Line #169:
' 	Ld theWorksheets 
' 	Ld vFreqKeys 
' 	Ld vAmpKeys 
' 	Ld _B_var_iColOffset 
' 	Ld _B_var_Const 
' 	Ld lMaxHistHeigh 
' 	ArgsCall (Call) ACTIVESPWLib 0x0006 
' Line #170:
' 	EndIfBlock 
' Line #171:
' Line #172:
' 	EndSub 
' Line #173:
' Line #174:
' 	FuncDefn (Function buildWorksheetArray() As Variant)
' Line #175:
' 	Dim 
' 	OptionBase 
' 	LitDI2 0x001F 
' 	VarDefn theWorksheets
' Line #176:
' 	Dim 
' 	VarDefn strWsname (As String)
' Line #177:
' 	Dim 
' 	VarDefn intWSNum (As Long)
' Line #178:
' Line #179:
' 	Dim 
' 	VarDefn i (As Integer)
' Line #180:
' Line #181:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0001 
' 	Ld Worksheets 
' 	MemLd Count 
' 	For 
' Line #182:
' 	Ld i 
' 	Ld Worksheets 
' 	ArgsMemLd Item 0x0001 
' 	MemLd Name 
' 	St strWsname 
' Line #183:
' 	Ld strWsname 
' 	LitDI2 0x0004 
' 	ArgsLd Left 0x0002 
' 	LitStr 0x0004 "Site"
' 	Eq 
' 	IfBlock 
' Line #184:
' 	Ld strWsname 
' 	Ld strWsname 
' 	FnLen 
' 	LitDI2 0x0004 
' 	Sub 
' 	ArgsLd Right 0x0002 
' 	ArgsLd IsNumeric 0x0001 
' 	IfBlock 
' Line #185:
' 	Ld strWsname 
' 	Ld strWsname 
' 	FnLen 
' 	LitDI2 0x0004 
' 	Sub 
' 	ArgsLd Right 0x0002 
' 	Coerce (Int) 
' 	St intWSNum 
' Line #186:
' 	Ld intWSNum 
' 	LitDI2 0x0021 
' 	Lt 
' 	Ld intWSNum 
' 	LitDI2 0x0000 
' 	Gt 
' 	And 
' 	IfBlock 
' Line #187:
' 	SetStmt 
' 	Ld i 
' 	Ld Worksheets 
' 	ArgsMemLd Item 0x0001 
' 	Ld intWSNum 
' 	LitDI2 0x0001 
' 	Sub 
' 	ArgsSet theWorksheets 0x0001 
' Line #188:
' 	EndIfBlock 
' Line #189:
' 	EndIfBlock 
' Line #190:
' 	EndIfBlock 
' Line #191:
' 	StartForVariable 
' 	Next 
' Line #192:
' Line #193:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	LitDI2 0x001F 
' 	For 
' Line #194:
' 	Ld i 
' 	ArgsLd theWorksheets 0x0001 
' 	ArgsLd Sheet70 0x0001 
' 	IfBlock 
' Line #195:
' 	Ld i 
' 	LitDI2 0x0000 
' 	Gt 
' 	IfBlock 
' Line #196:
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
' Line #197:
' 	ElseBlock 
' Line #198:
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
' Line #199:
' 	EndIfBlock 
' Line #200:
' 	LitStr 0x0004 "Site"
' 	Ld i 
' 	LitDI2 0x0001 
' 	Add 
' 	Coerce (Str) 
' 	Concat 
' 	Ld i 
' 	ArgsLd theWorksheets 0x0001 
' 	MemSt Name 
' Line #201:
' 	EndIfBlock 
' Line #202:
' 	StartForVariable 
' 	Next 
' Line #203:
' 	Ld theWorksheets 
' 	St buildWorksheetArray 
' Line #204:
' 	EndFunc 
' Line #205:
' Line #206:
' 	FuncDefn (Sub SubwriteAxes(theWorksheets As Variant, rowLabels As Variant, deleteWorksheets As Variant, _B_var_iColOffset, _B_var_Const))
' Line #207:
' 	Dim 
' 	VarDefn i (As Long)
' Line #208:
' 	Dim 
' 	VarDefn j (As Long)
' Line #209:
' Line #210:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld theWorksheets 
' 	FnUBound 0x0000 
' 	For 
' Line #211:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld deleteWorksheets 
' 	FnUBound 0x0000 
' 	For 
' Line #212:
' 	Ld j 
' 	ArgsLd deleteWorksheets 0x0001 
' 	Ld deleteWorksheets 
' 	FnUBound 0x0000 
' 	Ld _B_var_Const 
' 	Add 
' 	LitDI2 0x0002 
' 	Add 
' 	Paren 
' 	Ld j 
' 	Sub 
' 	Ld _B_var_iColOffset 
' 	LitDI2 0x0001 
' 	Add 
' 	Ld i 
' 	ArgsLd theWorksheets 0x0001 
' 	ArgsMemLd Cells 0x0002 
' 	MemSt Value 
' Line #213:
' 	StartForVariable 
' 	Next 
' Line #214:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld rowLabels 
' 	FnUBound 0x0000 
' 	For 
' Line #215:
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
' Line #216:
' 	StartForVariable 
' 	Next 
' Line #217:
' 	StartForVariable 
' 	Next 
' Line #218:
' Line #219:
' 	EndSub 
' Line #220:
' Line #221:
' 	FuncDefn (Sub Delete())
' Line #222:
' 	Dim 
' 	VarDefn strWsname (As String)
' Line #223:
' 	Dim 
' 	VarDefn intWSNum (As Long)
' Line #224:
' Line #225:
' 	Dim 
' 	VarDefn i (As Integer)
' Line #226:
' Line #227:
' 	Ld Worksheets 
' 	MemLd Count 
' 	St i 
' Line #228:
' Line #229:
' 	Do 
' Line #230:
' 	Ld i 
' 	LitDI2 0x0000 
' 	Eq 
' 	IfBlock 
' Line #231:
' 	ExitDo 
' Line #232:
' 	EndIfBlock 
' Line #233:
' 	Ld i 
' 	Ld Worksheets 
' 	ArgsMemLd Item 0x0001 
' 	MemLd Name 
' 	St strWsname 
' Line #234:
' 	Ld strWsname 
' 	LitDI2 0x0004 
' 	ArgsLd Left 0x0002 
' 	LitStr 0x0004 "Site"
' 	Eq 
' 	IfBlock 
' Line #235:
' 	Ld strWsname 
' 	Ld strWsname 
' 	FnLen 
' 	LitDI2 0x0004 
' 	Sub 
' 	ArgsLd Right 0x0002 
' 	ArgsLd IsNumeric 0x0001 
' 	IfBlock 
' Line #236:
' 	Ld strWsname 
' 	Ld strWsname 
' 	FnLen 
' 	LitDI2 0x0004 
' 	Sub 
' 	ArgsLd Right 0x0002 
' 	Coerce (Int) 
' 	St intWSNum 
' Line #237:
' 	Ld intWSNum 
' 	LitDI2 0x0021 
' 	Lt 
' 	Ld intWSNum 
' 	LitDI2 0x0000 
' 	Gt 
' 	And 
' 	IfBlock 
' Line #238:
' 	Ld i 
' 	Ld Worksheets 
' 	ArgsMemLd Item 0x0001 
' 	ArgsMemCall UserForm1 0x0000 
' Line #239:
' 	EndIfBlock 
' Line #240:
' 	EndIfBlock 
' Line #241:
' 	EndIfBlock 
' Line #242:
' 	Ld i 
' 	LitDI2 0x0001 
' 	Sub 
' 	St i 
' Line #243:
' 	Loop 
' Line #244:
' 	EndSub 
' Line #245:
' Line #246:
' 	FuncDefn (Sub ACTIVESPWLib(theWorksheets As Variant, rowLabels As Variant, deleteWorksheets As Variant, _B_var_iColOffset, _B_var_Const, lMaxHistHeigh))
' Line #247:
' Line #248:
' 	Dim (Const) 
' 	LitHI2 0x0406 
' 	VarDefn SAA_TOVAL
' Line #249:
' 	Dim (Const) 
' 	LitHI2 0x0407 
' 	VarDefn GraphPages
' Line #250:
' 	Dim (Const) 
' 	LitHI2 0x0301 
' 	VarDefn SLA_SELECTDIM
' Line #251:
' 	Dim (Const) 
' 	LitDI2 0x0401 
' 	VarDefn SEA_COLORCOL
' Line #252:
' 	Dim (Const) 
' 	LitDI2 0x0308 
' 	VarDefn SAA_OPTIONS
' Line #253:
' 	Dim (Const) 
' 	LitDI2 0x0403 
' 	VarDefn _B_var_GPM_SETPLOTATTR
' Line #254:
' 	Dim (Const) 
' 	LitDI2 0x0408 
' 	VarDefn SAA_FROMVAL
' Line #255:
' 	Dim (Const) 
' 	LitDI2 0x0615 
' 	VarDefn GPM_SETAXISATTRSTRING
' Line #256:
' 	Dim (Const) 
' 	LitDI2 0x0613 
' 	VarDefn SLA_CONTOURFILLTYPE
' Line #257:
' 	Dim (Const) 
' 	LitDI2 0x0358 
' 	VarDefn SAA_SELECTLINE
' Line #258:
' 	Dim (Const) 
' 	LitDI2 0x040A 
' 	VarDefn SEA_THICKNESS
' Line #259:
' 	Dim (Const) 
' 	LitDI2 0x0601 
' 	VarDefn SEA_COLOR
' Line #260:
' 	Dim (Const) 
' 	LitDI2 0x0606 
' 	VarDefn _B_var_SEA_THICKNESS
' Line #261:
' 	Dim (Const) 
' 	LitDI2 0x0410 
' 	VarDefn _B_var_SAA_SUB1OPTIONS
' Line #262:
' Line #263:
' 	Dim 
' 	VarDefn Module2 (As Object)
' Line #264:
' 	SetStmt 
' 	LitStr 0x0017 "SigmaPlot.Application.1"
' 	ArgsLd CreateObject 0x0001 
' 	Set Module2 
' Line #265:
' 	LitVarSpecial (True)
' 	Ld Module2 
' 	MemSt Application 
' Line #266:
' 	Ld Module2 
' 	MemLd Notebooks 
' 	MemLd buildTuningCurvesIntoSigmaplot 
' 	ArgsMemCall (Call) Add 0x0000 
' Line #267:
' Line #268:
' 	Dim 
' 	VarDefn i (As Integer)
' Line #269:
' 	Dim 
' 	VarDefn j (As Long)
' Line #270:
' 	Dim 
' 	VarDefn k (As Long)
' Line #271:
' Line #272:
' 	Dim 
' 	VarDefn SPApplication (As Object)
' Line #273:
' 	Dim 
' 	VarDefn spDT (As Object)
' Line #274:
' 	Dim 
' 	VarDefn DataTable (As Object)
' Line #275:
' 	Dim 
' 	VarDefn objSPWizard (As Object)
' Line #276:
' Line #277:
' Line #278:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld theWorksheets 
' 	FnUBound 0x0000 
' 	For 
' Line #279:
' 	SetStmt 
' 	Ld Module2 
' 	MemLd buildTuningCurvesIntoSigmaplot 
' 	MemLd Count 
' 	LitDI2 0x0001 
' 	Sub 
' 	Ld Module2 
' 	MemLd buildTuningCurvesIntoSigmaplot 
' 	ArgsMemLd Item 0x0001 
' 	Set SPApplication 
' Line #280:
' 	SetStmt 
' 	Ld SPApplication 
' 	MemLd SPWNotebookComponentType 
' 	MemLd Count 
' 	LitDI2 0x0001 
' 	Sub 
' 	Ld SPApplication 
' 	MemLd SPWNotebookComponentType 
' 	ArgsMemLd Item 0x0001 
' 	Set spDT 
' Line #281:
' 	Ld i 
' 	ArgsLd theWorksheets 0x0001 
' 	MemLd Name 
' 	Ld spDT 
' 	MemSt Name 
' Line #282:
' 	SetStmt 
' 	Ld spDT 
' 	MemLd Cell 
' 	Set DataTable 
' Line #283:
' Line #284:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld deleteWorksheets 
' 	FnUBound 0x0000 
' 	For 
' Line #285:
' 	Ld j 
' 	ArgsLd deleteWorksheets 0x0001 
' 	LitDI2 0x0001 
' 	Ld j 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #286:
' 	StartForVariable 
' 	Next 
' Line #287:
' Line #288:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld rowLabels 
' 	FnUBound 0x0000 
' 	For 
' Line #289:
' 	Ld j 
' 	ArgsLd rowLabels 0x0001 
' 	LitDI2 0x0000 
' 	Ld j 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #290:
' 	StartForVariable 
' 	Next 
' Line #291:
' Line #292:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld rowLabels 
' 	FnUBound 0x0000 
' 	For 
' Line #293:
' 	StartForVariable 
' 	Ld k 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld deleteWorksheets 
' 	FnUBound 0x0000 
' 	For 
' Line #294:
' 	Ld deleteWorksheets 
' 	FnUBound 0x0000 
' 	Ld _B_var_Const 
' 	Add 
' 	LitDI2 0x0002 
' 	Add 
' 	Paren 
' 	Ld k 
' 	Sub 
' 	Ld j 
' 	Ld _B_var_iColOffset 
' 	Add 
' 	LitDI2 0x0002 
' 	Add 
' 	Ld i 
' 	ArgsLd theWorksheets 0x0001 
' 	ArgsMemLd Cells 0x0002 
' 	MemLd Value 
' 	LitDI2 0x0003 
' 	Ld k 
' 	Add 
' 	Ld j 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #295:
' 	StartForVariable 
' 	Next 
' Line #296:
' 	StartForVariable 
' 	Next 
' Line #297:
' Line #298:
' 	LitStr 0x0011 "@rgb(255,255,255)"
' 	LitDI2 0x0002 
' 	LitDI2 0x0000 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #299:
' 	LitStr 0x000D "@rgb(0,0,255)"
' 	LitDI2 0x0002 
' 	LitDI2 0x0001 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #300:
' 	LitStr 0x000F "@rgb(0,255,255)"
' 	LitDI2 0x0002 
' 	LitDI2 0x0002 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #301:
' 	LitStr 0x000D "@rgb(0,255,0)"
' 	LitDI2 0x0002 
' 	LitDI2 0x0003 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #302:
' 	LitStr 0x000F "@rgb(255,255,0)"
' 	LitDI2 0x0002 
' 	LitDI2 0x0004 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #303:
' 	LitStr 0x000D "@rgb(255,0,0)"
' 	LitDI2 0x0002 
' 	LitDI2 0x0005 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #304:
' Line #305:
' 	QuoteRem 0x0008 0x001E "Call spNB.NotebookItems.Add(2)"
' Line #306:
' 	QuoteRem 0x0008 0x0042 "Set spGRPH = spNB.NotebookItems.Item(spNB.NotebookItems.Count - 1)"
' Line #307:
' Line #308:
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd SPWNotebookComponentType 
' 	ArgsMemCall (Call) Add 0x0001 
' Line #309:
' 	Dim 
' 	OptionBase 
' 	LitDI2 0x0002 
' 	OptionBase 
' 	LitDI2 0x0003 
' 	VarDefn PlotColumnCountArray
' Line #310:
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #311:
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	LitDI2 0x0000 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #312:
' 	LitDI4 0x47FF 0x01E8 
' 	LitDI2 0x0002 
' 	LitDI2 0x0000 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #313:
' 	LitDI2 0x0001 
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #314:
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	LitDI2 0x0001 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #315:
' 	LitDI4 0x47FF 0x01E8 
' 	LitDI2 0x0002 
' 	LitDI2 0x0001 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #316:
' 	LitDI2 0x0003 
' 	LitDI2 0x0000 
' 	LitDI2 0x0002 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #317:
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	LitDI2 0x0002 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #318:
' 	LitDI4 0x47FF 0x01E8 
' 	LitDI2 0x0002 
' 	LitDI2 0x0002 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #319:
' 	LitDI2 0x0003 
' 	Ld deleteWorksheets 
' 	FnUBound 0x0000 
' 	Add 
' 	LitDI2 0x0000 
' 	LitDI2 0x0003 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #320:
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	LitDI2 0x0003 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #321:
' 	LitDI4 0x47FF 0x01E8 
' 	LitDI2 0x0002 
' 	LitDI2 0x0003 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #322:
' Line #323:
' 	Dim 
' 	VarDefn CurrentPageItem
' Line #324:
' 	OptionBase 
' 	LitDI2 0x0000 
' 	Redim CurrentPageItem 0x0001 (As Variant)
' Line #325:
' Line #326:
' 	LitDI2 0x0004 
' 	LitDI2 0x0000 
' 	ArgsSt CurrentPageItem 0x0001 
' Line #327:
' 	LitStr 0x000C "Contour Plot"
' 	LitStr 0x0013 "Filled Contour Plot"
' 	LitStr 0x0009 "XY Many Z"
' 	Ld PlotColumnCountArray 
' 	Ld CurrentPageItem 
' 	LitStr 0x0011 "Worksheet Columns"
' 	LitStr 0x0012 "Standard Deviation"
' 	LitStr 0x0007 "Degrees"
' 	LitR8 0x0000 0x0000 0x0000 0x0000 
' 	LitR8 0x0000 0x0000 0x8000 0x4076 
' 	ParamOmitted 
' 	LitStr 0x0012 "Standard Deviation"
' 	LitVarSpecial (True)
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) _B_var_CT_GRAPHICPAGE 0x000D 
' Line #328:
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemLd Graphs 0x0001 
' 	ArgsMemLd SelectObject 0x0001 
' 	ArgsMemCall (Call) testSigmaPlot 0x0000 
' Line #329:
' Line #330:
' 	LitStr 0x0006 "Site y"
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemLd Graphs 0x0001 
' 	ArgsMemLd SelectObject 0x0001 
' 	MemSt Name 
' Line #331:
' 	LitStr 0x000B "Attenuation"
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemLd Graphs 0x0001 
' 	ArgsMemLd SelectObject 0x0001 
' 	ArgsMemLd _B_var_SAA_FROMVAL 0x0001 
' 	MemSt Name 
' Line #332:
' 	LitStr 0x0009 "Frequency"
' 	LitDI2 0x0001 
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemLd Graphs 0x0001 
' 	ArgsMemLd SelectObject 0x0001 
' 	ArgsMemLd _B_var_SAA_FROMVAL 0x0001 
' 	MemSt Name 
' Line #333:
' Line #334:
' 	Ld SLA_SELECTDIM 
' 	Ld SAA_OPTIONS 
' 	LitDI2 0x0003 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #335:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_GPM_SETPLOTATTR 
' 	LitDI2 0x000A 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #336:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_GPM_SETPLOTATTR 
' 	LitDI4 0x0001 0x0310 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #337:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_GPM_SETPLOTATTR 
' 	LitDI4 0x0402 0x00C0 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #338:
' 	Ld SAA_FROMVAL 
' 	Ld SAA_TOVAL 
' 	LitStr 0x0001 "0"
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #339:
' 	Ld SAA_FROMVAL 
' 	Ld GraphPages 
' 	Ld lMaxHistHeigh 
' 	Coerce (Str) 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #340:
' Line #341:
' 	Ld SLA_SELECTDIM 
' 	Ld SAA_OPTIONS 
' 	LitDI2 0x0003 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #342:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #343:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_COLOR 
' 	LitDI2 0x000A 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #344:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #345:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SEA_THICKNESS 
' 	LitHI4 0xFFFF 0x00FF 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #346:
' 	Ld SEA_COLORCOL 
' 	Ld GPM_SETAXISATTRSTRING 
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #347:
' 	Ld SEA_COLORCOL 
' 	Ld SLA_CONTOURFILLTYPE 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #348:
' 	Ld SLA_SELECTDIM 
' 	Ld SAA_SELECTLINE 
' 	LitDI2 0x0001 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #349:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0005 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #350:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SEA_THICKNESS 
' 	LitHI4 0xFFFF 0x00FF 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #351:
' 	Ld SEA_COLORCOL 
' 	Ld GPM_SETAXISATTRSTRING 
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #352:
' 	Ld SEA_COLORCOL 
' 	Ld SLA_CONTOURFILLTYPE 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #353:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #354:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SAA_SUB1OPTIONS 
' 	LitDI2 0x0512 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #355:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SAA_SUB1OPTIONS 
' 	LitDI2 0x0F31 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #356:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0001 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #357:
' Line #358:
' 	Ld i 
' 	Ld theWorksheets 
' 	FnUBound 0x0000 
' 	Lt 
' 	IfBlock 
' Line #359:
' 	LitDI2 0x0001 
' 	Ld SPApplication 
' 	MemLd SPWNotebookComponentType 
' 	ArgsMemCall (Call) Add 0x0001 
' Line #360:
' 	EndIfBlock 
' Line #361:
' 	StartForVariable 
' 	Next 
' Line #362:
' 	EndSub 
' Line #363:
' Line #364:
' 	FuncDefn (Sub _B_var_lMaxHistHeight())
' Line #365:
' 	Dim (Const) 
' 	LitHI2 0x0406 
' 	VarDefn SAA_TOVAL
' Line #366:
' 	Dim (Const) 
' 	LitHI2 0x0407 
' 	VarDefn GraphPages
' Line #367:
' 	Dim (Const) 
' 	LitHI2 0x0301 
' 	VarDefn SLA_SELECTDIM
' Line #368:
' 	Dim (Const) 
' 	LitDI2 0x0401 
' 	VarDefn SEA_COLORCOL
' Line #369:
' 	Dim (Const) 
' 	LitDI2 0x0308 
' 	VarDefn SAA_OPTIONS
' Line #370:
' 	Dim (Const) 
' 	LitDI2 0x0403 
' 	VarDefn _B_var_GPM_SETPLOTATTR
' Line #371:
' 	Dim (Const) 
' 	LitDI2 0x0408 
' 	VarDefn SAA_FROMVAL
' Line #372:
' 	Dim (Const) 
' 	LitDI2 0x0615 
' 	VarDefn GPM_SETAXISATTRSTRING
' Line #373:
' 	Dim (Const) 
' 	LitDI2 0x0613 
' 	VarDefn SLA_CONTOURFILLTYPE
' Line #374:
' 	Dim (Const) 
' 	LitDI2 0x0358 
' 	VarDefn SAA_SELECTLINE
' Line #375:
' 	Dim (Const) 
' 	LitDI2 0x040A 
' 	VarDefn SEA_THICKNESS
' Line #376:
' 	Dim (Const) 
' 	LitDI2 0x0601 
' 	VarDefn SEA_COLOR
' Line #377:
' 	Dim (Const) 
' 	LitDI2 0x0606 
' 	VarDefn _B_var_SEA_THICKNESS
' Line #378:
' 	Dim (Const) 
' 	LitDI2 0x0410 
' 	VarDefn _B_var_SAA_SUB1OPTIONS
' Line #379:
' Line #380:
' 	Dim 
' 	VarDefn Module2 (As Object)
' Line #381:
' 	SetStmt 
' 	LitStr 0x0017 "SigmaPlot.Application.1"
' 	ArgsLd CreateObject 0x0001 
' 	Set Module2 
' Line #382:
' 	LitVarSpecial (True)
' 	Ld Module2 
' 	MemSt Application 
' Line #383:
' 	LitStr 0x0006 "Site y"
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemLd Graphs 0x0001 
' 	ArgsMemLd SelectObject 0x0001 
' 	MemSt Name 
' Line #384:
' 	LitStr 0x000B "Attenuation"
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemLd Graphs 0x0001 
' 	ArgsMemLd SelectObject 0x0001 
' 	ArgsMemLd _B_var_SAA_FROMVAL 0x0001 
' 	MemSt Name 
' Line #385:
' 	LitStr 0x0009 "Frequency"
' 	LitDI2 0x0001 
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemLd Graphs 0x0001 
' 	ArgsMemLd SelectObject 0x0001 
' 	ArgsMemLd _B_var_SAA_FROMVAL 0x0001 
' 	MemSt Name 
' Line #386:
' Line #387:
' 	Ld SLA_SELECTDIM 
' 	Ld SAA_OPTIONS 
' 	LitDI2 0x0003 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #388:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_GPM_SETPLOTATTR 
' 	LitDI2 0x000A 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #389:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_GPM_SETPLOTATTR 
' 	LitDI4 0x0001 0x0310 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #390:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_GPM_SETPLOTATTR 
' 	LitDI4 0x0402 0x00C0 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #391:
' 	Ld SAA_FROMVAL 
' 	Ld SAA_TOVAL 
' 	LitStr 0x0001 "0"
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #392:
' 	Ld SAA_FROMVAL 
' 	Ld GraphPages 
' 	LitStr 0x0003 "150"
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #393:
' Line #394:
' 	Ld SLA_SELECTDIM 
' 	Ld SAA_OPTIONS 
' 	LitDI2 0x0003 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #395:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #396:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_COLOR 
' 	LitDI2 0x000A 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #397:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #398:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SEA_THICKNESS 
' 	LitHI4 0xFFFF 0x00FF 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #399:
' 	Ld SEA_COLORCOL 
' 	Ld GPM_SETAXISATTRSTRING 
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #400:
' 	Ld SEA_COLORCOL 
' 	Ld SLA_CONTOURFILLTYPE 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #401:
' 	Ld SLA_SELECTDIM 
' 	Ld SAA_SELECTLINE 
' 	LitDI2 0x0001 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #402:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0005 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #403:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SEA_THICKNESS 
' 	LitHI4 0xFFFF 0x00FF 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #404:
' 	Ld SEA_COLORCOL 
' 	Ld GPM_SETAXISATTRSTRING 
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #405:
' 	Ld SEA_COLORCOL 
' 	Ld SLA_CONTOURFILLTYPE 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #406:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #407:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SAA_SUB1OPTIONS 
' 	LitDI2 0x0512 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #408:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SAA_SUB1OPTIONS 
' 	LitDI2 0x0F31 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #409:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0001 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #410:
' Line #411:
' 	EndSub 
' _VBA_PROJECT_CUR/VBA/Sheet130 - 1152 bytes
' _VBA_PROJECT_CUR/VBA/Sheet131 - 1152 bytes
' _VBA_PROJECT_CUR/VBA/Sheet132 - 1152 bytes
' _VBA_PROJECT_CUR/VBA/Sheet133 - 1152 bytes
' _VBA_PROJECT_CUR/VBA/Sheet134 - 1152 bytes
' _VBA_PROJECT_CUR/VBA/Sheet135 - 1152 bytes
' _VBA_PROJECT_CUR/VBA/Sheet136 - 1152 bytes
' _VBA_PROJECT_CUR/VBA/Sheet137 - 1152 bytes
' _VBA_PROJECT_CUR/VBA/Sheet138 - 1152 bytes
' _VBA_PROJECT_CUR/VBA/Sheet139 - 1152 bytes
' _VBA_PROJECT_CUR/VBA/Sheet140 - 1152 bytes
' _VBA_PROJECT_CUR/VBA/Sheet141 - 1152 bytes
' _VBA_PROJECT_CUR/VBA/Sheet142 - 1152 bytes
' _VBA_PROJECT_CUR/VBA/Sheet143 - 1152 bytes
' _VBA_PROJECT_CUR/VBA/Sheet144 - 1152 bytes
' _VBA_PROJECT_CUR/VBA/Sheet145 - 1152 bytes
' _VBA_PROJECT_CUR/VBA/Sheet146 - 1152 bytes
' _VBA_PROJECT_CUR/VBA/Sheet147 - 1152 bytes
' _VBA_PROJECT_CUR/VBA/Sheet148 - 1152 bytes
' _VBA_PROJECT_CUR/VBA/Sheet149 - 1152 bytes
' _VBA_PROJECT_CUR/VBA/Sheet150 - 1152 bytes
' _VBA_PROJECT_CUR/VBA/Sheet151 - 1152 bytes
' _VBA_PROJECT_CUR/VBA/Sheet152 - 1152 bytes
' _VBA_PROJECT_CUR/VBA/Sheet153 - 1152 bytes
' _VBA_PROJECT_CUR/VBA/Sheet154 - 1152 bytes
' _VBA_PROJECT_CUR/VBA/Sheet155 - 1152 bytes
' _VBA_PROJECT_CUR/VBA/Sheet156 - 1152 bytes
' _VBA_PROJECT_CUR/VBA/Sheet157 - 1152 bytes
' _VBA_PROJECT_CUR/VBA/Sheet158 - 1152 bytes
' _VBA_PROJECT_CUR/VBA/Sheet159 - 1152 bytes
' _VBA_PROJECT_CUR/VBA/Sheet160 - 1152 bytes
' _VBA_PROJECT_CUR/VBA/Sheet161 - 1152 bytes
