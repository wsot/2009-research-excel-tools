' Processing file: Tuning curve generator - generic.xls
' ===============================================================================
' Module streams:
' _VBA_PROJECT_CUR/VBA/ThisWorkbook - 1210 bytes
' Line #0:
' 	Option  (Explicit)
' Line #1:
' _VBA_PROJECT_CUR/VBA/Sheet1 - 1150 bytes
' _VBA_PROJECT_CUR/VBA/ImportFrom - 22379 bytes
' Line #0:
' 	Dim 
' 	VarDefn objTTX (As Object)
' Line #1:
' 	Dim (Const) 
' 	LitDI2 0x0001 
' 	VarDefn BlockSelectFail
' Line #2:
' 	Dim (Const) 
' 	LitDI2 0x0001 
' 	VarDefn _B_var_ServerConnectFail
' Line #3:
' 	Dim (Const) 
' 	LitDI2 0x0002 
' 	VarDefn _B_var_TankConnectFail
' Line #4:
' 	Dim (Const) 
' 	LitDI2 0x0003 
' 	VarDefn _B_var_BlockConnectFail
' Line #5:
' Line #6:
' 	FuncDefn (Private Sub Cancel_Click())
' Line #7:
' 	LitVarSpecial (False)
' 	St doImport 
' Line #8:
' 	Ld id_FFFF 
' 	ArgsCall Unload 0x0001 
' 	QuoteRem 0x0015 0x0015 "Unloads the UserForm."
' Line #9:
' 	EndSub 
' Line #10:
' Line #11:
' 	FuncDefn (Private Sub ImportButton_Click())
' Line #12:
' 	Ld BlockSelect1 
' 	MemLd ActiveBlock 
' 	LitStr 0x0000 ""
' 	Ne 
' 	IfBlock 
' Line #13:
' 	LitVarSpecial (True)
' 	St doImport 
' Line #14:
' Line #15:
' 	QuoteRem 0x0008 0x0036 "set global variables to the selected block information"
' Line #16:
' 	Ld BlockSelect1 
' 	MemLd UseServer 
' 	St theServer 
' Line #17:
' 	Ld BlockSelect1 
' 	MemLd UseTank 
' 	St theTank 
' Line #18:
' 	Ld BlockSelect1 
' 	MemLd ActiveBlock 
' 	St theBlock 
' Line #19:
' Line #20:
' 	Ld BlockSelect1 
' 	MemLd UseServer 
' 	LitStr 0x0002 "B1"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #21:
' 	Ld BlockSelect1 
' 	MemLd UseTank 
' 	LitStr 0x0002 "B2"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #22:
' 	Ld BlockSelect1 
' 	MemLd ActiveBlock 
' 	LitStr 0x0002 "B3"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #23:
' Line #24:
' 	QuoteRem 0x0008 0x0031 "store the selected 'axis' and other grouping data"
' Line #25:
' 	Dim 
' 	VarDefn _B_var_dictOtherEp (As Dictionary)
' Line #26:
' 	SetStmt 
' 	New 0
' 	Set _B_var_dictOtherEp 
' Line #27:
' Line #28:
' 	Dim 
' 	VarDefn i (As Integer)
' Line #29:
' 	Dim 
' 	VarDefn j (As Integer)
' Line #30:
' Line #31:
' 	Dim 
' 	VarDefn _B_var_Exists (As Integer)
' Line #32:
' 	LitDI2 0x0009 
' 	St _B_var_Exists 
' Line #33:
' 	LitStr 0x0001 "B"
' 	Ld _B_var_Exists 
' 	Coerce (Str) 
' 	Concat 
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0000 ""
' 	Ne 
' 	While 
' Line #34:
' 	LitStr 0x0000 ""
' 	LitStr 0x0001 "B"
' 	Ld _B_var_Exists 
' 	Coerce (Str) 
' 	Concat 
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #35:
' 	Ld _B_var_Exists 
' 	LitDI2 0x0001 
' 	Add 
' 	St _B_var_Exists 
' Line #36:
' 	Wend 
' Line #37:
' Line #38:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld sOrigXAxis 
' 	MemLd ListIndex 
' 	LitDI2 0x0001 
' 	Sub 
' 	Paren 
' 	For 
' Line #39:
' 	Ld i 
' 	Ld sOrigXAxis 
' 	ArgsMemLd Clear 0x0001 
' 	IfBlock 
' Line #40:
' 	Ld i 
' 	Ld sOrigXAxis 
' 	ArgsMemLd Listn 0x0001 
' 	LitDI2 0x0001 
' 	Ld _B_var_dictOtherEp 
' 	ArgsMemCall (Call) Add 0x0002 
' Line #41:
' 	Ld i 
' 	Ld sOrigXAxis 
' 	ArgsMemLd Listn 0x0001 
' 	LitStr 0x0001 "B"
' 	LitDI2 0x0009 
' 	Ld j 
' 	Add 
' 	Coerce (Str) 
' 	Concat 
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #42:
' 	Ld j 
' 	LitDI2 0x0001 
' 	Add 
' 	St j 
' Line #43:
' 	EndIfBlock 
' Line #44:
' 	StartForVariable 
' 	Next 
' Line #45:
' Line #46:
' 	Ld _B_var_XAxis3 
' 	MemLd Value 
' 	St yAxisEp 
' Line #47:
' 	Ld yAxisEp 
' 	LitStr 0x0002 "B5"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #48:
' 	Ld OtherGroupings 
' 	MemLd Value 
' 	St otherEp 
' Line #49:
' 	Ld otherEp 
' 	LitStr 0x0002 "B6"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #50:
' 	Ld _B_var_dictOtherEp 
' 	MemLd Keys 
' 	St _B_var_arrOtherEp 
' Line #51:
' Line #52:
' 	Ld ReverseY 
' 	MemLd Value 
' 	LitVarSpecial (True)
' 	Eq 
' 	IfBlock 
' Line #53:
' 	LitVarSpecial (True)
' 	St bReverseY 
' Line #54:
' 	LitDI2 0x0001 
' 	LitStr 0x0002 "E1"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #55:
' 	ElseBlock 
' Line #56:
' 	LitVarSpecial (False)
' 	St bReverseY 
' Line #57:
' 	LitDI2 0x0000 
' 	LitStr 0x0002 "E1"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #58:
' 	EndIfBlock 
' Line #59:
' Line #60:
' 	Ld Reverse 
' 	MemLd Value 
' 	LitVarSpecial (True)
' 	Eq 
' 	IfBlock 
' Line #61:
' 	LitVarSpecial (True)
' 	St ReverseX 
' Line #62:
' 	LitDI2 0x0001 
' 	LitStr 0x0002 "E2"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #63:
' 	ElseBlock 
' Line #64:
' 	LitVarSpecial (False)
' 	St ReverseX 
' Line #65:
' 	LitDI2 0x0000 
' 	LitStr 0x0002 "E2"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #66:
' 	EndIfBlock 
' Line #67:
' Line #68:
' 	Ld id_FFFF 
' 	ArgsCall Unload 0x0001 
' Line #69:
' 	ElseBlock 
' Line #70:
' 	LitStr 0x001F "Please select a block to import"
' 	Paren 
' 	ArgsCall MsgBox 0x0001 
' Line #71:
' 	EndIfBlock 
' Line #72:
' 	EndSub 
' Line #73:
' Line #74:
' Line #75:
' 	FuncDefn (Private Sub TankSelect1_TankChanged(ActTank As String, ActServer As String))
' Line #76:
' 	QuoteRem 0x0004 0x0053 "When a different tank is selected, update the list of available blocks for the tank"
' Line #77:
' 	Ld ActServer 
' 	Ld BlockSelect1 
' 	MemSt UseServer 
' Line #78:
' 	Ld ActTank 
' 	Ld BlockSelect1 
' 	MemSt UseTank 
' Line #79:
' 	Ld BlockSelect1 
' 	ArgsMemCall (Call) Refresh 0x0000 
' Line #80:
' 	EndSub 
' Line #81:
' Line #82:
' 	FuncDefn (Private Sub UserForm_Activate())
' Line #83:
' Line #84:
' 	SetStmt 
' 	LitStr 0x0007 "TTank.X"
' 	ArgsLd CreateObject 0x0001 
' 	Set objTTX 
' 	QuoteRem 0x0029 0x0027 "establish connection to TDT Tank engine"
' Line #85:
' Line #86:
' 	QuoteRem 0x0004 0x004B "when the form loads, if tanks etc were already selected then re-select them"
' Line #87:
' 	Ld theServer 
' 	LitStr 0x0000 ""
' 	Ne 
' 	IfBlock 
' Line #88:
' 	Ld theServer 
' 	Ld TankSelect1 
' 	MemSt UseServer 
' Line #89:
' 	Ld theServer 
' 	Ld BlockSelect1 
' 	MemSt UseServer 
' Line #90:
' 	Ld theTank 
' 	LitStr 0x0000 ""
' 	Ne 
' 	IfBlock 
' Line #91:
' 	Ld theTank 
' 	Ld TankSelect1 
' 	MemSt ActiveTank 
' Line #92:
' 	Ld theTank 
' 	Ld BlockSelect1 
' 	MemSt UseTank 
' Line #93:
' 	Ld theBlock 
' 	LitStr 0x0000 ""
' 	Ne 
' 	IfBlock 
' Line #94:
' 	Ld theBlock 
' 	Ld BlockSelect1 
' 	MemSt ActiveBlock 
' Line #95:
' 	Ld theBlock 
' 	Ld theTank 
' 	Ld theServer 
' 	LitVarSpecial (True)
' 	ArgsCall (Call) _B_var_buildOptionLists 0x0004 
' Line #96:
' 	EndIfBlock 
' Line #97:
' 	Ld BlockSelect1 
' 	ArgsMemCall Refresh 0x0000 
' Line #98:
' 	EndIfBlock 
' Line #99:
' 	Ld TankSelect1 
' 	ArgsMemCall Refresh 0x0000 
' Line #100:
' 	EndIfBlock 
' Line #101:
' Line #102:
' 	QuoteRem 0x0004 0x0035 "try to read parameters from the spreadsheet variables"
' Line #103:
' 	Ld theServer 
' 	LitStr 0x0000 ""
' 	Eq 
' 	IfBlock 
' Line #104:
' 	LitStr 0x0002 "B1"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	St theServer 
' Line #105:
' 	LitStr 0x0002 "B2"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	St theTank 
' Line #106:
' 	LitStr 0x0002 "B3"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	St theBlock 
' Line #107:
' 	Ld theServer 
' 	Ld theTank 
' 	Ld theBlock 
' 	ArgsLd strServer 0x0003 
' 	SelectCase 
' Line #108:
' 	Ld BlockSelectFail 
' 	Case 
' 	CaseDone 
' 	BoS 0x0000 
' Line #109:
' 	Ld theServer 
' 	Ld TankSelect1 
' 	MemSt UseServer 
' Line #110:
' 	Ld theTank 
' 	Ld TankSelect1 
' 	MemSt ActiveTank 
' Line #111:
' 	Ld theServer 
' 	Ld BlockSelect1 
' 	MemSt UseServer 
' Line #112:
' 	Ld theTank 
' 	Ld BlockSelect1 
' 	MemSt UseTank 
' Line #113:
' 	Ld theBlock 
' 	Ld BlockSelect1 
' 	MemSt ActiveBlock 
' Line #114:
' 	Ld TankSelect1 
' 	ArgsMemCall Refresh 0x0000 
' Line #115:
' 	Ld BlockSelect1 
' 	ArgsMemCall Refresh 0x0000 
' Line #116:
' 	Ld theBlock 
' 	Ld theTank 
' 	Ld theServer 
' 	LitVarSpecial (True)
' 	ArgsCall (Call) _B_var_buildOptionLists 0x0004 
' Line #117:
' 	Ld _B_var_BlockConnectFail 
' 	Case 
' 	CaseDone 
' 	BoS 0x0000 
' Line #118:
' 	Ld theServer 
' 	Ld TankSelect1 
' 	MemSt UseServer 
' Line #119:
' 	Ld theTank 
' 	Ld TankSelect1 
' 	MemSt ActiveTank 
' Line #120:
' 	Ld TankSelect1 
' 	ArgsMemCall Refresh 0x0000 
' Line #121:
' 	EndSelect 
' Line #122:
' 	EndIfBlock 
' Line #123:
' Line #124:
' 	Ld bReverseY 
' 	LitVarSpecial (True)
' 	Eq 
' 	LitStr 0x0002 "E1"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitDI2 0x0001 
' 	Eq 
' 	Or 
' 	IfBlock 
' Line #125:
' 	LitVarSpecial (True)
' 	Ld ReverseY 
' 	MemSt Value 
' Line #126:
' 	ElseBlock 
' Line #127:
' 	LitVarSpecial (False)
' 	Ld ReverseY 
' 	MemSt Value 
' Line #128:
' 	EndIfBlock 
' Line #129:
' Line #130:
' 	Ld ReverseX 
' 	LitVarSpecial (True)
' 	Eq 
' 	LitStr 0x0002 "E2"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitDI2 0x0001 
' 	Eq 
' 	Or 
' 	IfBlock 
' Line #131:
' 	LitVarSpecial (True)
' 	Ld Reverse 
' 	MemSt Value 
' Line #132:
' 	ElseBlock 
' Line #133:
' 	LitVarSpecial (False)
' 	Ld Reverse 
' 	MemSt Value 
' Line #134:
' 	EndIfBlock 
' Line #135:
' Line #136:
' 	EndSub 
' Line #137:
' Line #138:
' 	FuncDefn (Private Sub ActBlock(Exnd As String, ActTank As String, ActServer As String))
' Line #139:
' 	Ld Exnd 
' 	Ld ActTank 
' 	Ld ActServer 
' 	LitVarSpecial (False)
' 	ArgsCall (Call) _B_var_buildOptionLists 0x0004 
' Line #140:
' 	EndSub 
' Line #141:
' Line #142:
' 	QuoteRem 0x0000 0x0059 "test the connection settings to see if it is possible to connect to the server/tank/block"
' Line #143:
' 	FuncDefn (Function strServer(ActServer, ActTank, Exnd, id_FFFE As Variant))
' Line #144:
' Line #145:
' 	Ld ActServer 
' 	LitStr 0x0002 "Me"
' 	Ld objTTX 
' 	ArgsMemLd ConnectServer 0x0002 
' 	LitDI2 0x0001 
' 	Coerce (Lng) 
' 	Ne 
' 	IfBlock 
' Line #146:
' 	Ld _B_var_ServerConnectFail 
' 	St strServer 
' Line #147:
' 	ExitFunc 
' Line #148:
' 	Ld ActTank 
' 	LitStr 0x0001 "R"
' 	Ld objTTX 
' 	ArgsMemLd OpenTank 0x0002 
' 	LitDI2 0x0001 
' 	Coerce (Lng) 
' 	Ne 
' 	ElseIfBlock 
' Line #149:
' 	Ld objTTX 
' 	ArgsMemCall ReleaseServer 0x0000 
' Line #150:
' 	Ld _B_var_TankConnectFail 
' 	St strServer 
' Line #151:
' 	ExitFunc 
' Line #152:
' 	Ld Exnd 
' 	Ld objTTX 
' 	ArgsMemLd SelectBlock 0x0001 
' 	LitDI2 0x0001 
' 	Coerce (Lng) 
' 	Ne 
' 	ElseIfBlock 
' Line #153:
' 	Ld objTTX 
' 	ArgsMemCall CloseTank 0x0000 
' Line #154:
' 	Ld objTTX 
' 	ArgsMemCall ReleaseServer 0x0000 
' Line #155:
' 	Ld _B_var_BlockConnectFail 
' 	St strServer 
' Line #156:
' 	EndIfBlock 
' Line #157:
' Line #158:
' 	EndFunc 
' Line #159:
' Line #160:
' 	FuncDefn (Sub _B_var_buildOptionLists(Exnd, ActTank, ActServer, _B_var_ElseIf))
' Line #161:
' 	QuoteRem 0x0004 0x0035 "if a different block is selcted, try to connect to it"
' Line #162:
' 	Dim (Const) 
' 	LitHI2 0x0101 
' 	VarDefn _B_var_EVTYPE_STRON
' Line #163:
' Line #164:
' 	Dim 
' 	VarDefn objTTX (As Object)
' Line #165:
' 	SetStmt 
' 	LitStr 0x0007 "TTank.X"
' 	ArgsLd CreateObject 0x0001 
' 	Set objTTX 
' Line #166:
' Line #167:
' 	Ld ActServer 
' 	LitStr 0x0002 "Me"
' 	Ld objTTX 
' 	ArgsMemLd ConnectServer 0x0002 
' 	LitDI2 0x0001 
' 	Coerce (Lng) 
' 	Ne 
' 	IfBlock 
' Line #168:
' 	LitStr 0x0015 "Connecting to server "
' 	Ld theServer 
' 	Concat 
' 	LitStr 0x0008 " failed."
' 	Concat 
' 	Paren 
' 	ArgsCall MsgBox 0x0001 
' Line #169:
' 	ExitSub 
' Line #170:
' 	EndIfBlock 
' Line #171:
' Line #172:
' 	Ld ActTank 
' 	LitStr 0x0001 "R"
' 	Ld objTTX 
' 	ArgsMemLd OpenTank 0x0002 
' 	LitDI2 0x0001 
' 	Coerce (Lng) 
' 	Ne 
' 	IfBlock 
' Line #173:
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
' Line #174:
' 	Ld objTTX 
' 	ArgsMemCall (Call) ReleaseServer 0x0000 
' Line #175:
' 	ExitSub 
' Line #176:
' 	EndIfBlock 
' Line #177:
' Line #178:
' 	Ld Exnd 
' 	Ld objTTX 
' 	ArgsMemLd SelectBlock 0x0001 
' 	LitDI2 0x0001 
' 	Coerce (Lng) 
' 	Ne 
' 	IfBlock 
' Line #179:
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
' Line #180:
' 	Ld objTTX 
' 	ArgsMemCall (Call) CloseTank 0x0000 
' Line #181:
' 	Ld objTTX 
' 	ArgsMemCall (Call) ReleaseServer 0x0000 
' Line #182:
' 	ExitSub 
' Line #183:
' 	EndIfBlock 
' Line #184:
' Line #185:
' 	QuoteRem 0x0004 0x001F "build a list of all event codes"
' Line #186:
' 	Dim 
' 	VarDefn GetEventCodes (As Long)
' Line #187:
' Line #188:
' 	Ld _B_var_EVTYPE_STRON 
' 	Ld objTTX 
' 	ArgsMemLd XAxis3 0x0001 
' 	St GetEventCodes 
' Line #189:
' Line #190:
' 	QuoteRem 0x0004 0x002A "fill the select boxes with the event lists"
' Line #191:
' 	Dim 
' 	VarDefn i (As Integer)
' Line #192:
' Line #193:
' 	Dim 
' 	VarDefn sOrigYAxis (As String)
' Line #194:
' 	Dim 
' 	VarDefn vOrigOtherGroupings (As String)
' Line #195:
' 	Dim 
' 	VarDefn Selected (As Dictionary)
' Line #196:
' 	SetStmt 
' 	New 0
' 	Set Selected 
' Line #197:
' Line #198:
' 	Ld _B_var_ElseIf 
' 	IfBlock 
' Line #199:
' 	Ld yAxisEp 
' 	LitStr 0x0000 ""
' 	Ne 
' 	IfBlock 
' Line #200:
' 	Ld yAxisEp 
' 	St sOrigYAxis 
' Line #201:
' 	Ld otherEp 
' 	St vOrigOtherGroupings 
' Line #202:
' 	LitStr 0x0002 "B5"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0000 ""
' 	Ne 
' 	ElseIfBlock 
' Line #203:
' 	LitStr 0x0002 "B5"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	St sOrigYAxis 
' Line #204:
' 	LitStr 0x0002 "B6"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	St vOrigOtherGroupings 
' Line #205:
' 	EndIfBlock 
' Line #206:
' Line #207:
' 	Dim 
' 	VarDefn _B_var_Exists (As Integer)
' Line #208:
' 	LitDI2 0x0009 
' 	St _B_var_Exists 
' Line #209:
' 	LitStr 0x0001 "B"
' 	Ld _B_var_Exists 
' 	Coerce (Str) 
' 	Concat 
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0000 ""
' 	Ne 
' 	While 
' Line #210:
' 	LitStr 0x0001 "B"
' 	Ld _B_var_Exists 
' 	Coerce (Str) 
' 	Concat 
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	Ld Selected 
' 	ArgsMemLd Exists 0x0001 
' 	Not 
' 	IfBlock 
' Line #211:
' 	LitStr 0x0001 "B"
' 	Ld _B_var_Exists 
' 	Coerce (Str) 
' 	Concat 
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitDI2 0x0001 
' 	Ld Selected 
' 	ArgsMemCall (Call) Add 0x0002 
' Line #212:
' 	EndIfBlock 
' Line #213:
' 	Ld _B_var_Exists 
' 	LitDI2 0x0001 
' 	Add 
' 	St _B_var_Exists 
' Line #214:
' 	Wend 
' Line #215:
' 	ElseBlock 
' Line #216:
' 	Ld _B_var_XAxis3 
' 	MemLd Value 
' 	St sOrigYAxis 
' Line #217:
' 	Ld OtherGroupings 
' 	MemLd Value 
' 	St vOrigOtherGroupings 
' Line #218:
' Line #219:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld sOrigXAxis 
' 	MemLd ListIndex 
' 	LitDI2 0x0001 
' 	Sub 
' 	Paren 
' 	For 
' Line #220:
' 	Ld i 
' 	Ld sOrigXAxis 
' 	ArgsMemLd Clear 0x0001 
' 	IfBlock 
' Line #221:
' 	Ld i 
' 	Ld sOrigXAxis 
' 	ArgsMemLd Listn 0x0001 
' 	LitDI2 0x0001 
' 	Ld Selected 
' 	ArgsMemCall (Call) Add 0x0002 
' Line #222:
' 	EndIfBlock 
' Line #223:
' 	StartForVariable 
' 	Next 
' Line #224:
' 	EndIfBlock 
' Line #225:
' Line #226:
' 	Dim 
' 	VarDefn _B_var_bMatchXAxis (As Boolean)
' Line #227:
' 	LitVarSpecial (False)
' 	St _B_var_bMatchXAxis 
' Line #228:
' 	Dim 
' 	VarDefn XAis (As Boolean)
' Line #229:
' 	LitVarSpecial (False)
' 	St XAis 
' Line #230:
' Line #231:
' 	Ld _B_var_XAxis3 
' 	ArgsMemCall (Call) _B_var_End 0x0000 
' Line #232:
' 	Ld OtherGroupings 
' 	ArgsMemCall (Call) _B_var_End 0x0000 
' Line #233:
' 	Ld sOrigXAxis 
' 	ArgsMemCall (Call) _B_var_End 0x0000 
' Line #234:
' Line #235:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld GetEventCodes 
' 	FnUBound 0x0000 
' 	For 
' Line #236:
' 	Ld i 
' 	ArgsLd GetEventCodes 0x0001 
' 	Ld objTTX 
' 	ArgsMemLd _B_var_CodeToString 0x0001 
' 	Ld i 
' 	Ld _B_var_XAxis3 
' 	ArgsMemCall (Call) EVTYPE_STRON 0x0002 
' Line #237:
' Line #238:
' 	Ld _B_var_bMatchXAxis 
' 	LitVarSpecial (False)
' 	Eq 
' 	Ld i 
' 	ArgsLd GetEventCodes 0x0001 
' 	Ld objTTX 
' 	ArgsMemLd _B_var_CodeToString 0x0001 
' 	LitStr 0x0004 "Frq1"
' 	Eq 
' 	And 
' 	IfBlock 
' 	QuoteRem 0x0057 0x002F "if no item was selected, choose Frq1 as default"
' Line #239:
' 	LitStr 0x0004 "Frq1"
' 	Ld _B_var_XAxis3 
' 	MemSt Value 
' Line #240:
' 	LitVarSpecial (True)
' 	St _B_var_bMatchXAxis 
' Line #241:
' 	Ld i 
' 	ArgsLd GetEventCodes 0x0001 
' 	Ld objTTX 
' 	ArgsMemLd _B_var_CodeToString 0x0001 
' 	Coerce (Str) 
' 	Ld sOrigYAxis 
' 	Coerce (Str) 
' 	Eq 
' 	ElseIfBlock 
' 	QuoteRem 0x0053 0x0044 "if item was selected before changing blocks, keep same name selected"
' Line #242:
' 	Ld sOrigYAxis 
' 	Coerce (Str) 
' 	Ld _B_var_XAxis3 
' 	MemSt Value 
' Line #243:
' 	LitVarSpecial (True)
' 	St _B_var_bMatchXAxis 
' Line #244:
' 	EndIfBlock 
' Line #245:
' 	Ld i 
' 	ArgsLd GetEventCodes 0x0001 
' 	Ld objTTX 
' 	ArgsMemLd _B_var_CodeToString 0x0001 
' 	Ld i 
' 	Ld OtherGroupings 
' 	ArgsMemCall (Call) EVTYPE_STRON 0x0002 
' Line #246:
' 	Ld XAis 
' 	LitVarSpecial (False)
' 	Eq 
' 	Ld i 
' 	ArgsLd GetEventCodes 0x0001 
' 	Ld objTTX 
' 	ArgsMemLd _B_var_CodeToString 0x0001 
' 	LitStr 0x0004 "Lev1"
' 	Eq 
' 	And 
' 	IfBlock 
' 	QuoteRem 0x0057 0x0036 "if no item previously selected, choose Lev1 as default"
' Line #247:
' 	LitStr 0x0004 "Lev1"
' 	Ld OtherGroupings 
' 	MemSt Value 
' Line #248:
' 	LitVarSpecial (True)
' 	St XAis 
' Line #249:
' 	Ld i 
' 	ArgsLd GetEventCodes 0x0001 
' 	Ld objTTX 
' 	ArgsMemLd _B_var_CodeToString 0x0001 
' 	Coerce (Str) 
' 	Ld vOrigOtherGroupings 
' 	Coerce (Str) 
' 	Eq 
' 	ElseIfBlock 
' 	QuoteRem 0x0053 0x0033 "if item was previously selected, try to reselect it"
' Line #250:
' 	Ld vOrigOtherGroupings 
' 	Coerce (Str) 
' 	Ld OtherGroupings 
' 	MemSt Value 
' Line #251:
' 	LitVarSpecial (True)
' 	St XAis 
' Line #252:
' 	EndIfBlock 
' Line #253:
' 	Ld i 
' 	ArgsLd GetEventCodes 0x0001 
' 	Ld objTTX 
' 	ArgsMemLd _B_var_CodeToString 0x0001 
' 	Ld i 
' 	Ld sOrigXAxis 
' 	ArgsMemCall (Call) EVTYPE_STRON 0x0002 
' Line #254:
' 	Ld i 
' 	ArgsLd GetEventCodes 0x0001 
' 	Ld objTTX 
' 	ArgsMemLd _B_var_CodeToString 0x0001 
' 	Ld Selected 
' 	ArgsMemLd Exists 0x0001 
' 	IfBlock 
' Line #255:
' 	LitVarSpecial (True)
' 	Ld i 
' 	Ld sOrigXAxis 
' 	ArgsMemSt Clear 0x0001 
' Line #256:
' 	EndIfBlock 
' Line #257:
' 	StartForVariable 
' 	Next 
' Line #258:
' Line #259:
' 	QuoteRem 0x0004 0x0036 "add the channel option, as it is not actually an epoch"
' Line #260:
' 	LitStr 0x0007 "Channel"
' 	Ld i 
' 	Ld _B_var_XAxis3 
' 	ArgsMemCall (Call) EVTYPE_STRON 0x0002 
' Line #261:
' 	LitStr 0x0007 "Channel"
' 	Ld i 
' 	Ld OtherGroupings 
' 	ArgsMemCall (Call) EVTYPE_STRON 0x0002 
' Line #262:
' 	LitStr 0x0007 "Channel"
' 	Ld i 
' 	Ld sOrigXAxis 
' 	ArgsMemCall (Call) EVTYPE_STRON 0x0002 
' Line #263:
' Line #264:
' 	QuoteRem 0x0004 0x005F "if the defaults were not available, and nothing was selected, choose the first items by default"
' Line #265:
' 	Ld _B_var_bMatchXAxis 
' 	LitVarSpecial (False)
' 	Eq 
' 	IfBlock 
' Line #266:
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	Ld _B_var_XAxis3 
' 	ArgsMemLd Listn 0x0002 
' 	Ld _B_var_XAxis3 
' 	MemSt Value 
' Line #267:
' 	EndIfBlock 
' Line #268:
' 	Ld XAis 
' 	LitVarSpecial (False)
' 	Eq 
' 	IfBlock 
' Line #269:
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	Ld OtherGroupings 
' 	ArgsMemLd Listn 0x0002 
' 	Ld OtherGroupings 
' 	MemSt Value 
' Line #270:
' 	EndIfBlock 
' Line #271:
' Line #272:
' 	Ld objTTX 
' 	ArgsMemCall (Call) CloseTank 0x0000 
' Line #273:
' 	Ld objTTX 
' 	ArgsMemCall (Call) ReleaseServer 0x0000 
' Line #274:
' Line #275:
' 	SetStmt 
' 	LitNothing 
' 	Set Selected 
' Line #276:
' Line #277:
' 	EndSub 
' _VBA_PROJECT_CUR/VBA/Module1 - 50276 bytes
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
' 	Dim (Global) 
' 	VarDefn yAxisEp
' 	VarDefn otherEp
' 	VarDefn _B_var_arrOtherEp
' Line #4:
' 	Dim (Global) 
' 	VarDefn lBinWidth (As Double)
' Line #5:
' 	Dim (Global) 
' 	VarDefn _B_var_lIgnoreFirstMsec (As Double)
' Line #6:
' 	Dim (Global) 
' 	VarDefn _B_var_Const (As Integer)
' Line #7:
' 	Dim (Global) 
' 	VarDefn _B_var_iColOffset (As Integer)
' Line #8:
' 	Dim (Global) 
' 	VarDefn bReverseY
' 	VarDefn ReverseX (As Boolean)
' Line #9:
' Line #10:
' 	Dim (Global) 
' 	VarDefn dHeading (As Dictionary)
' Line #11:
' 	Dim (Global) 
' 	VarDefn bXAxisLog (As Dictionary)
' Line #12:
' 	Dim (Global) 
' 	VarDefn buildHeadingsList (As Boolean)
' Line #13:
' Line #14:
' 	Dim 
' 	VarDefn vYAxisKeys (As Variant)
' Line #15:
' 	Dim 
' 	VarDefn buildOptionLists (As Variant)
' Line #16:
' Line #17:
' Line #18:
' 	FuncDefn (Sub buildTuningCurves())
' Line #19:
' 	Ld ImportFrom 
' 	ArgsMemCall Show 0x0000 
' Line #20:
' Line #21:
' 	Ld doImport 
' 	IfBlock 
' Line #22:
' 	LitVarSpecial (False)
' 	ArgsCall (Call) processImport 0x0001 
' Line #23:
' 	EndIfBlock 
' Line #24:
' 	EndSub 
' Line #25:
' Line #26:
' 	FuncDefn (Sub importIntoSigmaplot())
' Line #27:
' 	QuoteRem 0x0000 0x0013 "    ImportFrom.Show"
' Line #28:
' Line #29:
' 	QuoteRem 0x0000 0x0014 "    If doImport Then"
' Line #30:
' 	QuoteRem 0x0000 0x0020 "        Call processImport(True)"
' Line #31:
' 	QuoteRem 0x0000 0x000A "    End If"
' Line #32:
' 	ArgsCall (Call) ACTIVESPWLib 0x0000 
' Line #33:
' 	EndSub 
' Line #34:
' Line #35:
' 	FuncDefn (Sub processImport(spNB As Boolean))
' Line #36:
' Line #37:
' 	QuoteRem 0x0004 0x002B "load the bin width for histogram generation"
' Line #38:
' 	LitStr 0x0002 "B1"
' 	LitStr 0x0008 "Settings"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	St lBinWidth 
' Line #39:
' Line #40:
' 	QuoteRem 0x0004 0x004D "load the # of msec to ignore at the start (for filtering stimulation artifact"
' Line #41:
' 	LitStr 0x0002 "B2"
' 	LitStr 0x0008 "Settings"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	St _B_var_lIgnoreFirstMsec 
' Line #42:
' Line #43:
' 	QuoteRem 0x0004 0x003A "used to store the maximum histogram peak for normalisation"
' Line #44:
' 	Dim 
' 	VarDefn lMaxHistHeigh (As Double)
' Line #45:
' 	LitDI2 0x0000 
' 	St lMaxHistHeigh 
' Line #46:
' Line #47:
' 	Dim 
' 	VarDefn theWorksheets (As Variant)
' 	QuoteRem 0x0021 0x0029 "stores the created worksheets to write to"
' Line #48:
' 	Dim 
' 	VarDefn _B_var_arrHistTmp (As Long)
' 	QuoteRem 0x001D 0x0044 "used to store the histogram data for each channel as it is generated"
' Line #49:
' 	OptionBase 
' 	LitDI2 0x001F 
' 	Redim _B_var_arrHistTmp 0x0001 (As Variant)
' Line #50:
' Line #51:
' 	Dim 
' 	VarDefn xCount (As Long)
' Line #52:
' 	Dim 
' 	VarDefn yCoun (As Long)
' Line #53:
' 	Dim 
' 	VarDefn lMaxHistHe (As Long)
' Line #54:
' Line #55:
' 	QuoteRem 0x0004 0x0037 "offsets to leave space at the top and left of the chart"
' Line #56:
' 	LitDI2 0x0001 
' 	St _B_var_Const 
' Line #57:
' 	LitDI2 0x0000 
' 	St _B_var_iColOffset 
' Line #58:
' Line #59:
' 	QuoteRem 0x0000 0x0050 "    theWorksheets = buildWorksheetArray() 'build the worksheets for writing data"
' Line #60:
' Line #61:
' 	QuoteRem 0x0004 0x0013 "connect to the tank"
' Line #62:
' 	Dim 
' 	VarDefn objTTX
' Line #63:
' 	SetStmt 
' 	LitStr 0x0007 "TTank.X"
' 	ArgsLd CreateObject 0x0001 
' 	Set objTTX 
' Line #64:
' Line #65:
' 	Ld theServer 
' 	LitStr 0x0002 "Me"
' 	Ld objTTX 
' 	ArgsMemLd ConnectServer 0x0002 
' 	LitDI2 0x0001 
' 	Coerce (Lng) 
' 	Ne 
' 	IfBlock 
' Line #66:
' 	LitStr 0x0015 "Connecting to server "
' 	Ld theServer 
' 	Concat 
' 	LitStr 0x0008 " failed."
' 	Concat 
' 	Paren 
' 	ArgsCall MsgBox 0x0001 
' Line #67:
' 	ExitSub 
' Line #68:
' 	EndIfBlock 
' Line #69:
' Line #70:
' 	Ld theTank 
' 	LitStr 0x0001 "R"
' 	Ld objTTX 
' 	ArgsMemLd OpenTank 0x0002 
' 	LitDI2 0x0001 
' 	Coerce (Lng) 
' 	Ne 
' 	IfBlock 
' Line #71:
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
' Line #72:
' 	Ld objTTX 
' 	ArgsMemCall (Call) ReleaseServer 0x0000 
' Line #73:
' 	ExitSub 
' Line #74:
' 	EndIfBlock 
' Line #75:
' Line #76:
' 	Ld theBlock 
' 	Ld objTTX 
' 	ArgsMemLd SelectBlock 0x0001 
' 	LitDI2 0x0001 
' 	Coerce (Lng) 
' 	Ne 
' 	IfBlock 
' Line #77:
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
' Line #78:
' 	Ld objTTX 
' 	ArgsMemCall (Call) CloseTank 0x0000 
' Line #79:
' 	Ld objTTX 
' 	ArgsMemCall (Call) ReleaseServer 0x0000 
' Line #80:
' 	ExitSub 
' Line #81:
' 	EndIfBlock 
' Line #82:
' Line #83:
' 	QuoteRem 0x0004 0x0026 "index epochs - required to use filters"
' Line #84:
' 	Ld objTTX 
' 	ArgsMemCall (Call) CreateEpocIndexing 0x0000 
' Line #85:
' Line #86:
' 	Dim 
' 	VarDefn dblStartTime (As Double)
' Line #87:
' 	Dim 
' 	VarDefn dblEndTime (As Double)
' Line #88:
' Line #89:
' 	Dim 
' 	VarDefn varReturn (As Variant)
' Line #90:
' Line #91:
' 	QuoteRem 0x0000 0x001D "    Dim vXAxisKeys As Variant"
' Line #92:
' 	QuoteRem 0x0000 0x001D "    Dim vYAxisKeys As Variant"
' Line #93:
' Line #94:
' 	Ld objTTX 
' 	Ld yAxisEp 
' 	Ld bReverseY 
' 	ArgsLd _B_var_buildEpocList 0x0003 
' 	St vYAxisKeys 
' Line #95:
' 	Ld objTTX 
' 	Ld otherEp 
' 	Ld ReverseX 
' 	ArgsLd _B_var_buildEpocList 0x0003 
' 	St buildOptionLists 
' Line #96:
' Line #97:
' 	Dim 
' 	VarDefn i (As Long)
' Line #98:
' 	Dim 
' 	VarDefn j (As Long)
' Line #99:
' 	Dim 
' 	VarDefn k (As Long)
' Line #100:
' 	Dim 
' 	VarDefn l (As Long)
' Line #101:
' Line #102:
' 	Dim 
' 	VarDefn strSearchString (As Variant)
' Line #103:
' 	Ld _B_var_arrOtherEp 
' 	FnUBound 0x0000 
' 	LitDI2 0x0001 
' 	UMi 
' 	Ne 
' 	IfBlock 
' Line #104:
' 	OptionBase 
' 	Ld _B_var_arrOtherEp 
' 	FnUBound 0x0000 
' 	Redim strSearchString 0x0001 (As Variant)
' Line #105:
' Line #106:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld _B_var_arrOtherEp 
' 	FnUBound 0x0000 
' 	For 
' Line #107:
' 	Ld objTTX 
' 	Ld i 
' 	ArgsLd _B_var_arrOtherEp 0x0001 
' 	LitVarSpecial (False)
' 	ArgsLd _B_var_buildEpocList 0x0003 
' 	Ld i 
' 	ArgsSt strSearchString 0x0001 
' Line #108:
' 	StartForVariable 
' 	Next 
' Line #109:
' 	EndIfBlock 
' Line #110:
' Line #111:
' 	LitDI2 0x0000 
' 	St i 
' Line #112:
' 	LitDI2 0x0000 
' 	St j 
' Line #113:
' Line #114:
' 	Dim 
' 	VarDefn iYAxisIndex (As Integer)
' Line #115:
' 	Dim 
' 	VarDefn otherEpocList (As Integer)
' Line #116:
' 	Dim 
' 	VarDefn _B_var_arrOtherEpocIndex (As Integer)
' Line #117:
' 	Ld _B_var_arrOtherEp 
' 	FnUBound 0x0000 
' 	LitDI2 0x0001 
' 	UMi 
' 	Ne 
' 	IfBlock 
' Line #118:
' 	OptionBase 
' 	Ld _B_var_arrOtherEp 
' 	FnUBound 0x0000 
' 	Redim _B_var_arrOtherEpocIndex 0x0001 (As Variant)
' Line #119:
' 	EndIfBlock 
' Line #120:
' Line #121:
' 	Dim 
' 	VarDefn varChanData (As Variant)
' Line #122:
' 	Dim 
' 	VarDefn IsEmpty (As Double)
' Line #123:
' Line #124:
' 	Dim 
' 	VarDefn yAxisSearchString (As String)
' Line #125:
' 	Dim 
' 	VarDefn otherAxisSearchString (As String)
' Line #126:
' 	Dim 
' 	VarDefn processSearch (As String)
' Line #127:
' 	Dim 
' 	VarDefn arrOtherEpFor (As String)
' Line #128:
' 	Ld _B_var_arrOtherEp 
' 	FnUBound 0x0000 
' 	LitDI2 0x0001 
' 	UMi 
' 	Ne 
' 	IfBlock 
' Line #129:
' 	OptionBase 
' 	Ld _B_var_arrOtherEp 
' 	FnUBound 0x0000 
' 	Redim processSearch 0x0001 (As Variant)
' Line #130:
' 	EndIfBlock 
' Line #131:
' Line #132:
' 	Dim 
' 	VarDefn 1 (As Integer)
' Line #133:
' 	LitDI2 0x0000 
' 	St 1 
' Line #134:
' Line #135:
' 	Ld _B_var_arrOtherEp 
' 	FnUBound 0x0000 
' 	LitDI2 0x0001 
' 	UMi 
' 	Ne 
' 	IfBlock 
' Line #136:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld vYAxisKeys 
' 	FnUBound 0x0000 
' 	For 
' Line #137:
' 	Ld yAxisEp 
' 	LitStr 0x0007 "Channel"
' 	Eq 
' 	IfBlock 
' Line #138:
' 	Ld i 
' 	ArgsLd vYAxisKeys 0x0001 
' 	St 1 
' Line #139:
' 	LitStr 0x0000 ""
' 	St yAxisSearchString 
' Line #140:
' 	ElseBlock 
' Line #141:
' 	Ld yAxisEp 
' 	LitStr 0x0003 " = "
' 	Concat 
' 	Ld i 
' 	ArgsLd vYAxisKeys 0x0001 
' 	Coerce (Str) 
' 	Concat 
' 	LitStr 0x0005 " and "
' 	Concat 
' 	St yAxisSearchString 
' Line #142:
' 	EndIfBlock 
' Line #143:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld buildOptionLists 
' 	FnUBound 0x0000 
' 	For 
' Line #144:
' 	Ld otherEp 
' 	LitStr 0x0007 "Channel"
' 	Eq 
' 	IfBlock 
' Line #145:
' 	Ld j 
' 	ArgsLd buildOptionLists 0x0001 
' 	St 1 
' Line #146:
' 	LitStr 0x0000 ""
' 	St otherAxisSearchString 
' Line #147:
' 	ElseBlock 
' Line #148:
' 	Ld otherEp 
' 	LitStr 0x0003 " = "
' 	Concat 
' 	Ld j 
' 	ArgsLd buildOptionLists 0x0001 
' 	Coerce (Str) 
' 	Concat 
' 	LitStr 0x0005 " and "
' 	Concat 
' 	St otherAxisSearchString 
' Line #149:
' 	EndIfBlock 
' Line #150:
' 	Ld objTTX 
' 	Ld _B_var_arrOtherEp 
' 	Ld strSearchString 
' 	LitDI2 0x0000 
' 	Ld yAxisSearchString 
' 	Ld otherAxisSearchString 
' 	Concat 
' 	Ld i 
' 	LitDI2 0x0001 
' 	Add 
' 	Ld j 
' 	LitDI2 0x0001 
' 	Add 
' 	Ld buildOptionLists 
' 	FnUBound 0x0000 
' 	LitDI2 0x0003 
' 	Add 
' 	Ld 1 
' 	LitStr 0x0000 ""
' 	Ld yCoun 
' 	Ld xCount 
' 	Ld lMaxHistHe 
' 	Ld lMaxHistHeigh 
' 	ArgsCall (Call) _B_var_processSearch 0x000E 
' Line #151:
' 	StartForVariable 
' 	Next 
' Line #152:
' 	StartForVariable 
' 	Next 
' Line #153:
' 	EndIfBlock 
' Line #154:
' Line #155:
' 	QuoteRem 0x0000 0x0051 "    Call writeAxes(theWorksheets, vXAxisKeys, vYAxisKeys, iColOffset, iRowOffset)"
' Line #156:
' Line #157:
' 	Ld objTTX 
' 	ArgsMemCall (Call) CloseTank 0x0000 
' Line #158:
' 	Ld objTTX 
' 	ArgsMemCall (Call) ReleaseServer 0x0000 
' Line #159:
' Line #160:
' 	Ld yCoun 
' 	LitStr 0x0002 "H1"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #161:
' 	Ld xCount 
' 	LitStr 0x0002 "H2"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #162:
' 	Ld lMaxHistHe 
' 	LitStr 0x0002 "H3"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #163:
' 	Ld lMaxHistHeigh 
' 	LitStr 0x0002 "H4"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #164:
' 	Ld _B_var_iColOffset 
' 	LitStr 0x0002 "H5"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #165:
' 	Ld _B_var_Const 
' 	LitStr 0x0002 "H6"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #166:
' Line #167:
' 	QuoteRem 0x0004 0x001B "If importIntoSigmaplot Then"
' Line #168:
' 	QuoteRem 0x0008 0x005D "Call transferToSigmaplot(xCount, yCount, zOffsetSize, iColOffset, iRowOffset, lMaxHistHeight)"
' Line #169:
' 	QuoteRem 0x0004 0x0006 "End If"
' Line #170:
' Line #171:
' 	EndSub 
' Line #172:
' Line #173:
' 	FuncDefn (Function buildWorksheetArray() As Variant)
' Line #174:
' 	Dim 
' 	OptionBase 
' 	LitDI2 0x001F 
' 	VarDefn theWorksheets
' Line #175:
' 	Dim 
' 	VarDefn strWsname (As String)
' Line #176:
' 	Dim 
' 	VarDefn intWSNum (As Long)
' Line #177:
' Line #178:
' 	Dim 
' 	VarDefn i (As Integer)
' Line #179:
' Line #180:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0001 
' 	Ld Worksheets 
' 	MemLd Count 
' 	For 
' Line #181:
' 	Ld i 
' 	Ld Worksheets 
' 	ArgsMemLd Item 0x0001 
' 	MemLd Name 
' 	St strWsname 
' Line #182:
' 	Ld strWsname 
' 	LitDI2 0x0004 
' 	ArgsLd Left 0x0002 
' 	LitStr 0x0004 "Site"
' 	Eq 
' 	IfBlock 
' Line #183:
' 	Ld strWsname 
' 	Ld strWsname 
' 	FnLen 
' 	LitDI2 0x0004 
' 	Sub 
' 	ArgsLd Right 0x0002 
' 	ArgsLd IsNumeric 0x0001 
' 	IfBlock 
' Line #184:
' 	Ld strWsname 
' 	Ld strWsname 
' 	FnLen 
' 	LitDI2 0x0004 
' 	Sub 
' 	ArgsLd Right 0x0002 
' 	Coerce (Int) 
' 	St intWSNum 
' Line #185:
' 	Ld intWSNum 
' 	LitDI2 0x0021 
' 	Lt 
' 	Ld intWSNum 
' 	LitDI2 0x0000 
' 	Gt 
' 	And 
' 	IfBlock 
' Line #186:
' 	SetStmt 
' 	Ld i 
' 	Ld Worksheets 
' 	ArgsMemLd Item 0x0001 
' 	Ld intWSNum 
' 	LitDI2 0x0001 
' 	Sub 
' 	ArgsSet theWorksheets 0x0001 
' Line #187:
' 	EndIfBlock 
' Line #188:
' 	EndIfBlock 
' Line #189:
' 	EndIfBlock 
' Line #190:
' 	StartForVariable 
' 	Next 
' Line #191:
' Line #192:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	LitDI2 0x001F 
' 	For 
' Line #193:
' 	Ld i 
' 	ArgsLd theWorksheets 0x0001 
' 	ArgsLd Sheet70 0x0001 
' 	IfBlock 
' Line #194:
' 	Ld i 
' 	LitDI2 0x0000 
' 	Gt 
' 	IfBlock 
' Line #195:
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
' Line #196:
' 	ElseBlock 
' Line #197:
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
' Line #198:
' 	EndIfBlock 
' Line #199:
' 	LitStr 0x0004 "Site"
' 	Ld i 
' 	LitDI2 0x0001 
' 	Add 
' 	Coerce (Str) 
' 	Concat 
' 	Ld i 
' 	ArgsLd theWorksheets 0x0001 
' 	MemSt Name 
' Line #200:
' 	EndIfBlock 
' Line #201:
' 	StartForVariable 
' 	Next 
' Line #202:
' 	Ld theWorksheets 
' 	St buildWorksheetArray 
' Line #203:
' 	EndFunc 
' Line #204:
' Line #205:
' 	FuncDefn (Sub SubwriteAxes(rowLabels As Variant, deleteWorksheets As Variant, _B_var_iColOffset, _B_var_Const, xOffes))
' Line #206:
' 	Dim 
' 	VarDefn j (As Long)
' Line #207:
' Line #208:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld deleteWorksheets 
' 	FnUBound 0x0000 
' 	For 
' Line #209:
' 	Ld j 
' 	ArgsLd deleteWorksheets 0x0001 
' 	Ld _B_var_Const 
' 	Ld j 
' 	Add 
' 	LitDI2 0x0002 
' 	Add 
' 	Ld xOffes 
' 	Add 
' 	Ld _B_var_iColOffset 
' 	LitDI2 0x0001 
' 	Add 
' 	LitStr 0x0006 "Output"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Cells 0x0002 
' 	MemSt Value 
' Line #210:
' 	StartForVariable 
' 	Next 
' Line #211:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld rowLabels 
' 	FnUBound 0x0000 
' 	For 
' Line #212:
' 	Ld j 
' 	ArgsLd rowLabels 0x0001 
' 	Ld _B_var_Const 
' 	Ld xOffes 
' 	Add 
' 	LitDI2 0x0001 
' 	Add 
' 	Ld j 
' 	LitDI2 0x0002 
' 	Add 
' 	LitStr 0x0006 "Output"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Cells 0x0002 
' 	MemSt Value 
' Line #213:
' 	StartForVariable 
' 	Next 
' Line #214:
' Line #215:
' 	EndSub 
' Line #216:
' Line #217:
' 	FuncDefn (Sub Delete())
' Line #218:
' 	Dim 
' 	VarDefn strWsname (As String)
' Line #219:
' 	Dim 
' 	VarDefn intWSNum (As Long)
' Line #220:
' Line #221:
' 	Dim 
' 	VarDefn i (As Integer)
' Line #222:
' Line #223:
' 	Ld Worksheets 
' 	MemLd Count 
' 	St i 
' Line #224:
' Line #225:
' 	Do 
' Line #226:
' 	Ld i 
' 	LitDI2 0x0000 
' 	Eq 
' 	IfBlock 
' Line #227:
' 	ExitDo 
' Line #228:
' 	EndIfBlock 
' Line #229:
' 	Ld i 
' 	Ld Worksheets 
' 	ArgsMemLd Item 0x0001 
' 	MemLd Name 
' 	St strWsname 
' Line #230:
' 	Ld strWsname 
' 	LitDI2 0x0004 
' 	ArgsLd Left 0x0002 
' 	LitStr 0x0004 "Site"
' 	Eq 
' 	IfBlock 
' Line #231:
' 	Ld strWsname 
' 	Ld strWsname 
' 	FnLen 
' 	LitDI2 0x0004 
' 	Sub 
' 	ArgsLd Right 0x0002 
' 	ArgsLd IsNumeric 0x0001 
' 	IfBlock 
' Line #232:
' 	Ld strWsname 
' 	Ld strWsname 
' 	FnLen 
' 	LitDI2 0x0004 
' 	Sub 
' 	ArgsLd Right 0x0002 
' 	Coerce (Int) 
' 	St intWSNum 
' Line #233:
' 	Ld intWSNum 
' 	LitDI2 0x0021 
' 	Lt 
' 	Ld intWSNum 
' 	LitDI2 0x0000 
' 	Gt 
' 	And 
' 	IfBlock 
' Line #234:
' 	Ld i 
' 	Ld Worksheets 
' 	ArgsMemLd Item 0x0001 
' 	ArgsMemCall UserForm1 0x0000 
' Line #235:
' 	EndIfBlock 
' Line #236:
' 	EndIfBlock 
' Line #237:
' 	EndIfBlock 
' Line #238:
' 	Ld i 
' 	LitDI2 0x0001 
' 	Sub 
' 	St i 
' Line #239:
' 	Loop 
' Line #240:
' 	EndSub 
' Line #241:
' Line #242:
' 	FuncDefn (Sub ACTIVESPWLib())
' Line #243:
' Line #244:
' 	Dim 
' 	VarDefn yCoun (As Long)
' Line #245:
' 	Dim 
' 	VarDefn xCount (As Long)
' Line #246:
' 	Dim 
' 	VarDefn lMaxHistHe (As Long)
' Line #247:
' 	Dim 
' 	VarDefn lMaxHistHeigh (As Long)
' Line #248:
' 	Dim 
' 	VarDefn _B_var_iColOffset (As Integer)
' Line #249:
' 	Dim 
' 	VarDefn _B_var_Const (As Integer)
' Line #250:
' Line #251:
' 	LitStr 0x0002 "H1"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	St yCoun 
' Line #252:
' 	LitStr 0x0002 "H2"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	St xCount 
' Line #253:
' 	LitStr 0x0002 "H3"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	St lMaxHistHe 
' Line #254:
' 	LitStr 0x0002 "H4"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	St lMaxHistHeigh 
' Line #255:
' 	LitStr 0x0002 "H5"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	St _B_var_iColOffset 
' Line #256:
' 	LitStr 0x0002 "H6"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	St _B_var_Const 
' Line #257:
' Line #258:
' 	Dim (Const) 
' 	LitHI2 0x0406 
' 	VarDefn SAA_TOVAL
' Line #259:
' 	Dim (Const) 
' 	LitHI2 0x0407 
' 	VarDefn GraphPages
' Line #260:
' 	Dim (Const) 
' 	LitHI2 0x0301 
' 	VarDefn SLA_SELECTDIM
' Line #261:
' 	Dim (Const) 
' 	LitDI2 0x0401 
' 	VarDefn SEA_COLORCOL
' Line #262:
' 	Dim (Const) 
' 	LitDI2 0x0308 
' 	VarDefn SAA_OPTIONS
' Line #263:
' 	Dim (Const) 
' 	LitDI2 0x0403 
' 	VarDefn _B_var_GPM_SETPLOTATTR
' Line #264:
' 	Dim (Const) 
' 	LitDI2 0x0408 
' 	VarDefn SAA_FROMVAL
' Line #265:
' 	Dim (Const) 
' 	LitDI2 0x0615 
' 	VarDefn GPM_SETAXISATTRSTRING
' Line #266:
' 	Dim (Const) 
' 	LitDI2 0x0613 
' 	VarDefn SLA_CONTOURFILLTYPE
' Line #267:
' 	Dim (Const) 
' 	LitDI2 0x0358 
' 	VarDefn SAA_SELECTLINE
' Line #268:
' 	Dim (Const) 
' 	LitDI2 0x040A 
' 	VarDefn SEA_THICKNESS
' Line #269:
' 	Dim (Const) 
' 	LitDI2 0x0601 
' 	VarDefn SEA_COLOR
' Line #270:
' 	Dim (Const) 
' 	LitDI2 0x0606 
' 	VarDefn _B_var_SEA_THICKNESS
' Line #271:
' 	Dim (Const) 
' 	LitDI2 0x0410 
' 	VarDefn _B_var_SAA_SUB1OPTIONS
' Line #272:
' Line #273:
' 	Dim 
' 	VarDefn Module2 (As Object)
' Line #274:
' 	SetStmt 
' 	LitStr 0x0017 "SigmaPlot.Application.1"
' 	ArgsLd CreateObject 0x0001 
' 	Set Module2 
' Line #275:
' 	LitVarSpecial (True)
' 	Ld Module2 
' 	MemSt Application 
' Line #276:
' 	Ld Module2 
' 	MemLd Notebooks 
' 	MemLd buildTuningCurvesIntoSigmaplot 
' 	ArgsMemCall (Call) Add 0x0000 
' Line #277:
' Line #278:
' 	Dim 
' 	VarDefn i (As Integer)
' Line #279:
' 	Dim 
' 	VarDefn j (As Long)
' Line #280:
' 	Dim 
' 	VarDefn k (As Long)
' Line #281:
' Line #282:
' 	Dim 
' 	VarDefn yPos (As Long)
' Line #283:
' 	Dim 
' 	VarDefn Whi0l (As Long)
' Line #284:
' Line #285:
' 	Dim 
' 	VarDefn SPApplication (As Object)
' Line #286:
' 	Dim 
' 	VarDefn spDT (As Object)
' Line #287:
' 	Dim 
' 	VarDefn DataTable (As Object)
' Line #288:
' 	Dim 
' 	VarDefn objSPWizard (As Object)
' Line #289:
' Line #290:
' 	Ld _B_var_iColOffset 
' 	LitDI2 0x0001 
' 	Add 
' 	St yPos 
' Line #291:
' 	Ld _B_var_Const 
' 	St Whi0l 
' Line #292:
' Line #293:
' 	Do 
' Line #294:
' 	Ld Whi0l 
' 	Ld yPos 
' 	LitStr 0x0006 "Output"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Cells 0x0002 
' 	MemLd Value 
' 	LitStr 0x0000 ""
' 	Ne 
' 	IfBlock 
' Line #295:
' 	Ld Whi0l 
' 	Ld yPos 
' 	LitStr 0x0006 "Output"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Cells 0x0002 
' 	MemLd Value 
' 	Ld bXAxisLog 
' 	ArgsMemLd Exists 0x0001 
' 	IfBlock 
' Line #296:
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
' Line #297:
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
' Line #298:
' 	Ld Whi0l 
' 	Ld yPos 
' 	LitStr 0x0006 "Output"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Cells 0x0002 
' 	MemLd Value 
' 	Ld spDT 
' 	MemSt Name 
' Line #299:
' 	SetStmt 
' 	Ld spDT 
' 	MemLd Cell 
' 	Set DataTable 
' Line #300:
' Line #301:
' 	Ld Whi0l 
' 	St Whi0l 
' Line #302:
' Line #303:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld yCoun 
' 	LitDI2 0x0001 
' 	Sub 
' 	Paren 
' 	For 
' Line #304:
' 	Ld Whi0l 
' 	LitDI2 0x0001 
' 	Add 
' 	Ld yPos 
' 	Ld j 
' 	Add 
' 	LitDI2 0x0001 
' 	Add 
' 	LitStr 0x0006 "Output"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Cells 0x0002 
' 	MemLd Value 
' 	LitDI2 0x0000 
' 	Ld j 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #305:
' 	StartForVariable 
' 	Next 
' Line #306:
' Line #307:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld xCount 
' 	LitDI2 0x0001 
' 	Sub 
' 	Paren 
' 	For 
' Line #308:
' 	Ld Whi0l 
' 	Ld j 
' 	Add 
' 	LitDI2 0x0002 
' 	Add 
' 	Ld yPos 
' 	LitStr 0x0006 "Output"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Cells 0x0002 
' 	MemLd Value 
' 	LitDI2 0x0001 
' 	Ld j 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #309:
' 	StartForVariable 
' 	Next 
' Line #310:
' Line #311:
' Line #312:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld yCoun 
' 	LitDI2 0x0001 
' 	Sub 
' 	Paren 
' 	For 
' Line #313:
' 	StartForVariable 
' 	Ld k 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld xCount 
' 	LitDI2 0x0001 
' 	Sub 
' 	Paren 
' 	For 
' Line #314:
' 	Ld Whi0l 
' 	Ld k 
' 	Add 
' 	LitDI2 0x0002 
' 	Add 
' 	Ld yPos 
' 	Ld j 
' 	Add 
' 	LitDI2 0x0001 
' 	Add 
' 	LitStr 0x0006 "Output"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Cells 0x0002 
' 	MemLd Value 
' 	LitDI2 0x0003 
' 	Ld k 
' 	Add 
' 	Ld j 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #315:
' 	StartForVariable 
' 	Next 
' Line #316:
' 	StartForVariable 
' 	Next 
' Line #317:
' Line #318:
' 	LitStr 0x0011 "@rgb(255,255,255)"
' 	LitDI2 0x0002 
' 	LitDI2 0x0000 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #319:
' 	LitStr 0x000D "@rgb(0,0,255)"
' 	LitDI2 0x0002 
' 	LitDI2 0x0001 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #320:
' 	LitStr 0x000F "@rgb(0,255,255)"
' 	LitDI2 0x0002 
' 	LitDI2 0x0002 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #321:
' 	LitStr 0x000D "@rgb(0,255,0)"
' 	LitDI2 0x0002 
' 	LitDI2 0x0003 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #322:
' 	LitStr 0x000F "@rgb(255,255,0)"
' 	LitDI2 0x0002 
' 	LitDI2 0x0004 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #323:
' 	LitStr 0x000D "@rgb(255,0,0)"
' 	LitDI2 0x0002 
' 	LitDI2 0x0005 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #324:
' Line #325:
' 	QuoteRem 0x0010 0x001E "Call spNB.NotebookItems.Add(2)"
' Line #326:
' 	QuoteRem 0x0010 0x0042 "Set spGRPH = spNB.NotebookItems.Item(spNB.NotebookItems.Count - 1)"
' Line #327:
' Line #328:
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd SPWNotebookComponentType 
' 	ArgsMemCall (Call) Add 0x0001 
' Line #329:
' 	Dim 
' 	OptionBase 
' 	LitDI2 0x0002 
' 	OptionBase 
' 	LitDI2 0x0003 
' 	VarDefn PlotColumnCountArray
' Line #330:
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #331:
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	LitDI2 0x0000 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #332:
' 	LitDI4 0x47FF 0x01E8 
' 	LitDI2 0x0002 
' 	LitDI2 0x0000 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #333:
' 	LitDI2 0x0001 
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #334:
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	LitDI2 0x0001 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #335:
' 	LitDI4 0x47FF 0x01E8 
' 	LitDI2 0x0002 
' 	LitDI2 0x0001 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #336:
' 	LitDI2 0x0003 
' 	LitDI2 0x0000 
' 	LitDI2 0x0002 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #337:
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	LitDI2 0x0002 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #338:
' 	LitDI4 0x47FF 0x01E8 
' 	LitDI2 0x0002 
' 	LitDI2 0x0002 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #339:
' 	LitDI2 0x0003 
' 	Ld xCount 
' 	LitDI2 0x0001 
' 	Sub 
' 	Paren 
' 	Add 
' 	LitDI2 0x0000 
' 	LitDI2 0x0003 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #340:
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	LitDI2 0x0003 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #341:
' 	LitDI4 0x47FF 0x01E8 
' 	LitDI2 0x0002 
' 	LitDI2 0x0003 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #342:
' Line #343:
' 	Dim 
' 	VarDefn CurrentPageItem
' Line #344:
' 	OptionBase 
' 	LitDI2 0x0000 
' 	Redim CurrentPageItem 0x0001 (As Variant)
' Line #345:
' Line #346:
' 	LitDI2 0x0004 
' 	LitDI2 0x0000 
' 	ArgsSt CurrentPageItem 0x0001 
' Line #347:
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
' Line #348:
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemLd Graphs 0x0001 
' 	ArgsMemLd SelectObject 0x0001 
' 	ArgsMemCall (Call) testSigmaPlot 0x0000 
' Line #349:
' Line #350:
' 	LitStr 0x0006 "Site y"
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemLd Graphs 0x0001 
' 	ArgsMemLd SelectObject 0x0001 
' 	MemSt Name 
' Line #351:
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
' Line #352:
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
' Line #353:
' Line #354:
' 	Ld SLA_SELECTDIM 
' 	Ld SAA_OPTIONS 
' 	LitDI2 0x0003 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #355:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_GPM_SETPLOTATTR 
' 	LitDI2 0x000A 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #356:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_GPM_SETPLOTATTR 
' 	LitDI4 0x0001 0x0310 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #357:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_GPM_SETPLOTATTR 
' 	LitDI4 0x0402 0x00C0 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #358:
' 	Ld SAA_FROMVAL 
' 	Ld SAA_TOVAL 
' 	LitStr 0x0001 "0"
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #359:
' 	Ld SAA_FROMVAL 
' 	Ld GraphPages 
' 	Ld lMaxHistHeigh 
' 	Coerce (Str) 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #360:
' Line #361:
' 	Ld SLA_SELECTDIM 
' 	Ld SAA_OPTIONS 
' 	LitDI2 0x0003 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #362:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #363:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_COLOR 
' 	LitDI2 0x000A 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #364:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #365:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SEA_THICKNESS 
' 	LitHI4 0xFFFF 0x00FF 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #366:
' 	Ld SEA_COLORCOL 
' 	Ld GPM_SETAXISATTRSTRING 
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #367:
' 	Ld SEA_COLORCOL 
' 	Ld SLA_CONTOURFILLTYPE 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #368:
' 	Ld SLA_SELECTDIM 
' 	Ld SAA_SELECTLINE 
' 	LitDI2 0x0001 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #369:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0005 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #370:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SEA_THICKNESS 
' 	LitHI4 0xFFFF 0x00FF 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #371:
' 	Ld SEA_COLORCOL 
' 	Ld GPM_SETAXISATTRSTRING 
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #372:
' 	Ld SEA_COLORCOL 
' 	Ld SLA_CONTOURFILLTYPE 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #373:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #374:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SAA_SUB1OPTIONS 
' 	LitDI2 0x0512 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #375:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SAA_SUB1OPTIONS 
' 	LitDI2 0x0F31 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #376:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0001 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #377:
' Line #378:
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd SPWNotebookComponentType 
' 	ArgsMemCall (Call) Add 0x0001 
' Line #379:
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #380:
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	LitDI2 0x0000 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #381:
' 	LitDI4 0x47FF 0x01E8 
' 	LitDI2 0x0002 
' 	LitDI2 0x0000 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #382:
' 	LitDI2 0x0001 
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #383:
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	LitDI2 0x0001 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #384:
' 	LitDI4 0x47FF 0x01E8 
' 	LitDI2 0x0002 
' 	LitDI2 0x0001 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #385:
' 	LitDI2 0x0003 
' 	LitDI2 0x0000 
' 	LitDI2 0x0002 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #386:
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	LitDI2 0x0002 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #387:
' 	LitDI4 0x47FF 0x01E8 
' 	LitDI2 0x0002 
' 	LitDI2 0x0002 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #388:
' 	LitDI2 0x0003 
' 	Ld xCount 
' 	LitDI2 0x0001 
' 	Sub 
' 	Paren 
' 	Add 
' 	LitDI2 0x0000 
' 	LitDI2 0x0003 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #389:
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	LitDI2 0x0003 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #390:
' 	LitDI4 0x47FF 0x01E8 
' 	LitDI2 0x0002 
' 	LitDI2 0x0003 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #391:
' Line #392:
' 	OptionBase 
' 	LitDI2 0x0000 
' 	Redim CurrentPageItem 0x0001 (As Variant)
' Line #393:
' Line #394:
' 	LitDI2 0x0004 
' 	LitDI2 0x0000 
' 	ArgsSt CurrentPageItem 0x0001 
' Line #395:
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
' Line #396:
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemLd Graphs 0x0001 
' 	ArgsMemLd SelectObject 0x0001 
' 	ArgsMemCall (Call) testSigmaPlot 0x0000 
' Line #397:
' Line #398:
' 	LitStr 0x0006 "Site y"
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemLd Graphs 0x0001 
' 	ArgsMemLd SelectObject 0x0001 
' 	MemSt Name 
' Line #399:
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
' Line #400:
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
' Line #401:
' Line #402:
' 	Ld SLA_SELECTDIM 
' 	Ld SAA_OPTIONS 
' 	LitDI2 0x0003 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #403:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #404:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_COLOR 
' 	LitDI2 0x000A 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #405:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #406:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SEA_THICKNESS 
' 	LitHI4 0xFFFF 0x00FF 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #407:
' 	Ld SEA_COLORCOL 
' 	Ld GPM_SETAXISATTRSTRING 
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #408:
' 	Ld SEA_COLORCOL 
' 	Ld SLA_CONTOURFILLTYPE 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #409:
' 	Ld SLA_SELECTDIM 
' 	Ld SAA_SELECTLINE 
' 	LitDI2 0x0001 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #410:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0005 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #411:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SEA_THICKNESS 
' 	LitHI4 0xFFFF 0x00FF 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #412:
' 	Ld SEA_COLORCOL 
' 	Ld GPM_SETAXISATTRSTRING 
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #413:
' 	Ld SEA_COLORCOL 
' 	Ld SLA_CONTOURFILLTYPE 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #414:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #415:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SAA_SUB1OPTIONS 
' 	LitDI2 0x0512 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #416:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SAA_SUB1OPTIONS 
' 	LitDI2 0x0F31 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #417:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0001 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #418:
' Line #419:
' 	LitDI2 0x0001 
' 	Ld SPApplication 
' 	MemLd SPWNotebookComponentType 
' 	ArgsMemCall (Call) Add 0x0001 
' Line #420:
' 	EndIfBlock 
' Line #421:
' 	Ld Whi0l 
' 	Ld lMaxHistHe 
' 	Add 
' 	St Whi0l 
' Line #422:
' 	ElseBlock 
' Line #423:
' 	ExitDo 
' Line #424:
' 	EndIfBlock 
' Line #425:
' 	Loop 
' Line #426:
' 	EndSub 
' Line #427:
' Line #428:
' 	FuncDefn (Sub _B_var_lMaxHistHeight())
' Line #429:
' 	Dim (Const) 
' 	LitHI2 0x0406 
' 	VarDefn SAA_TOVAL
' Line #430:
' 	Dim (Const) 
' 	LitHI2 0x0407 
' 	VarDefn GraphPages
' Line #431:
' 	Dim (Const) 
' 	LitHI2 0x0301 
' 	VarDefn SLA_SELECTDIM
' Line #432:
' 	Dim (Const) 
' 	LitDI2 0x0401 
' 	VarDefn SEA_COLORCOL
' Line #433:
' 	Dim (Const) 
' 	LitDI2 0x0308 
' 	VarDefn SAA_OPTIONS
' Line #434:
' 	Dim (Const) 
' 	LitDI2 0x0403 
' 	VarDefn _B_var_GPM_SETPLOTATTR
' Line #435:
' 	Dim (Const) 
' 	LitDI2 0x0408 
' 	VarDefn SAA_FROMVAL
' Line #436:
' 	Dim (Const) 
' 	LitDI2 0x0615 
' 	VarDefn GPM_SETAXISATTRSTRING
' Line #437:
' 	Dim (Const) 
' 	LitDI2 0x0613 
' 	VarDefn SLA_CONTOURFILLTYPE
' Line #438:
' 	Dim (Const) 
' 	LitDI2 0x0358 
' 	VarDefn SAA_SELECTLINE
' Line #439:
' 	Dim (Const) 
' 	LitDI2 0x040A 
' 	VarDefn SEA_THICKNESS
' Line #440:
' 	Dim (Const) 
' 	LitDI2 0x0601 
' 	VarDefn SEA_COLOR
' Line #441:
' 	Dim (Const) 
' 	LitDI2 0x0606 
' 	VarDefn _B_var_SEA_THICKNESS
' Line #442:
' 	Dim (Const) 
' 	LitDI2 0x0410 
' 	VarDefn _B_var_SAA_SUB1OPTIONS
' Line #443:
' Line #444:
' 	Dim 
' 	VarDefn Module2 (As Object)
' Line #445:
' 	SetStmt 
' 	LitStr 0x0017 "SigmaPlot.Application.1"
' 	ArgsLd CreateObject 0x0001 
' 	Set Module2 
' Line #446:
' 	LitVarSpecial (True)
' 	Ld Module2 
' 	MemSt Application 
' Line #447:
' 	LitStr 0x0006 "Site y"
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemLd Graphs 0x0001 
' 	ArgsMemLd SelectObject 0x0001 
' 	MemSt Name 
' Line #448:
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
' Line #449:
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
' Line #450:
' Line #451:
' 	Ld SLA_SELECTDIM 
' 	Ld SAA_OPTIONS 
' 	LitDI2 0x0003 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #452:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_GPM_SETPLOTATTR 
' 	LitDI2 0x000A 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #453:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_GPM_SETPLOTATTR 
' 	LitDI4 0x0001 0x0310 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #454:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_GPM_SETPLOTATTR 
' 	LitDI4 0x0402 0x00C0 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #455:
' 	Ld SAA_FROMVAL 
' 	Ld SAA_TOVAL 
' 	LitStr 0x0001 "0"
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #456:
' 	Ld SAA_FROMVAL 
' 	Ld GraphPages 
' 	LitStr 0x0003 "150"
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #457:
' Line #458:
' 	Ld SLA_SELECTDIM 
' 	Ld SAA_OPTIONS 
' 	LitDI2 0x0003 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #459:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #460:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_COLOR 
' 	LitDI2 0x000A 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #461:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #462:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SEA_THICKNESS 
' 	LitHI4 0xFFFF 0x00FF 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #463:
' 	Ld SEA_COLORCOL 
' 	Ld GPM_SETAXISATTRSTRING 
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #464:
' 	Ld SEA_COLORCOL 
' 	Ld SLA_CONTOURFILLTYPE 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #465:
' 	Ld SLA_SELECTDIM 
' 	Ld SAA_SELECTLINE 
' 	LitDI2 0x0001 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #466:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0005 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #467:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SEA_THICKNESS 
' 	LitHI4 0xFFFF 0x00FF 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #468:
' 	Ld SEA_COLORCOL 
' 	Ld GPM_SETAXISATTRSTRING 
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #469:
' 	Ld SEA_COLORCOL 
' 	Ld SLA_CONTOURFILLTYPE 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #470:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #471:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SAA_SUB1OPTIONS 
' 	LitDI2 0x0512 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #472:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SAA_SUB1OPTIONS 
' 	LitDI2 0x0F31 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #473:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0001 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #474:
' Line #475:
' 	EndSub 
' Line #476:
' Line #477:
' 	FuncDefn (Function _B_var_buildEpocList(objTTX, iXAxisIndexAs, returnArr))
' Line #478:
' 	QuoteRem 0x0004 0x0030 "build list of epocs for the given axis epoc name"
' Line #479:
' Line #480:
' 	Dim 
' 	VarDefn AxisEp (As Dictionary)
' Line #481:
' 	SetStmt 
' 	New id_FFFF
' 	Set AxisEp 
' Line #482:
' Line #483:
' 	Dim 
' 	VarDefn dblStartTime (As Double)
' Line #484:
' 	Dim 
' 	VarDefn varReturn (As Variant)
' Line #485:
' Line #486:
' 	Dim 
' 	VarDefn i (As Integer)
' Line #487:
' 	Dim 
' 	VarDefn j (As Integer)
' Line #488:
' Line #489:
' 	Ld iXAxisIndexAs 
' 	LitStr 0x0007 "Channel"
' 	Eq 
' 	IfBlock 
' Line #490:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0001 
' 	LitDI2 0x0020 
' 	For 
' Line #491:
' 	Ld i 
' 	LitDI2 0x0000 
' 	Ld AxisEp 
' 	ArgsMemCall (Call) Add 0x0002 
' Line #492:
' 	StartForVariable 
' 	Next 
' Line #493:
' 	ElseBlock 
' Line #494:
' 	Do 
' Line #495:
' 	LitDI2 0x01F4 
' 	Ld iXAxisIndexAs 
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	Ld dblStartTime 
' 	LitR8 0x0000 0x0000 0x0000 0x0000 
' 	LitStr 0x0003 "ALL"
' 	Ld objTTX 
' 	ArgsMemLd ReadEventsV 0x0007 
' 	St i 
' Line #496:
' 	Ld i 
' 	LitDI2 0x0000 
' 	Eq 
' 	IfBlock 
' Line #497:
' 	ExitDo 
' Line #498:
' 	EndIfBlock 
' Line #499:
' Line #500:
' 	LitDI2 0x0000 
' 	Ld i 
' 	LitDI2 0x0000 
' 	Ld objTTX 
' 	ArgsMemLd ParseEvInfoV 0x0003 
' 	St varReturn 
' Line #501:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld i 
' 	LitDI2 0x0001 
' 	Sub 
' 	Paren 
' 	For 
' Line #502:
' 	LitDI2 0x0006 
' 	Ld j 
' 	ArgsLd varReturn 0x0002 
' 	Ld AxisEp 
' 	ArgsMemLd Exists 0x0001 
' 	Not 
' 	IfBlock 
' Line #503:
' 	LitDI2 0x0006 
' 	Ld j 
' 	ArgsLd varReturn 0x0002 
' 	LitStr 0x0000 ""
' 	Ld AxisEp 
' 	ArgsMemCall (Call) Add 0x0002 
' Line #504:
' 	EndIfBlock 
' Line #505:
' 	LitDI2 0x0005 
' 	Ld j 
' 	ArgsLd varReturn 0x0002 
' 	LitDI2 0x0001 
' 	LitDI4 0x86A0 0x0001 
' 	Div 
' 	Paren 
' 	Add 
' 	St dblStartTime 
' Line #506:
' 	StartForVariable 
' 	Next 
' Line #507:
' Line #508:
' 	Ld i 
' 	LitDI2 0x01F4 
' 	Lt 
' 	IfBlock 
' Line #509:
' 	ExitDo 
' Line #510:
' 	EndIfBlock 
' Line #511:
' 	Loop 
' Line #512:
' 	EndIfBlock 
' Line #513:
' Line #514:
' Line #515:
' Line #516:
' 	Ld returnArr 
' 	IfBlock 
' Line #517:
' 	Dim 
' 	VarDefn _B_var_returnArr
' Line #518:
' 	Dim 
' 	VarDefn zOffsets (As Variant)
' Line #519:
' 	Ld AxisEp 
' 	MemLd Keys 
' 	St zOffsets 
' Line #520:
' 	OptionBase 
' 	Ld zOffsets 
' 	FnUBound 0x0000 
' 	Redim _B_var_returnArr 0x0001 (As Variant)
' Line #521:
' Line #522:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld zOffsets 
' 	FnUBound 0x0000 
' 	For 
' Line #523:
' 	Ld zOffsets 
' 	FnUBound 0x0000 
' 	Ld i 
' 	Sub 
' 	ArgsLd zOffsets 0x0001 
' 	Ld i 
' 	ArgsSt _B_var_returnArr 0x0001 
' Line #524:
' 	StartForVariable 
' 	Next 
' Line #525:
' 	Ld _B_var_returnArr 
' 	St _B_var_buildEpocList 
' Line #526:
' 	ElseBlock 
' Line #527:
' 	Ld AxisEp 
' 	MemLd Keys 
' 	St _B_var_buildEpocList 
' Line #528:
' 	EndIfBlock 
' Line #529:
' Line #530:
' 	EndFunc 
' Line #531:
' Line #532:
' Line #533:
' 	FuncDefn (Function _B_var_processSearch(ByRef objTTX, ByRef _B_var_arrOtherEp, ByRef strSearchString, iOtherEpocIndex, xAxisSearchString As String, yOffset, zOffset, xOffes, 1, Le, ByRef yCoun, ByRef xCount, ByRef lMaxHistHe, ByRef lMaxHistHeigh))
' Line #534:
' 	Dim 
' 	VarDefn i (As Integer)
' Line #535:
' 	Dim 
' 	VarDefn j (As Integer)
' Line #536:
' 	Dim 
' 	VarDefn _B_var_objTTX (As String)
' Line #537:
' 	Dim 
' 	VarDefn strHeading (As String)
' Line #538:
' 	Dim 
' 	VarDefn Label1 (As String)
' Line #539:
' Line #540:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld iOtherEpocIndex 
' 	ArgsLd strSearchString 0x0001 
' 	FnUBound 0x0000 
' 	For 
' Line #541:
' 	Ld iOtherEpocIndex 
' 	ArgsLd _B_var_arrOtherEp 0x0001 
' 	LitStr 0x0007 "Channel"
' 	Ne 
' 	IfBlock 
' Line #542:
' 	QuoteRem 0x000C 0x0014 "add to search string"
' Line #543:
' 	Ld xAxisSearchString 
' 	Ld iOtherEpocIndex 
' 	ArgsLd _B_var_arrOtherEp 0x0001 
' 	Concat 
' 	LitStr 0x0003 " = "
' 	Concat 
' 	Ld i 
' 	Ld iOtherEpocIndex 
' 	ArgsLd strSearchString 0x0001 
' 	IndexLd 0x0001 
' 	Coerce (Str) 
' 	Concat 
' 	LitStr 0x0005 " and "
' 	Concat 
' 	St _B_var_objTTX 
' Line #544:
' 	Ld Le 
' 	Ld iOtherEpocIndex 
' 	ArgsLd _B_var_arrOtherEp 0x0001 
' 	Concat 
' 	LitStr 0x0003 " = "
' 	Concat 
' 	Ld i 
' 	Ld iOtherEpocIndex 
' 	ArgsLd strSearchString 0x0001 
' 	IndexLd 0x0001 
' 	Coerce (Str) 
' 	Concat 
' 	LitStr 0x0002 ", "
' 	Concat 
' 	St Label1 
' Line #545:
' 	ElseBlock 
' Line #546:
' 	Ld xAxisSearchString 
' 	St _B_var_objTTX 
' Line #547:
' 	Ld Le 
' 	LitStr 0x000A "Channel = "
' 	Concat 
' 	Ld i 
' 	Ld iOtherEpocIndex 
' 	ArgsLd strSearchString 0x0001 
' 	IndexLd 0x0001 
' 	Coerce (Str) 
' 	Concat 
' 	LitStr 0x0002 ", "
' 	Concat 
' 	St Label1 
' Line #548:
' 	Ld i 
' 	Ld iOtherEpocIndex 
' 	ArgsLd strSearchString 0x0001 
' 	IndexLd 0x0001 
' 	St 1 
' Line #549:
' 	EndIfBlock 
' Line #550:
' 	Ld iOtherEpocIndex 
' 	Ld _B_var_arrOtherEp 
' 	FnUBound 0x0000 
' 	Lt 
' 	IfBlock 
' Line #551:
' 	QuoteRem 0x000C 0x002F "there are still more epocs to add to the search"
' Line #552:
' 	Ld objTTX 
' 	Ld _B_var_arrOtherEp 
' 	Ld strSearchString 
' 	Ld iOtherEpocIndex 
' 	LitDI2 0x0001 
' 	Add 
' 	Ld _B_var_objTTX 
' 	Ld yOffset 
' 	Ld zOffset 
' 	Ld xOffes 
' 	Ld iOtherEpocIndex 
' 	ArgsLd strSearchString 0x0001 
' 	FnUBound 0x0000 
' 	Mul 
' 	Paren 
' 	Ld i 
' 	Add 
' 	Ld 1 
' 	Ld Label1 
' 	Ld yCoun 
' 	Ld xCount 
' 	Ld lMaxHistHe 
' 	Ld lMaxHistHeigh 
' 	ArgsCall (Call) _B_var_processSearch 0x000E 
' Line #553:
' 	ElseBlock 
' Line #554:
' 	QuoteRem 0x000C 0x004B "we have reached the end of the list of epocs - can actually do a search now"
' Line #555:
' 	Ld _B_var_objTTX 
' 	LitDI2 0x0005 
' 	ArgsLd Right 0x0002 
' 	LitStr 0x0005 " and "
' 	Eq 
' 	IfBlock 
' 	QuoteRem 0x003D 0x0045 "this should always be the case - should be a trailing 'and' to remove"
' Line #556:
' 	Ld _B_var_objTTX 
' 	Ld _B_var_objTTX 
' 	FnLen 
' 	LitDI2 0x0005 
' 	Sub 
' 	ArgsLd Left 0x0002 
' 	St strHeading 
' Line #557:
' 	ElseBlock 
' Line #558:
' 	Ld _B_var_objTTX 
' 	St strHeading 
' Line #559:
' 	EndIfBlock 
' Line #560:
' 	Ld strHeading 
' 	Ld objTTX 
' 	ArgsMemCall (Call) SetFilterWithDescEx 0x0001 
' Line #561:
' Line #562:
' 	Ld yOffset 
' 	LitDI2 0x0001 
' 	Eq 
' 	Ld zOffset 
' 	LitDI2 0x0001 
' 	Eq 
' 	And 
' 	IfBlock 
' Line #563:
' 	Ld Label1 
' 	Ld Label1 
' 	FnLen 
' 	LitDI2 0x0002 
' 	Sub 
' 	ArgsLd Left 0x0002 
' 	Ld _B_var_Const 
' 	Ld i 
' 	Ld xOffes 
' 	Mul 
' 	Paren 
' 	Add 
' 	Ld _B_var_iColOffset 
' 	LitDI2 0x0001 
' 	Add 
' 	LitStr 0x0006 "Output"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Cells 0x0002 
' 	MemSt Value 
' Line #564:
' 	Ld vYAxisKeys 
' 	Ld buildOptionLists 
' 	Ld _B_var_iColOffset 
' 	Ld _B_var_Const 
' 	Ld i 
' 	Ld xOffes 
' 	Mul 
' 	Paren 
' 	ArgsCall (Call) SubwriteAxes 0x0005 
' Line #565:
' 	EndIfBlock 
' Line #566:
' Line #567:
' 	Ld objTTX 
' 	Ld yOffset 
' 	Ld zOffset 
' 	Ld i 
' 	Ld xOffes 
' 	Mul 
' 	Ld 1 
' 	Ld lMaxHistHeigh 
' 	ArgsCall (Call) _B_var_writeResults 0x0006 
' Line #568:
' 	Ld yOffset 
' 	Ld yCoun 
' 	Gt 
' 	IfBlock 
' Line #569:
' 	Ld yOffset 
' 	St yCoun 
' Line #570:
' 	EndIfBlock 
' Line #571:
' 	Ld zOffset 
' 	Ld xCount 
' 	Gt 
' 	IfBlock 
' Line #572:
' 	Ld zOffset 
' 	St xCount 
' Line #573:
' 	EndIfBlock 
' Line #574:
' 	Ld xOffes 
' 	St lMaxHistHe 
' Line #575:
' 	EndIfBlock 
' Line #576:
' 	StartForVariable 
' 	Next 
' Line #577:
' Line #578:
' 	EndFunc 
' Line #579:
' Line #580:
' 	FuncDefn (Sub _B_var_writeResults(ByRef objTTX, yOffset, zOffset, xOffes, 1, ByRef lMaxHistHeigh))
' Line #581:
' 	Dim 
' 	VarDefn varReturn (As Variant)
' Line #582:
' 	Dim 
' 	VarDefn varChanData (As Variant)
' Line #583:
' Line #584:
' 	Dim 
' 	VarDefn dblStartTime (As Double)
' Line #585:
' 	Dim 
' 	VarDefn dblEndTime (As Double)
' Line #586:
' 	Dim 
' 	VarDefn IsEmpty (As Double)
' Line #587:
' Line #588:
' 	Dim 
' 	VarDefn i (As Long)
' Line #589:
' 	Dim 
' 	VarDefn j (As Long)
' Line #590:
' 	Dim 
' 	VarDefn k (As Long)
' Line #591:
' Line #592:
' 	Dim 
' 	VarDefn histTmp (As Long)
' Line #593:
' Line #594:
' 	LitStr 0x0004 "Swep"
' 	LitDI2 0x0000 
' 	Ld objTTX 
' 	ArgsMemLd GetEpocsExV 0x0002 
' 	St varReturn 
' Line #595:
' 	Ld varReturn 
' 	ArgsLd Dib 0x0001 
' 	IfBlock 
' Line #596:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld varReturn 
' 	LitDI2 0x0002 
' 	FnUBound 0x0001 
' 	For 
' Line #597:
' 	LitDI2 0x0002 
' 	Ld i 
' 	ArgsLd varReturn 0x0002 
' 	Ld _B_var_lIgnoreFirstMsec 
' 	Add 
' 	St dblStartTime 
' Line #598:
' 	Ld dblStartTime 
' 	Ld lBinWidth 
' 	Add 
' 	Ld _B_var_lIgnoreFirstMsec 
' 	Add 
' 	St dblEndTime 
' Line #599:
' 	Ld dblStartTime 
' 	St IsEmpty 
' Line #600:
' 	Do 
' Line #601:
' 	LitDI2 0x01F4 
' 	LitStr 0x0004 "CSPK"
' 	Ld 1 
' 	LitDI2 0x0000 
' 	Ld dblStartTime 
' 	Ld dblEndTime 
' 	LitStr 0x0009 "JUSTTIMES"
' 	Ld objTTX 
' 	ArgsMemLd ReadEventsV 0x0007 
' 	St k 
' Line #602:
' 	Ld k 
' 	LitDI2 0x0000 
' 	Eq 
' 	IfBlock 
' Line #603:
' 	ExitDo 
' Line #604:
' 	EndIfBlock 
' Line #605:
' Line #606:
' 	Ld histTmp 
' 	Coerce (Lng) 
' 	Ld k 
' 	Coerce (Lng) 
' 	Add 
' 	St histTmp 
' Line #607:
' 	Ld k 
' 	LitDI2 0x01F4 
' 	Lt 
' 	IfBlock 
' Line #608:
' 	ExitDo 
' Line #609:
' 	ElseBlock 
' Line #610:
' 	Ld k 
' 	LitDI2 0x0001 
' 	Sub 
' 	LitDI2 0x0001 
' 	LitDI2 0x0006 
' 	Ld objTTX 
' 	ArgsMemLd ParseEvInfoV 0x0003 
' 	St varChanData 
' Line #611:
' 	LitDI2 0x0000 
' 	ArgsLd varChanData 0x0001 
' 	LitDI2 0x0001 
' 	LitDI4 0x86A0 0x0001 
' 	Div 
' 	Paren 
' 	Add 
' 	St dblStartTime 
' Line #612:
' 	EndIfBlock 
' Line #613:
' 	Loop 
' Line #614:
' 	Ld IsEmpty 
' 	St dblStartTime 
' Line #615:
' 	StartForVariable 
' 	Next 
' Line #616:
' Line #617:
' 	Ld yAxisEp 
' 	LitStr 0x0007 "Channel"
' 	Eq 
' 	IfBlock 
' Line #618:
' 	Ld histTmp 
' 	Ld zOffset 
' 	Ld _B_var_Const 
' 	Add 
' 	Ld xOffes 
' 	Add 
' 	LitDI2 0x0001 
' 	Add 
' 	Ld yOffset 
' 	Ld _B_var_iColOffset 
' 	Add 
' 	LitDI2 0x0001 
' 	Add 
' 	LitStr 0x0006 "Output"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Cells 0x0002 
' 	MemSt Value 
' Line #619:
' 	Ld otherEp 
' 	LitStr 0x0007 "Channel"
' 	Eq 
' 	ElseIfBlock 
' Line #620:
' 	Ld histTmp 
' 	Ld zOffset 
' 	Ld _B_var_Const 
' 	Add 
' 	Ld xOffes 
' 	Add 
' 	LitDI2 0x0001 
' 	Add 
' 	Ld yOffset 
' 	Ld _B_var_iColOffset 
' 	Add 
' 	LitDI2 0x0001 
' 	Add 
' 	LitStr 0x0006 "Output"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Cells 0x0002 
' 	MemSt Value 
' Line #621:
' 	ElseBlock 
' Line #622:
' 	Ld histTmp 
' 	Ld zOffset 
' 	Ld _B_var_Const 
' 	Add 
' 	Ld xOffes 
' 	Add 
' 	LitDI2 0x0001 
' 	Add 
' 	Ld yOffset 
' 	Ld _B_var_iColOffset 
' 	Add 
' 	LitDI2 0x0001 
' 	Add 
' 	LitStr 0x0006 "Output"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Cells 0x0002 
' 	MemSt Value 
' Line #623:
' 	EndIfBlock 
' Line #624:
' 	Ld histTmp 
' 	Ld lMaxHistHeigh 
' 	Gt 
' 	IfBlock 
' Line #625:
' 	Ld histTmp 
' 	St lMaxHistHeigh 
' Line #626:
' 	EndIfBlock 
' Line #627:
' 	LitDI2 0x0000 
' 	St histTmp 
' Line #628:
' 	EndIfBlock 
' Line #629:
' Line #630:
' 	EndSub 
' Line #631:
' Line #632:
' 	FuncDefn (Sub Ra())
' Line #633:
' 	Dim 
' 	VarDefn lMaxHistHe (As Long)
' Line #634:
' 	Dim 
' 	VarDefn _B_var_iColOffset (As Integer)
' Line #635:
' 	Dim 
' 	VarDefn _B_var_Const (As Integer)
' Line #636:
' Line #637:
' 	LitStr 0x0002 "H3"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	St lMaxHistHe 
' Line #638:
' 	LitStr 0x0002 "H5"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	St _B_var_iColOffset 
' Line #639:
' 	LitStr 0x0002 "H6"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	St _B_var_Const 
' Line #640:
' Line #641:
' 	Dim 
' 	VarDefn yPos (As Long)
' Line #642:
' 	Dim 
' 	VarDefn Whi0l (As Long)
' Line #643:
' Line #644:
' 	Ld _B_var_iColOffset 
' 	LitDI2 0x0001 
' 	Add 
' 	St yPos 
' Line #645:
' 	Ld _B_var_Const 
' 	St Whi0l 
' Line #646:
' Line #647:
' 	SetStmt 
' 	New id_FFFF
' 	Set dHeading 
' Line #648:
' Line #649:
' 	Do 
' Line #650:
' 	Ld Whi0l 
' 	Ld yPos 
' 	LitStr 0x0006 "Output"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Cells 0x0002 
' 	MemLd Value 
' 	LitStr 0x0000 ""
' 	Ne 
' 	IfBlock 
' Line #651:
' 	Ld Whi0l 
' 	Ld yPos 
' 	LitStr 0x0006 "Output"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Cells 0x0002 
' 	MemLd Value 
' 	Ld dHeading 
' 	ArgsMemLd Exists 0x0001 
' 	Not 
' 	IfBlock 
' Line #652:
' 	Ld Whi0l 
' 	Ld yPos 
' 	LitStr 0x0006 "Output"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Cells 0x0002 
' 	MemLd Value 
' 	LitDI2 0x0000 
' 	Ld dHeading 
' 	ArgsMemCall (Call) Add 0x0002 
' Line #653:
' 	EndIfBlock 
' Line #654:
' 	Ld Whi0l 
' 	Ld lMaxHistHe 
' 	Add 
' 	St Whi0l 
' Line #655:
' 	ElseBlock 
' Line #656:
' 	ExitDo 
' Line #657:
' 	EndIfBlock 
' Line #658:
' 	Loop 
' Line #659:
' Line #660:
' 	Ld transferToSigmaplotButton 
' 	ArgsMemCall Show 0x0000 
' Line #661:
' 	Ld doImport 
' 	IfBlock 
' Line #662:
' 	ArgsCall (Call) ACTIVESPWLib 0x0000 
' Line #663:
' 	EndIfBlock 
' Line #664:
' Line #665:
' 	EndSub 
' _VBA_PROJECT_CUR/VBA/Sheet2 - 1166 bytes
' _VBA_PROJECT_CUR/VBA/Sheet3 - 1150 bytes
' _VBA_PROJECT_CUR/VBA/Sheet4 - 1166 bytes
' _VBA_PROJECT_CUR/VBA/TransferToSigmaplotFrm - 6120 bytes
' Line #0:
' 	FuncDefn (Private Sub Cancel_Click())
' Line #1:
' 	LitVarSpecial (False)
' 	St doImport 
' Line #2:
' 	Ld id_FFFF 
' 	ArgsCall Unload 0x0001 
' 	QuoteRem 0x0015 0x0015 "Unloads the UserForm."
' Line #3:
' 	EndSub 
' Line #4:
' Line #5:
' 	FuncDefn (Private Sub SelectAll())
' Line #6:
' 	Dim 
' 	VarDefn i (As Integer)
' Line #7:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld HeadingsList 
' 	MemLd ListIndex 
' 	LitDI2 0x0001 
' 	Sub 
' 	Paren 
' 	For 
' Line #8:
' 	LitVarSpecial (True)
' 	Ld i 
' 	Ld HeadingsList 
' 	ArgsMemSt Clear 0x0001 
' Line #9:
' 	StartForVariable 
' 	Next 
' Line #10:
' 	EndSub 
' Line #11:
' Line #12:
' 	FuncDefn (Private Sub SelectAll_Click())
' Line #13:
' 	Dim 
' 	VarDefn i (As Integer)
' Line #14:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld HeadingsList 
' 	MemLd ListIndex 
' 	LitDI2 0x0001 
' 	Sub 
' 	Paren 
' 	For 
' Line #15:
' 	LitVarSpecial (False)
' 	Ld i 
' 	Ld HeadingsList 
' 	ArgsMemSt Clear 0x0001 
' Line #16:
' 	StartForVariable 
' 	Next 
' Line #17:
' 	EndSub 
' Line #18:
' Line #19:
' 	FuncDefn (Private Sub dSelected())
' Line #20:
' 	SetStmt 
' 	LitNothing 
' 	Set bXAxisLog 
' Line #21:
' 	SetStmt 
' 	New id_FFFF
' 	Set bXAxisLog 
' Line #22:
' Line #23:
' 	Dim 
' 	VarDefn i (As Integer)
' Line #24:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld HeadingsList 
' 	MemLd ListIndex 
' 	LitDI2 0x0001 
' 	Sub 
' 	Paren 
' 	For 
' Line #25:
' 	Ld i 
' 	Ld HeadingsList 
' 	ArgsMemLd Clear 0x0001 
' 	LitVarSpecial (True)
' 	Eq 
' 	IfBlock 
' Line #26:
' 	Ld i 
' 	Ld HeadingsList 
' 	ArgsMemLd Listn 0x0001 
' 	Ld bXAxisLog 
' 	ArgsMemLd Exists 0x0001 
' 	Not 
' 	IfBlock 
' Line #27:
' 	Ld i 
' 	Ld HeadingsList 
' 	ArgsMemLd Listn 0x0001 
' 	Ld i 
' 	Ld bXAxisLog 
' 	ArgsMemCall (Call) Add 0x0002 
' Line #28:
' 	EndIfBlock 
' Line #29:
' 	EndIfBlock 
' Line #30:
' 	StartForVariable 
' 	Next 
' Line #31:
' Line #32:
' 	LitVarSpecial (True)
' 	St doImport 
' Line #33:
' 	Ld id_FFFF 
' 	ArgsCall Unload 0x0001 
' 	QuoteRem 0x0015 0x0015 "Unloads the UserForm."
' Line #34:
' 	EndSub 
' Line #35:
' Line #36:
' 	FuncDefn (Private Sub UserForm_Activate())
' Line #37:
' 	Dim 
' 	VarDefn _B_var_HeadingsList
' Line #38:
' 	Ld dHeading 
' 	MemLd Keys 
' 	St _B_var_HeadingsList 
' Line #39:
' Line #40:
' 	Dim 
' 	VarDefn i (As Integer)
' Line #41:
' Line #42:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld _B_var_HeadingsList 
' 	FnUBound 0x0000 
' 	For 
' Line #43:
' 	Ld i 
' 	ArgsLd _B_var_HeadingsList 0x0001 
' 	Ld i 
' 	Ld HeadingsList 
' 	ArgsMemCall (Call) EVTYPE_STRON 0x0002 
' Line #44:
' 	LitVarSpecial (True)
' 	Ld i 
' 	Ld HeadingsList 
' 	ArgsMemSt Clear 0x0001 
' Line #45:
' 	StartForVariable 
' 	Next 
' Line #46:
' 	EndSub 
