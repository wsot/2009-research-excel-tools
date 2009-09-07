' Processing file: Tuning curve generator - generic.xls
' ===============================================================================
' Module streams:
' _VBA_PROJECT_CUR/VBA/ThisWorkbook - 1210 bytes
' Line #0:
' 	Option  (Explicit)
' Line #1:
' _VBA_PROJECT_CUR/VBA/Sheet1 - 1150 bytes
' _VBA_PROJECT_CUR/VBA/ImportFrom - 22376 bytes
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
' _VBA_PROJECT_CUR/VBA/Module1 - 47991 bytes
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
' 	Dim 
' 	VarDefn vYAxisKeys (As Variant)
' Line #11:
' 	Dim 
' 	VarDefn buildOptionLists (As Variant)
' Line #12:
' Line #13:
' Line #14:
' 	FuncDefn (Sub buildTuningCurves())
' Line #15:
' 	Ld ImportFrom 
' 	ArgsMemCall Show 0x0000 
' Line #16:
' Line #17:
' 	Ld doImport 
' 	IfBlock 
' Line #18:
' 	LitVarSpecial (False)
' 	ArgsCall (Call) processImport 0x0001 
' Line #19:
' 	EndIfBlock 
' Line #20:
' 	EndSub 
' Line #21:
' Line #22:
' 	FuncDefn (Sub importIntoSigmaplot())
' Line #23:
' 	QuoteRem 0x0000 0x0013 "    ImportFrom.Show"
' Line #24:
' Line #25:
' 	QuoteRem 0x0000 0x0014 "    If doImport Then"
' Line #26:
' 	QuoteRem 0x0000 0x0020 "        Call processImport(True)"
' Line #27:
' 	QuoteRem 0x0000 0x000A "    End If"
' Line #28:
' 	ArgsCall (Call) ACTIVESPWLib 0x0000 
' Line #29:
' 	EndSub 
' Line #30:
' Line #31:
' 	FuncDefn (Sub processImport(spNB As Boolean))
' Line #32:
' Line #33:
' 	QuoteRem 0x0004 0x002B "load the bin width for histogram generation"
' Line #34:
' 	LitStr 0x0002 "B1"
' 	LitStr 0x0008 "Settings"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	St lBinWidth 
' Line #35:
' Line #36:
' 	QuoteRem 0x0004 0x004D "load the # of msec to ignore at the start (for filtering stimulation artifact"
' Line #37:
' 	LitStr 0x0002 "B2"
' 	LitStr 0x0008 "Settings"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	St _B_var_lIgnoreFirstMsec 
' Line #38:
' Line #39:
' 	QuoteRem 0x0004 0x003A "used to store the maximum histogram peak for normalisation"
' Line #40:
' 	Dim 
' 	VarDefn lMaxHistHeigh (As Double)
' Line #41:
' 	LitDI2 0x0000 
' 	St lMaxHistHeigh 
' Line #42:
' Line #43:
' 	Dim 
' 	VarDefn theWorksheets (As Variant)
' 	QuoteRem 0x0021 0x0029 "stores the created worksheets to write to"
' Line #44:
' 	Dim 
' 	VarDefn _B_var_arrHistTmp (As Long)
' 	QuoteRem 0x001D 0x0044 "used to store the histogram data for each channel as it is generated"
' Line #45:
' 	OptionBase 
' 	LitDI2 0x001F 
' 	Redim _B_var_arrHistTmp 0x0001 (As Variant)
' Line #46:
' Line #47:
' 	Dim 
' 	VarDefn xCount (As Long)
' Line #48:
' 	Dim 
' 	VarDefn yCoun (As Long)
' Line #49:
' 	Dim 
' 	VarDefn lMaxHistHe (As Long)
' Line #50:
' Line #51:
' 	QuoteRem 0x0004 0x0037 "offsets to leave space at the top and left of the chart"
' Line #52:
' 	LitDI2 0x0001 
' 	St _B_var_Const 
' Line #53:
' 	LitDI2 0x0000 
' 	St _B_var_iColOffset 
' Line #54:
' Line #55:
' 	QuoteRem 0x0000 0x0050 "    theWorksheets = buildWorksheetArray() 'build the worksheets for writing data"
' Line #56:
' Line #57:
' 	QuoteRem 0x0004 0x0013 "connect to the tank"
' Line #58:
' 	Dim 
' 	VarDefn objTTX
' Line #59:
' 	SetStmt 
' 	LitStr 0x0007 "TTank.X"
' 	ArgsLd CreateObject 0x0001 
' 	Set objTTX 
' Line #60:
' Line #61:
' 	Ld theServer 
' 	LitStr 0x0002 "Me"
' 	Ld objTTX 
' 	ArgsMemLd ConnectServer 0x0002 
' 	LitDI2 0x0001 
' 	Coerce (Lng) 
' 	Ne 
' 	IfBlock 
' Line #62:
' 	LitStr 0x0015 "Connecting to server "
' 	Ld theServer 
' 	Concat 
' 	LitStr 0x0008 " failed."
' 	Concat 
' 	Paren 
' 	ArgsCall MsgBox 0x0001 
' Line #63:
' 	ExitSub 
' Line #64:
' 	EndIfBlock 
' Line #65:
' Line #66:
' 	Ld theTank 
' 	LitStr 0x0001 "R"
' 	Ld objTTX 
' 	ArgsMemLd OpenTank 0x0002 
' 	LitDI2 0x0001 
' 	Coerce (Lng) 
' 	Ne 
' 	IfBlock 
' Line #67:
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
' Line #68:
' 	Ld objTTX 
' 	ArgsMemCall (Call) ReleaseServer 0x0000 
' Line #69:
' 	ExitSub 
' Line #70:
' 	EndIfBlock 
' Line #71:
' Line #72:
' 	Ld theBlock 
' 	Ld objTTX 
' 	ArgsMemLd SelectBlock 0x0001 
' 	LitDI2 0x0001 
' 	Coerce (Lng) 
' 	Ne 
' 	IfBlock 
' Line #73:
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
' Line #74:
' 	Ld objTTX 
' 	ArgsMemCall (Call) CloseTank 0x0000 
' Line #75:
' 	Ld objTTX 
' 	ArgsMemCall (Call) ReleaseServer 0x0000 
' Line #76:
' 	ExitSub 
' Line #77:
' 	EndIfBlock 
' Line #78:
' Line #79:
' 	QuoteRem 0x0004 0x0026 "index epochs - required to use filters"
' Line #80:
' 	Ld objTTX 
' 	ArgsMemCall (Call) CreateEpocIndexing 0x0000 
' Line #81:
' Line #82:
' 	Dim 
' 	VarDefn dblStartTime (As Double)
' Line #83:
' 	Dim 
' 	VarDefn dblEndTime (As Double)
' Line #84:
' Line #85:
' 	Dim 
' 	VarDefn varReturn (As Variant)
' Line #86:
' Line #87:
' 	QuoteRem 0x0000 0x001D "    Dim vXAxisKeys As Variant"
' Line #88:
' 	QuoteRem 0x0000 0x001D "    Dim vYAxisKeys As Variant"
' Line #89:
' Line #90:
' 	Ld objTTX 
' 	Ld yAxisEp 
' 	Ld bReverseY 
' 	ArgsLd _B_var_buildEpocList 0x0003 
' 	St vYAxisKeys 
' Line #91:
' 	Ld objTTX 
' 	Ld otherEp 
' 	Ld ReverseX 
' 	ArgsLd _B_var_buildEpocList 0x0003 
' 	St buildOptionLists 
' Line #92:
' Line #93:
' 	Dim 
' 	VarDefn i (As Long)
' Line #94:
' 	Dim 
' 	VarDefn j (As Long)
' Line #95:
' 	Dim 
' 	VarDefn k (As Long)
' Line #96:
' 	Dim 
' 	VarDefn l (As Long)
' Line #97:
' Line #98:
' 	Dim 
' 	VarDefn strSearchString (As Variant)
' Line #99:
' 	Ld _B_var_arrOtherEp 
' 	FnUBound 0x0000 
' 	LitDI2 0x0001 
' 	UMi 
' 	Ne 
' 	IfBlock 
' Line #100:
' 	OptionBase 
' 	Ld _B_var_arrOtherEp 
' 	FnUBound 0x0000 
' 	Redim strSearchString 0x0001 (As Variant)
' Line #101:
' Line #102:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld _B_var_arrOtherEp 
' 	FnUBound 0x0000 
' 	For 
' Line #103:
' 	Ld objTTX 
' 	Ld i 
' 	ArgsLd _B_var_arrOtherEp 0x0001 
' 	LitVarSpecial (False)
' 	ArgsLd _B_var_buildEpocList 0x0003 
' 	Ld i 
' 	ArgsSt strSearchString 0x0001 
' Line #104:
' 	StartForVariable 
' 	Next 
' Line #105:
' 	EndIfBlock 
' Line #106:
' Line #107:
' 	LitDI2 0x0000 
' 	St i 
' Line #108:
' 	LitDI2 0x0000 
' 	St j 
' Line #109:
' Line #110:
' 	Dim 
' 	VarDefn iYAxisIndex (As Integer)
' Line #111:
' 	Dim 
' 	VarDefn otherEpocList (As Integer)
' Line #112:
' 	Dim 
' 	VarDefn _B_var_arrOtherEpocIndex (As Integer)
' Line #113:
' 	Ld _B_var_arrOtherEp 
' 	FnUBound 0x0000 
' 	LitDI2 0x0001 
' 	UMi 
' 	Ne 
' 	IfBlock 
' Line #114:
' 	OptionBase 
' 	Ld _B_var_arrOtherEp 
' 	FnUBound 0x0000 
' 	Redim _B_var_arrOtherEpocIndex 0x0001 (As Variant)
' Line #115:
' 	EndIfBlock 
' Line #116:
' Line #117:
' 	Dim 
' 	VarDefn varChanData (As Variant)
' Line #118:
' 	Dim 
' 	VarDefn IsEmpty (As Double)
' Line #119:
' Line #120:
' 	Dim 
' 	VarDefn yAxisSearchString (As String)
' Line #121:
' 	Dim 
' 	VarDefn otherAxisSearchString (As String)
' Line #122:
' 	Dim 
' 	VarDefn processSearch (As String)
' Line #123:
' 	Dim 
' 	VarDefn arrOtherEpFor (As String)
' Line #124:
' 	Ld _B_var_arrOtherEp 
' 	FnUBound 0x0000 
' 	LitDI2 0x0001 
' 	UMi 
' 	Ne 
' 	IfBlock 
' Line #125:
' 	OptionBase 
' 	Ld _B_var_arrOtherEp 
' 	FnUBound 0x0000 
' 	Redim processSearch 0x0001 (As Variant)
' Line #126:
' 	EndIfBlock 
' Line #127:
' Line #128:
' 	Dim 
' 	VarDefn 1 (As Integer)
' Line #129:
' 	LitDI2 0x0000 
' 	St 1 
' Line #130:
' Line #131:
' 	Ld _B_var_arrOtherEp 
' 	FnUBound 0x0000 
' 	LitDI2 0x0001 
' 	UMi 
' 	Ne 
' 	IfBlock 
' Line #132:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld vYAxisKeys 
' 	FnUBound 0x0000 
' 	For 
' Line #133:
' 	Ld yAxisEp 
' 	LitStr 0x0007 "Channel"
' 	Eq 
' 	IfBlock 
' Line #134:
' 	Ld i 
' 	ArgsLd vYAxisKeys 0x0001 
' 	St 1 
' Line #135:
' 	LitStr 0x0000 ""
' 	St yAxisSearchString 
' Line #136:
' 	ElseBlock 
' Line #137:
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
' Line #138:
' 	EndIfBlock 
' Line #139:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld buildOptionLists 
' 	FnUBound 0x0000 
' 	For 
' Line #140:
' 	Ld otherEp 
' 	LitStr 0x0007 "Channel"
' 	Eq 
' 	IfBlock 
' Line #141:
' 	Ld j 
' 	ArgsLd buildOptionLists 0x0001 
' 	St 1 
' Line #142:
' 	LitStr 0x0000 ""
' 	St otherAxisSearchString 
' Line #143:
' 	ElseBlock 
' Line #144:
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
' Line #145:
' 	EndIfBlock 
' Line #146:
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
' Line #147:
' 	StartForVariable 
' 	Next 
' Line #148:
' 	StartForVariable 
' 	Next 
' Line #149:
' 	EndIfBlock 
' Line #150:
' Line #151:
' 	QuoteRem 0x0000 0x0051 "    Call writeAxes(theWorksheets, vXAxisKeys, vYAxisKeys, iColOffset, iRowOffset)"
' Line #152:
' Line #153:
' 	Ld objTTX 
' 	ArgsMemCall (Call) CloseTank 0x0000 
' Line #154:
' 	Ld objTTX 
' 	ArgsMemCall (Call) ReleaseServer 0x0000 
' Line #155:
' Line #156:
' 	Ld yCoun 
' 	LitStr 0x0002 "H1"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #157:
' 	Ld xCount 
' 	LitStr 0x0002 "H2"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #158:
' 	Ld lMaxHistHe 
' 	LitStr 0x0002 "H3"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #159:
' 	Ld lMaxHistHeigh 
' 	LitStr 0x0002 "H4"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #160:
' 	Ld _B_var_iColOffset 
' 	LitStr 0x0002 "H5"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #161:
' 	Ld _B_var_Const 
' 	LitStr 0x0002 "H6"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #162:
' Line #163:
' 	QuoteRem 0x0004 0x001B "If importIntoSigmaplot Then"
' Line #164:
' 	QuoteRem 0x0008 0x005D "Call transferToSigmaplot(xCount, yCount, zOffsetSize, iColOffset, iRowOffset, lMaxHistHeight)"
' Line #165:
' 	QuoteRem 0x0004 0x0006 "End If"
' Line #166:
' Line #167:
' 	EndSub 
' Line #168:
' Line #169:
' 	FuncDefn (Function buildWorksheetArray() As Variant)
' Line #170:
' 	Dim 
' 	OptionBase 
' 	LitDI2 0x001F 
' 	VarDefn theWorksheets
' Line #171:
' 	Dim 
' 	VarDefn strWsname (As String)
' Line #172:
' 	Dim 
' 	VarDefn intWSNum (As Long)
' Line #173:
' Line #174:
' 	Dim 
' 	VarDefn i (As Integer)
' Line #175:
' Line #176:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0001 
' 	Ld Worksheets 
' 	MemLd Count 
' 	For 
' Line #177:
' 	Ld i 
' 	Ld Worksheets 
' 	ArgsMemLd Item 0x0001 
' 	MemLd Name 
' 	St strWsname 
' Line #178:
' 	Ld strWsname 
' 	LitDI2 0x0004 
' 	ArgsLd Left 0x0002 
' 	LitStr 0x0004 "Site"
' 	Eq 
' 	IfBlock 
' Line #179:
' 	Ld strWsname 
' 	Ld strWsname 
' 	FnLen 
' 	LitDI2 0x0004 
' 	Sub 
' 	ArgsLd Right 0x0002 
' 	ArgsLd IsNumeric 0x0001 
' 	IfBlock 
' Line #180:
' 	Ld strWsname 
' 	Ld strWsname 
' 	FnLen 
' 	LitDI2 0x0004 
' 	Sub 
' 	ArgsLd Right 0x0002 
' 	Coerce (Int) 
' 	St intWSNum 
' Line #181:
' 	Ld intWSNum 
' 	LitDI2 0x0021 
' 	Lt 
' 	Ld intWSNum 
' 	LitDI2 0x0000 
' 	Gt 
' 	And 
' 	IfBlock 
' Line #182:
' 	SetStmt 
' 	Ld i 
' 	Ld Worksheets 
' 	ArgsMemLd Item 0x0001 
' 	Ld intWSNum 
' 	LitDI2 0x0001 
' 	Sub 
' 	ArgsSet theWorksheets 0x0001 
' Line #183:
' 	EndIfBlock 
' Line #184:
' 	EndIfBlock 
' Line #185:
' 	EndIfBlock 
' Line #186:
' 	StartForVariable 
' 	Next 
' Line #187:
' Line #188:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	LitDI2 0x001F 
' 	For 
' Line #189:
' 	Ld i 
' 	ArgsLd theWorksheets 0x0001 
' 	ArgsLd Sheet70 0x0001 
' 	IfBlock 
' Line #190:
' 	Ld i 
' 	LitDI2 0x0000 
' 	Gt 
' 	IfBlock 
' Line #191:
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
' Line #192:
' 	ElseBlock 
' Line #193:
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
' Line #194:
' 	EndIfBlock 
' Line #195:
' 	LitStr 0x0004 "Site"
' 	Ld i 
' 	LitDI2 0x0001 
' 	Add 
' 	Coerce (Str) 
' 	Concat 
' 	Ld i 
' 	ArgsLd theWorksheets 0x0001 
' 	MemSt Name 
' Line #196:
' 	EndIfBlock 
' Line #197:
' 	StartForVariable 
' 	Next 
' Line #198:
' 	Ld theWorksheets 
' 	St buildWorksheetArray 
' Line #199:
' 	EndFunc 
' Line #200:
' Line #201:
' 	FuncDefn (Sub SubwriteAxes(rowLabels As Variant, deleteWorksheets As Variant, _B_var_iColOffset, _B_var_Const, xOffes))
' Line #202:
' 	Dim 
' 	VarDefn j (As Long)
' Line #203:
' Line #204:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld deleteWorksheets 
' 	FnUBound 0x0000 
' 	For 
' Line #205:
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
' Line #206:
' 	StartForVariable 
' 	Next 
' Line #207:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld rowLabels 
' 	FnUBound 0x0000 
' 	For 
' Line #208:
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
' Line #209:
' 	StartForVariable 
' 	Next 
' Line #210:
' Line #211:
' 	EndSub 
' Line #212:
' Line #213:
' 	FuncDefn (Sub Delete())
' Line #214:
' 	Dim 
' 	VarDefn strWsname (As String)
' Line #215:
' 	Dim 
' 	VarDefn intWSNum (As Long)
' Line #216:
' Line #217:
' 	Dim 
' 	VarDefn i (As Integer)
' Line #218:
' Line #219:
' 	Ld Worksheets 
' 	MemLd Count 
' 	St i 
' Line #220:
' Line #221:
' 	Do 
' Line #222:
' 	Ld i 
' 	LitDI2 0x0000 
' 	Eq 
' 	IfBlock 
' Line #223:
' 	ExitDo 
' Line #224:
' 	EndIfBlock 
' Line #225:
' 	Ld i 
' 	Ld Worksheets 
' 	ArgsMemLd Item 0x0001 
' 	MemLd Name 
' 	St strWsname 
' Line #226:
' 	Ld strWsname 
' 	LitDI2 0x0004 
' 	ArgsLd Left 0x0002 
' 	LitStr 0x0004 "Site"
' 	Eq 
' 	IfBlock 
' Line #227:
' 	Ld strWsname 
' 	Ld strWsname 
' 	FnLen 
' 	LitDI2 0x0004 
' 	Sub 
' 	ArgsLd Right 0x0002 
' 	ArgsLd IsNumeric 0x0001 
' 	IfBlock 
' Line #228:
' 	Ld strWsname 
' 	Ld strWsname 
' 	FnLen 
' 	LitDI2 0x0004 
' 	Sub 
' 	ArgsLd Right 0x0002 
' 	Coerce (Int) 
' 	St intWSNum 
' Line #229:
' 	Ld intWSNum 
' 	LitDI2 0x0021 
' 	Lt 
' 	Ld intWSNum 
' 	LitDI2 0x0000 
' 	Gt 
' 	And 
' 	IfBlock 
' Line #230:
' 	Ld i 
' 	Ld Worksheets 
' 	ArgsMemLd Item 0x0001 
' 	ArgsMemCall UserForm1 0x0000 
' Line #231:
' 	EndIfBlock 
' Line #232:
' 	EndIfBlock 
' Line #233:
' 	EndIfBlock 
' Line #234:
' 	Ld i 
' 	LitDI2 0x0001 
' 	Sub 
' 	St i 
' Line #235:
' 	Loop 
' Line #236:
' 	EndSub 
' Line #237:
' Line #238:
' 	FuncDefn (Sub ACTIVESPWLib())
' Line #239:
' Line #240:
' 	Dim 
' 	VarDefn yCoun (As Long)
' Line #241:
' 	Dim 
' 	VarDefn xCount (As Long)
' Line #242:
' 	Dim 
' 	VarDefn lMaxHistHe (As Long)
' Line #243:
' 	Dim 
' 	VarDefn lMaxHistHeigh (As Long)
' Line #244:
' 	Dim 
' 	VarDefn _B_var_iColOffset (As Integer)
' Line #245:
' 	Dim 
' 	VarDefn _B_var_Const (As Integer)
' Line #246:
' Line #247:
' 	LitStr 0x0002 "H1"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	St yCoun 
' Line #248:
' 	LitStr 0x0002 "H2"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	St xCount 
' Line #249:
' 	LitStr 0x0002 "H3"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	St lMaxHistHe 
' Line #250:
' 	LitStr 0x0002 "H4"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	St lMaxHistHeigh 
' Line #251:
' 	LitStr 0x0002 "H5"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	St _B_var_iColOffset 
' Line #252:
' 	LitStr 0x0002 "H6"
' 	LitStr 0x0017 "Variables (do not edit)"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	St _B_var_Const 
' Line #253:
' Line #254:
' 	Dim (Const) 
' 	LitHI2 0x0406 
' 	VarDefn SAA_TOVAL
' Line #255:
' 	Dim (Const) 
' 	LitHI2 0x0407 
' 	VarDefn GraphPages
' Line #256:
' 	Dim (Const) 
' 	LitHI2 0x0301 
' 	VarDefn SLA_SELECTDIM
' Line #257:
' 	Dim (Const) 
' 	LitDI2 0x0401 
' 	VarDefn SEA_COLORCOL
' Line #258:
' 	Dim (Const) 
' 	LitDI2 0x0308 
' 	VarDefn SAA_OPTIONS
' Line #259:
' 	Dim (Const) 
' 	LitDI2 0x0403 
' 	VarDefn _B_var_GPM_SETPLOTATTR
' Line #260:
' 	Dim (Const) 
' 	LitDI2 0x0408 
' 	VarDefn SAA_FROMVAL
' Line #261:
' 	Dim (Const) 
' 	LitDI2 0x0615 
' 	VarDefn GPM_SETAXISATTRSTRING
' Line #262:
' 	Dim (Const) 
' 	LitDI2 0x0613 
' 	VarDefn SLA_CONTOURFILLTYPE
' Line #263:
' 	Dim (Const) 
' 	LitDI2 0x0358 
' 	VarDefn SAA_SELECTLINE
' Line #264:
' 	Dim (Const) 
' 	LitDI2 0x040A 
' 	VarDefn SEA_THICKNESS
' Line #265:
' 	Dim (Const) 
' 	LitDI2 0x0601 
' 	VarDefn SEA_COLOR
' Line #266:
' 	Dim (Const) 
' 	LitDI2 0x0606 
' 	VarDefn _B_var_SEA_THICKNESS
' Line #267:
' 	Dim (Const) 
' 	LitDI2 0x0410 
' 	VarDefn _B_var_SAA_SUB1OPTIONS
' Line #268:
' Line #269:
' 	Dim 
' 	VarDefn Module2 (As Object)
' Line #270:
' 	SetStmt 
' 	LitStr 0x0017 "SigmaPlot.Application.1"
' 	ArgsLd CreateObject 0x0001 
' 	Set Module2 
' Line #271:
' 	LitVarSpecial (True)
' 	Ld Module2 
' 	MemSt Application 
' Line #272:
' 	Ld Module2 
' 	MemLd Notebooks 
' 	MemLd buildTuningCurvesIntoSigmaplot 
' 	ArgsMemCall (Call) Add 0x0000 
' Line #273:
' Line #274:
' 	Dim 
' 	VarDefn i (As Integer)
' Line #275:
' 	Dim 
' 	VarDefn j (As Long)
' Line #276:
' 	Dim 
' 	VarDefn k (As Long)
' Line #277:
' Line #278:
' 	Dim 
' 	VarDefn yPos (As Long)
' Line #279:
' 	Dim 
' 	VarDefn Whi0l (As Long)
' Line #280:
' Line #281:
' 	Dim 
' 	VarDefn SPApplication (As Object)
' Line #282:
' 	Dim 
' 	VarDefn spDT (As Object)
' Line #283:
' 	Dim 
' 	VarDefn DataTable (As Object)
' Line #284:
' 	Dim 
' 	VarDefn objSPWizard (As Object)
' Line #285:
' Line #286:
' 	Ld _B_var_iColOffset 
' 	LitDI2 0x0001 
' 	Add 
' 	St yPos 
' Line #287:
' 	Ld _B_var_Const 
' 	St Whi0l 
' Line #288:
' Line #289:
' 	Do 
' Line #290:
' 	Ld Whi0l 
' 	Ld yPos 
' 	LitStr 0x0006 "Output"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Cells 0x0002 
' 	MemLd Value 
' 	LitStr 0x0000 ""
' 	Ne 
' 	IfBlock 
' Line #291:
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
' Line #292:
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
' Line #293:
' 	Ld Whi0l 
' 	Ld yPos 
' 	LitStr 0x0006 "Output"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Cells 0x0002 
' 	MemLd Value 
' 	Ld spDT 
' 	MemSt Name 
' Line #294:
' 	SetStmt 
' 	Ld spDT 
' 	MemLd Cell 
' 	Set DataTable 
' Line #295:
' Line #296:
' 	Ld Whi0l 
' 	St Whi0l 
' Line #297:
' Line #298:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld yCoun 
' 	LitDI2 0x0001 
' 	Sub 
' 	Paren 
' 	For 
' Line #299:
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
' Line #300:
' 	StartForVariable 
' 	Next 
' Line #301:
' Line #302:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld xCount 
' 	LitDI2 0x0001 
' 	Sub 
' 	Paren 
' 	For 
' Line #303:
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
' Line #304:
' 	StartForVariable 
' 	Next 
' Line #305:
' Line #306:
' Line #307:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld yCoun 
' 	LitDI2 0x0001 
' 	Sub 
' 	Paren 
' 	For 
' Line #308:
' 	StartForVariable 
' 	Ld k 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld xCount 
' 	LitDI2 0x0001 
' 	Sub 
' 	Paren 
' 	For 
' Line #309:
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
' Line #310:
' 	StartForVariable 
' 	Next 
' Line #311:
' 	StartForVariable 
' 	Next 
' Line #312:
' Line #313:
' 	LitStr 0x0011 "@rgb(255,255,255)"
' 	LitDI2 0x0002 
' 	LitDI2 0x0000 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #314:
' 	LitStr 0x000D "@rgb(0,0,255)"
' 	LitDI2 0x0002 
' 	LitDI2 0x0001 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #315:
' 	LitStr 0x000F "@rgb(0,255,255)"
' 	LitDI2 0x0002 
' 	LitDI2 0x0002 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #316:
' 	LitStr 0x000D "@rgb(0,255,0)"
' 	LitDI2 0x0002 
' 	LitDI2 0x0003 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #317:
' 	LitStr 0x000F "@rgb(255,255,0)"
' 	LitDI2 0x0002 
' 	LitDI2 0x0004 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #318:
' 	LitStr 0x000D "@rgb(255,0,0)"
' 	LitDI2 0x0002 
' 	LitDI2 0x0005 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #319:
' Line #320:
' 	QuoteRem 0x000C 0x001E "Call spNB.NotebookItems.Add(2)"
' Line #321:
' 	QuoteRem 0x000C 0x0042 "Set spGRPH = spNB.NotebookItems.Item(spNB.NotebookItems.Count - 1)"
' Line #322:
' Line #323:
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd SPWNotebookComponentType 
' 	ArgsMemCall (Call) Add 0x0001 
' Line #324:
' 	Dim 
' 	OptionBase 
' 	LitDI2 0x0002 
' 	OptionBase 
' 	LitDI2 0x0003 
' 	VarDefn PlotColumnCountArray
' Line #325:
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #326:
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	LitDI2 0x0000 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #327:
' 	LitDI4 0x47FF 0x01E8 
' 	LitDI2 0x0002 
' 	LitDI2 0x0000 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #328:
' 	LitDI2 0x0001 
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #329:
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	LitDI2 0x0001 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #330:
' 	LitDI4 0x47FF 0x01E8 
' 	LitDI2 0x0002 
' 	LitDI2 0x0001 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #331:
' 	LitDI2 0x0003 
' 	LitDI2 0x0000 
' 	LitDI2 0x0002 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #332:
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	LitDI2 0x0002 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #333:
' 	LitDI4 0x47FF 0x01E8 
' 	LitDI2 0x0002 
' 	LitDI2 0x0002 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #334:
' 	LitDI2 0x0003 
' 	Ld xCount 
' 	LitDI2 0x0001 
' 	Sub 
' 	Paren 
' 	Add 
' 	LitDI2 0x0000 
' 	LitDI2 0x0003 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #335:
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	LitDI2 0x0003 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #336:
' 	LitDI4 0x47FF 0x01E8 
' 	LitDI2 0x0002 
' 	LitDI2 0x0003 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #337:
' Line #338:
' 	Dim 
' 	VarDefn CurrentPageItem
' Line #339:
' 	OptionBase 
' 	LitDI2 0x0000 
' 	Redim CurrentPageItem 0x0001 (As Variant)
' Line #340:
' Line #341:
' 	LitDI2 0x0004 
' 	LitDI2 0x0000 
' 	ArgsSt CurrentPageItem 0x0001 
' Line #342:
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
' Line #343:
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemLd Graphs 0x0001 
' 	ArgsMemLd SelectObject 0x0001 
' 	ArgsMemCall (Call) testSigmaPlot 0x0000 
' Line #344:
' Line #345:
' 	LitStr 0x0006 "Site y"
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemLd Graphs 0x0001 
' 	ArgsMemLd SelectObject 0x0001 
' 	MemSt Name 
' Line #346:
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
' Line #347:
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
' Line #348:
' Line #349:
' 	Ld SLA_SELECTDIM 
' 	Ld SAA_OPTIONS 
' 	LitDI2 0x0003 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #350:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_GPM_SETPLOTATTR 
' 	LitDI2 0x000A 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #351:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_GPM_SETPLOTATTR 
' 	LitDI4 0x0001 0x0310 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #352:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_GPM_SETPLOTATTR 
' 	LitDI4 0x0402 0x00C0 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #353:
' 	Ld SAA_FROMVAL 
' 	Ld SAA_TOVAL 
' 	LitStr 0x0001 "0"
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #354:
' 	Ld SAA_FROMVAL 
' 	Ld GraphPages 
' 	Ld lMaxHistHeigh 
' 	Coerce (Str) 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #355:
' Line #356:
' 	Ld SLA_SELECTDIM 
' 	Ld SAA_OPTIONS 
' 	LitDI2 0x0003 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #357:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #358:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_COLOR 
' 	LitDI2 0x000A 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #359:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #360:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SEA_THICKNESS 
' 	LitHI4 0xFFFF 0x00FF 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #361:
' 	Ld SEA_COLORCOL 
' 	Ld GPM_SETAXISATTRSTRING 
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #362:
' 	Ld SEA_COLORCOL 
' 	Ld SLA_CONTOURFILLTYPE 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #363:
' 	Ld SLA_SELECTDIM 
' 	Ld SAA_SELECTLINE 
' 	LitDI2 0x0001 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #364:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0005 
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
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #369:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SAA_SUB1OPTIONS 
' 	LitDI2 0x0512 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #370:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SAA_SUB1OPTIONS 
' 	LitDI2 0x0F31 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #371:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0001 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #372:
' Line #373:
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd SPWNotebookComponentType 
' 	ArgsMemCall (Call) Add 0x0001 
' Line #374:
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #375:
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	LitDI2 0x0000 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #376:
' 	LitDI4 0x47FF 0x01E8 
' 	LitDI2 0x0002 
' 	LitDI2 0x0000 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #377:
' 	LitDI2 0x0001 
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #378:
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	LitDI2 0x0001 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #379:
' 	LitDI4 0x47FF 0x01E8 
' 	LitDI2 0x0002 
' 	LitDI2 0x0001 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #380:
' 	LitDI2 0x0003 
' 	LitDI2 0x0000 
' 	LitDI2 0x0002 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #381:
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	LitDI2 0x0002 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #382:
' 	LitDI4 0x47FF 0x01E8 
' 	LitDI2 0x0002 
' 	LitDI2 0x0002 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #383:
' 	LitDI2 0x0003 
' 	Ld xCount 
' 	LitDI2 0x0001 
' 	Sub 
' 	Paren 
' 	Add 
' 	LitDI2 0x0000 
' 	LitDI2 0x0003 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #384:
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	LitDI2 0x0003 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #385:
' 	LitDI4 0x47FF 0x01E8 
' 	LitDI2 0x0002 
' 	LitDI2 0x0003 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #386:
' Line #387:
' 	OptionBase 
' 	LitDI2 0x0000 
' 	Redim CurrentPageItem 0x0001 (As Variant)
' Line #388:
' Line #389:
' 	LitDI2 0x0004 
' 	LitDI2 0x0000 
' 	ArgsSt CurrentPageItem 0x0001 
' Line #390:
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
' Line #391:
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemLd Graphs 0x0001 
' 	ArgsMemLd SelectObject 0x0001 
' 	ArgsMemCall (Call) testSigmaPlot 0x0000 
' Line #392:
' Line #393:
' 	LitStr 0x0006 "Site y"
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemLd Graphs 0x0001 
' 	ArgsMemLd SelectObject 0x0001 
' 	MemSt Name 
' Line #394:
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
' Line #395:
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
' Line #396:
' Line #397:
' 	Ld SLA_SELECTDIM 
' 	Ld SAA_OPTIONS 
' 	LitDI2 0x0003 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #398:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #399:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_COLOR 
' 	LitDI2 0x000A 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #400:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #401:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SEA_THICKNESS 
' 	LitHI4 0xFFFF 0x00FF 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #402:
' 	Ld SEA_COLORCOL 
' 	Ld GPM_SETAXISATTRSTRING 
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #403:
' 	Ld SEA_COLORCOL 
' 	Ld SLA_CONTOURFILLTYPE 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #404:
' 	Ld SLA_SELECTDIM 
' 	Ld SAA_SELECTLINE 
' 	LitDI2 0x0001 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #405:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0005 
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
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #410:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SAA_SUB1OPTIONS 
' 	LitDI2 0x0512 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #411:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SAA_SUB1OPTIONS 
' 	LitDI2 0x0F31 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #412:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0001 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #413:
' Line #414:
' 	QuoteRem 0x0000 0x002D "            If i < UBound(theWorksheets) Then"
' Line #415:
' 	LitDI2 0x0001 
' 	Ld SPApplication 
' 	MemLd SPWNotebookComponentType 
' 	ArgsMemCall (Call) Add 0x0001 
' Line #416:
' 	QuoteRem 0x0000 0x0012 "            End If"
' Line #417:
' 	Ld Whi0l 
' 	Ld lMaxHistHe 
' 	Add 
' 	St Whi0l 
' Line #418:
' 	ElseBlock 
' Line #419:
' 	ExitDo 
' Line #420:
' 	EndIfBlock 
' Line #421:
' 	Loop 
' Line #422:
' 	EndSub 
' Line #423:
' Line #424:
' 	FuncDefn (Sub _B_var_lMaxHistHeight())
' Line #425:
' 	Dim (Const) 
' 	LitHI2 0x0406 
' 	VarDefn SAA_TOVAL
' Line #426:
' 	Dim (Const) 
' 	LitHI2 0x0407 
' 	VarDefn GraphPages
' Line #427:
' 	Dim (Const) 
' 	LitHI2 0x0301 
' 	VarDefn SLA_SELECTDIM
' Line #428:
' 	Dim (Const) 
' 	LitDI2 0x0401 
' 	VarDefn SEA_COLORCOL
' Line #429:
' 	Dim (Const) 
' 	LitDI2 0x0308 
' 	VarDefn SAA_OPTIONS
' Line #430:
' 	Dim (Const) 
' 	LitDI2 0x0403 
' 	VarDefn _B_var_GPM_SETPLOTATTR
' Line #431:
' 	Dim (Const) 
' 	LitDI2 0x0408 
' 	VarDefn SAA_FROMVAL
' Line #432:
' 	Dim (Const) 
' 	LitDI2 0x0615 
' 	VarDefn GPM_SETAXISATTRSTRING
' Line #433:
' 	Dim (Const) 
' 	LitDI2 0x0613 
' 	VarDefn SLA_CONTOURFILLTYPE
' Line #434:
' 	Dim (Const) 
' 	LitDI2 0x0358 
' 	VarDefn SAA_SELECTLINE
' Line #435:
' 	Dim (Const) 
' 	LitDI2 0x040A 
' 	VarDefn SEA_THICKNESS
' Line #436:
' 	Dim (Const) 
' 	LitDI2 0x0601 
' 	VarDefn SEA_COLOR
' Line #437:
' 	Dim (Const) 
' 	LitDI2 0x0606 
' 	VarDefn _B_var_SEA_THICKNESS
' Line #438:
' 	Dim (Const) 
' 	LitDI2 0x0410 
' 	VarDefn _B_var_SAA_SUB1OPTIONS
' Line #439:
' Line #440:
' 	Dim 
' 	VarDefn Module2 (As Object)
' Line #441:
' 	SetStmt 
' 	LitStr 0x0017 "SigmaPlot.Application.1"
' 	ArgsLd CreateObject 0x0001 
' 	Set Module2 
' Line #442:
' 	LitVarSpecial (True)
' 	Ld Module2 
' 	MemSt Application 
' Line #443:
' 	LitStr 0x0006 "Site y"
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemLd Graphs 0x0001 
' 	ArgsMemLd SelectObject 0x0001 
' 	MemSt Name 
' Line #444:
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
' Line #445:
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
' Line #446:
' Line #447:
' 	Ld SLA_SELECTDIM 
' 	Ld SAA_OPTIONS 
' 	LitDI2 0x0003 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #448:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_GPM_SETPLOTATTR 
' 	LitDI2 0x000A 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #449:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_GPM_SETPLOTATTR 
' 	LitDI4 0x0001 0x0310 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #450:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_GPM_SETPLOTATTR 
' 	LitDI4 0x0402 0x00C0 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #451:
' 	Ld SAA_FROMVAL 
' 	Ld SAA_TOVAL 
' 	LitStr 0x0001 "0"
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #452:
' 	Ld SAA_FROMVAL 
' 	Ld GraphPages 
' 	LitStr 0x0003 "150"
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #453:
' Line #454:
' 	Ld SLA_SELECTDIM 
' 	Ld SAA_OPTIONS 
' 	LitDI2 0x0003 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #455:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #456:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_COLOR 
' 	LitDI2 0x000A 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #457:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #458:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SEA_THICKNESS 
' 	LitHI4 0xFFFF 0x00FF 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #459:
' 	Ld SEA_COLORCOL 
' 	Ld GPM_SETAXISATTRSTRING 
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #460:
' 	Ld SEA_COLORCOL 
' 	Ld SLA_CONTOURFILLTYPE 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #461:
' 	Ld SLA_SELECTDIM 
' 	Ld SAA_SELECTLINE 
' 	LitDI2 0x0001 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #462:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0005 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #463:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SEA_THICKNESS 
' 	LitHI4 0xFFFF 0x00FF 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #464:
' 	Ld SEA_COLORCOL 
' 	Ld GPM_SETAXISATTRSTRING 
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #465:
' 	Ld SEA_COLORCOL 
' 	Ld SLA_CONTOURFILLTYPE 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #466:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #467:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SAA_SUB1OPTIONS 
' 	LitDI2 0x0512 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #468:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SAA_SUB1OPTIONS 
' 	LitDI2 0x0F31 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #469:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0001 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #470:
' Line #471:
' 	EndSub 
' Line #472:
' Line #473:
' 	FuncDefn (Function _B_var_buildEpocList(objTTX, iXAxisIndexAs, returnArr))
' Line #474:
' 	QuoteRem 0x0004 0x0030 "build list of epocs for the given axis epoc name"
' Line #475:
' Line #476:
' 	Dim 
' 	VarDefn AxisEp (As Dictionary)
' Line #477:
' 	SetStmt 
' 	New id_FFFF
' 	Set AxisEp 
' Line #478:
' Line #479:
' 	Dim 
' 	VarDefn dblStartTime (As Double)
' Line #480:
' 	Dim 
' 	VarDefn varReturn (As Variant)
' Line #481:
' Line #482:
' 	Dim 
' 	VarDefn i (As Integer)
' Line #483:
' 	Dim 
' 	VarDefn j (As Integer)
' Line #484:
' Line #485:
' 	Ld iXAxisIndexAs 
' 	LitStr 0x0007 "Channel"
' 	Eq 
' 	IfBlock 
' Line #486:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0001 
' 	LitDI2 0x0020 
' 	For 
' Line #487:
' 	Ld i 
' 	LitDI2 0x0000 
' 	Ld AxisEp 
' 	ArgsMemCall (Call) Add 0x0002 
' Line #488:
' 	StartForVariable 
' 	Next 
' Line #489:
' 	ElseBlock 
' Line #490:
' 	Do 
' Line #491:
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
' Line #492:
' 	Ld i 
' 	LitDI2 0x0000 
' 	Eq 
' 	IfBlock 
' Line #493:
' 	ExitDo 
' Line #494:
' 	EndIfBlock 
' Line #495:
' Line #496:
' 	LitDI2 0x0000 
' 	Ld i 
' 	LitDI2 0x0000 
' 	Ld objTTX 
' 	ArgsMemLd ParseEvInfoV 0x0003 
' 	St varReturn 
' Line #497:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld i 
' 	LitDI2 0x0001 
' 	Sub 
' 	Paren 
' 	For 
' Line #498:
' 	LitDI2 0x0006 
' 	Ld j 
' 	ArgsLd varReturn 0x0002 
' 	Ld AxisEp 
' 	ArgsMemLd Exists 0x0001 
' 	Not 
' 	IfBlock 
' Line #499:
' 	LitDI2 0x0006 
' 	Ld j 
' 	ArgsLd varReturn 0x0002 
' 	LitStr 0x0000 ""
' 	Ld AxisEp 
' 	ArgsMemCall (Call) Add 0x0002 
' Line #500:
' 	EndIfBlock 
' Line #501:
' 	LitDI2 0x0005 
' 	Ld j 
' 	ArgsLd varReturn 0x0002 
' 	LitDI2 0x0001 
' 	LitDI4 0x86A0 0x0001 
' 	Div 
' 	Paren 
' 	Add 
' 	St dblStartTime 
' Line #502:
' 	StartForVariable 
' 	Next 
' Line #503:
' Line #504:
' 	Ld i 
' 	LitDI2 0x01F4 
' 	Lt 
' 	IfBlock 
' Line #505:
' 	ExitDo 
' Line #506:
' 	EndIfBlock 
' Line #507:
' 	Loop 
' Line #508:
' 	EndIfBlock 
' Line #509:
' Line #510:
' Line #511:
' Line #512:
' 	Ld returnArr 
' 	IfBlock 
' Line #513:
' 	Dim 
' 	VarDefn _B_var_returnArr
' Line #514:
' 	Dim 
' 	VarDefn zOffsets (As Variant)
' Line #515:
' 	Ld AxisEp 
' 	MemLd Keys 
' 	St zOffsets 
' Line #516:
' 	OptionBase 
' 	Ld zOffsets 
' 	FnUBound 0x0000 
' 	Redim _B_var_returnArr 0x0001 (As Variant)
' Line #517:
' Line #518:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld zOffsets 
' 	FnUBound 0x0000 
' 	For 
' Line #519:
' 	Ld zOffsets 
' 	FnUBound 0x0000 
' 	Ld i 
' 	Sub 
' 	ArgsLd zOffsets 0x0001 
' 	Ld i 
' 	ArgsSt _B_var_returnArr 0x0001 
' Line #520:
' 	StartForVariable 
' 	Next 
' Line #521:
' 	Ld _B_var_returnArr 
' 	St _B_var_buildEpocList 
' Line #522:
' 	ElseBlock 
' Line #523:
' 	Ld AxisEp 
' 	MemLd Keys 
' 	St _B_var_buildEpocList 
' Line #524:
' 	EndIfBlock 
' Line #525:
' Line #526:
' 	EndFunc 
' Line #527:
' Line #528:
' Line #529:
' 	FuncDefn (Function _B_var_processSearch(ByRef objTTX, ByRef _B_var_arrOtherEp, ByRef strSearchString, iOtherEpocIndex, xAxisSearchString As String, yOffset, zOffset, xOffes, 1, Le, ByRef yCoun, ByRef xCount, ByRef lMaxHistHe, ByRef lMaxHistHeigh))
' Line #530:
' 	Dim 
' 	VarDefn i (As Integer)
' Line #531:
' 	Dim 
' 	VarDefn j (As Integer)
' Line #532:
' 	Dim 
' 	VarDefn _B_var_objTTX (As String)
' Line #533:
' 	Dim 
' 	VarDefn strHeading (As String)
' Line #534:
' 	Dim 
' 	VarDefn Label1 (As String)
' Line #535:
' Line #536:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld iOtherEpocIndex 
' 	ArgsLd strSearchString 0x0001 
' 	FnUBound 0x0000 
' 	For 
' Line #537:
' 	Ld iOtherEpocIndex 
' 	ArgsLd _B_var_arrOtherEp 0x0001 
' 	LitStr 0x0007 "Channel"
' 	Ne 
' 	IfBlock 
' Line #538:
' 	QuoteRem 0x000C 0x0014 "add to search string"
' Line #539:
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
' Line #540:
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
' Line #541:
' 	ElseBlock 
' Line #542:
' 	Ld xAxisSearchString 
' 	St _B_var_objTTX 
' Line #543:
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
' Line #544:
' 	Ld i 
' 	Ld iOtherEpocIndex 
' 	ArgsLd strSearchString 0x0001 
' 	IndexLd 0x0001 
' 	St 1 
' Line #545:
' 	EndIfBlock 
' Line #546:
' 	Ld iOtherEpocIndex 
' 	Ld _B_var_arrOtherEp 
' 	FnUBound 0x0000 
' 	Lt 
' 	IfBlock 
' Line #547:
' 	QuoteRem 0x000C 0x002F "there are still more epocs to add to the search"
' Line #548:
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
' Line #549:
' 	ElseBlock 
' Line #550:
' 	QuoteRem 0x000C 0x004B "we have reached the end of the list of epocs - can actually do a search now"
' Line #551:
' 	Ld _B_var_objTTX 
' 	LitDI2 0x0005 
' 	ArgsLd Right 0x0002 
' 	LitStr 0x0005 " and "
' 	Eq 
' 	IfBlock 
' 	QuoteRem 0x003D 0x0045 "this should always be the case - should be a trailing 'and' to remove"
' Line #552:
' 	Ld _B_var_objTTX 
' 	Ld _B_var_objTTX 
' 	FnLen 
' 	LitDI2 0x0005 
' 	Sub 
' 	ArgsLd Left 0x0002 
' 	St strHeading 
' Line #553:
' 	ElseBlock 
' Line #554:
' 	Ld _B_var_objTTX 
' 	St strHeading 
' Line #555:
' 	EndIfBlock 
' Line #556:
' 	Ld strHeading 
' 	Ld objTTX 
' 	ArgsMemCall (Call) SetFilterWithDescEx 0x0001 
' Line #557:
' Line #558:
' 	Ld yOffset 
' 	LitDI2 0x0001 
' 	Eq 
' 	Ld zOffset 
' 	LitDI2 0x0001 
' 	Eq 
' 	And 
' 	IfBlock 
' Line #559:
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
' Line #560:
' 	Ld vYAxisKeys 
' 	Ld buildOptionLists 
' 	Ld _B_var_iColOffset 
' 	Ld _B_var_Const 
' 	Ld i 
' 	Ld xOffes 
' 	Mul 
' 	Paren 
' 	ArgsCall (Call) SubwriteAxes 0x0005 
' Line #561:
' 	EndIfBlock 
' Line #562:
' Line #563:
' 	Ld objTTX 
' 	Ld yOffset 
' 	Ld zOffset 
' 	Ld i 
' 	Ld xOffes 
' 	Mul 
' 	Ld 1 
' 	Ld lMaxHistHeigh 
' 	ArgsCall (Call) _B_var_writeResults 0x0006 
' Line #564:
' 	Ld yOffset 
' 	Ld yCoun 
' 	Gt 
' 	IfBlock 
' Line #565:
' 	Ld yOffset 
' 	St yCoun 
' Line #566:
' 	EndIfBlock 
' Line #567:
' 	Ld zOffset 
' 	Ld xCount 
' 	Gt 
' 	IfBlock 
' Line #568:
' 	Ld zOffset 
' 	St xCount 
' Line #569:
' 	EndIfBlock 
' Line #570:
' 	Ld xOffes 
' 	St lMaxHistHe 
' Line #571:
' 	EndIfBlock 
' Line #572:
' 	StartForVariable 
' 	Next 
' Line #573:
' Line #574:
' 	EndFunc 
' Line #575:
' Line #576:
' 	FuncDefn (Sub _B_var_writeResults(ByRef objTTX, yOffset, zOffset, xOffes, 1, ByRef lMaxHistHeigh))
' Line #577:
' 	Dim 
' 	VarDefn varReturn (As Variant)
' Line #578:
' 	Dim 
' 	VarDefn varChanData (As Variant)
' Line #579:
' Line #580:
' 	Dim 
' 	VarDefn dblStartTime (As Double)
' Line #581:
' 	Dim 
' 	VarDefn dblEndTime (As Double)
' Line #582:
' 	Dim 
' 	VarDefn IsEmpty (As Double)
' Line #583:
' Line #584:
' 	Dim 
' 	VarDefn i (As Long)
' Line #585:
' 	Dim 
' 	VarDefn j (As Long)
' Line #586:
' 	Dim 
' 	VarDefn k (As Long)
' Line #587:
' Line #588:
' 	Dim 
' 	VarDefn histTmp (As Long)
' Line #589:
' Line #590:
' 	LitStr 0x0004 "Swep"
' 	LitDI2 0x0000 
' 	Ld objTTX 
' 	ArgsMemLd GetEpocsExV 0x0002 
' 	St varReturn 
' Line #591:
' 	Ld varReturn 
' 	ArgsLd Dib 0x0001 
' 	IfBlock 
' Line #592:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld varReturn 
' 	LitDI2 0x0002 
' 	FnUBound 0x0001 
' 	For 
' Line #593:
' 	LitDI2 0x0002 
' 	Ld i 
' 	ArgsLd varReturn 0x0002 
' 	Ld _B_var_lIgnoreFirstMsec 
' 	Add 
' 	St dblStartTime 
' Line #594:
' 	Ld dblStartTime 
' 	Ld lBinWidth 
' 	Add 
' 	Ld _B_var_lIgnoreFirstMsec 
' 	Add 
' 	St dblEndTime 
' Line #595:
' 	Ld dblStartTime 
' 	St IsEmpty 
' Line #596:
' 	Do 
' Line #597:
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
' Line #598:
' 	Ld k 
' 	LitDI2 0x0000 
' 	Eq 
' 	IfBlock 
' Line #599:
' 	ExitDo 
' Line #600:
' 	EndIfBlock 
' Line #601:
' Line #602:
' 	Ld histTmp 
' 	Coerce (Lng) 
' 	Ld k 
' 	Coerce (Lng) 
' 	Add 
' 	St histTmp 
' Line #603:
' 	Ld k 
' 	LitDI2 0x01F4 
' 	Lt 
' 	IfBlock 
' Line #604:
' 	ExitDo 
' Line #605:
' 	ElseBlock 
' Line #606:
' 	Ld k 
' 	LitDI2 0x0001 
' 	Sub 
' 	LitDI2 0x0001 
' 	LitDI2 0x0006 
' 	Ld objTTX 
' 	ArgsMemLd ParseEvInfoV 0x0003 
' 	St varChanData 
' Line #607:
' 	LitDI2 0x0000 
' 	ArgsLd varChanData 0x0001 
' 	LitDI2 0x0001 
' 	LitDI4 0x86A0 0x0001 
' 	Div 
' 	Paren 
' 	Add 
' 	St dblStartTime 
' Line #608:
' 	EndIfBlock 
' Line #609:
' 	Loop 
' Line #610:
' 	Ld IsEmpty 
' 	St dblStartTime 
' Line #611:
' 	StartForVariable 
' 	Next 
' Line #612:
' Line #613:
' 	Ld yAxisEp 
' 	LitStr 0x0007 "Channel"
' 	Eq 
' 	IfBlock 
' Line #614:
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
' Line #615:
' 	Ld otherEp 
' 	LitStr 0x0007 "Channel"
' 	Eq 
' 	ElseIfBlock 
' Line #616:
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
' Line #617:
' 	ElseBlock 
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
' 	EndIfBlock 
' Line #620:
' 	Ld histTmp 
' 	Ld lMaxHistHeigh 
' 	Gt 
' 	IfBlock 
' Line #621:
' 	Ld histTmp 
' 	St lMaxHistHeigh 
' Line #622:
' 	EndIfBlock 
' Line #623:
' 	LitDI2 0x0000 
' 	St histTmp 
' Line #624:
' 	EndIfBlock 
' Line #625:
' Line #626:
' 	EndSub 
' _VBA_PROJECT_CUR/VBA/Sheet2 - 1166 bytes
' _VBA_PROJECT_CUR/VBA/Sheet3 - 1150 bytes
' _VBA_PROJECT_CUR/VBA/Sheet4 - 1166 bytes
