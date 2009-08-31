' Processing file: Tuning curve generator - generic.xls
' ===============================================================================
' Module streams:
' _VBA_PROJECT_CUR/VBA/ThisWorkbook - 1210 bytes
' Line #0:
' 	Option  (Explicit)
' Line #1:
' _VBA_PROJECT_CUR/VBA/Sheet1 - 1150 bytes
' _VBA_PROJECT_CUR/VBA/ImportFrom - 22380 bytes
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
' 	Ld UseTakn 
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
' _VBA_PROJECT_CUR/VBA/Module1 - 42311 bytes
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
' 	Ld ImportFrom 
' 	ArgsMemCall Show 0x0000 
' Line #24:
' Line #25:
' 	Ld doImport 
' 	IfBlock 
' Line #26:
' 	LitVarSpecial (True)
' 	ArgsCall (Call) processImport 0x0001 
' Line #27:
' 	EndIfBlock 
' Line #28:
' 	EndSub 
' Line #29:
' Line #30:
' 	FuncDefn (Sub processImport(spNB As Boolean))
' Line #31:
' Line #32:
' 	QuoteRem 0x0004 0x002B "load the bin width for histogram generation"
' Line #33:
' 	LitStr 0x0002 "B1"
' 	LitStr 0x0008 "Settings"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	St lBinWidth 
' Line #34:
' Line #35:
' 	QuoteRem 0x0004 0x004D "load the # of msec to ignore at the start (for filtering stimulation artifact"
' Line #36:
' 	LitStr 0x0002 "B2"
' 	LitStr 0x0008 "Settings"
' 	ArgsLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	St _B_var_lIgnoreFirstMsec 
' Line #37:
' Line #38:
' 	QuoteRem 0x0004 0x003A "used to store the maximum histogram peak for normalisation"
' Line #39:
' 	Dim 
' 	VarDefn lMaxHistHeigh (As Double)
' Line #40:
' 	LitDI2 0x0000 
' 	St lMaxHistHeigh 
' Line #41:
' Line #42:
' 	Dim 
' 	VarDefn theWorksheets (As Variant)
' 	QuoteRem 0x0021 0x0029 "stores the created worksheets to write to"
' Line #43:
' 	Dim 
' 	VarDefn _B_var_arrHistTmp (As Long)
' 	QuoteRem 0x001D 0x0044 "used to store the histogram data for each channel as it is generated"
' Line #44:
' 	OptionBase 
' 	LitDI2 0x001F 
' 	Redim _B_var_arrHistTmp 0x0001 (As Variant)
' Line #45:
' Line #46:
' 	QuoteRem 0x0004 0x0037 "offsets to leave space at the top and left of the chart"
' Line #47:
' 	LitDI2 0x0001 
' 	St _B_var_Const 
' Line #48:
' 	LitDI2 0x0000 
' 	St _B_var_iColOffset 
' Line #49:
' Line #50:
' 	QuoteRem 0x0000 0x0050 "    theWorksheets = buildWorksheetArray() 'build the worksheets for writing data"
' Line #51:
' Line #52:
' 	QuoteRem 0x0004 0x0013 "connect to the tank"
' Line #53:
' 	Dim 
' 	VarDefn objTTX
' Line #54:
' 	SetStmt 
' 	LitStr 0x0007 "TTank.X"
' 	ArgsLd CreateObject 0x0001 
' 	Set objTTX 
' Line #55:
' Line #56:
' 	Ld theServer 
' 	LitStr 0x0002 "Me"
' 	Ld objTTX 
' 	ArgsMemLd ConnectServer 0x0002 
' 	LitDI2 0x0001 
' 	Coerce (Lng) 
' 	Ne 
' 	IfBlock 
' Line #57:
' 	LitStr 0x0015 "Connecting to server "
' 	Ld theServer 
' 	Concat 
' 	LitStr 0x0008 " failed."
' 	Concat 
' 	Paren 
' 	ArgsCall MsgBox 0x0001 
' Line #58:
' 	ExitSub 
' Line #59:
' 	EndIfBlock 
' Line #60:
' Line #61:
' 	Ld theTank 
' 	LitStr 0x0001 "R"
' 	Ld objTTX 
' 	ArgsMemLd OpenTank 0x0002 
' 	LitDI2 0x0001 
' 	Coerce (Lng) 
' 	Ne 
' 	IfBlock 
' Line #62:
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
' Line #63:
' 	Ld objTTX 
' 	ArgsMemCall (Call) ReleaseServer 0x0000 
' Line #64:
' 	ExitSub 
' Line #65:
' 	EndIfBlock 
' Line #66:
' Line #67:
' 	Ld theBlock 
' 	Ld objTTX 
' 	ArgsMemLd SelectBlock 0x0001 
' 	LitDI2 0x0001 
' 	Coerce (Lng) 
' 	Ne 
' 	IfBlock 
' Line #68:
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
' Line #69:
' 	Ld objTTX 
' 	ArgsMemCall (Call) CloseTank 0x0000 
' Line #70:
' 	Ld objTTX 
' 	ArgsMemCall (Call) ReleaseServer 0x0000 
' Line #71:
' 	ExitSub 
' Line #72:
' 	EndIfBlock 
' Line #73:
' Line #74:
' 	QuoteRem 0x0004 0x0026 "index epochs - required to use filters"
' Line #75:
' 	Ld objTTX 
' 	ArgsMemCall (Call) CreateEpocIndexing 0x0000 
' Line #76:
' Line #77:
' 	Dim 
' 	VarDefn dblStartTime (As Double)
' Line #78:
' 	Dim 
' 	VarDefn dblEndTime (As Double)
' Line #79:
' Line #80:
' 	Dim 
' 	VarDefn varReturn (As Variant)
' Line #81:
' Line #82:
' 	QuoteRem 0x0000 0x001D "    Dim vXAxisKeys As Variant"
' Line #83:
' 	QuoteRem 0x0000 0x001D "    Dim vYAxisKeys As Variant"
' Line #84:
' Line #85:
' 	Ld objTTX 
' 	Ld yAxisEp 
' 	Ld bReverseY 
' 	ArgsLd _B_var_buildEpocList 0x0003 
' 	St vYAxisKeys 
' Line #86:
' 	Ld objTTX 
' 	Ld otherEp 
' 	Ld ReverseX 
' 	ArgsLd _B_var_buildEpocList 0x0003 
' 	St buildOptionLists 
' Line #87:
' Line #88:
' 	Dim 
' 	VarDefn i (As Long)
' Line #89:
' 	Dim 
' 	VarDefn j (As Long)
' Line #90:
' 	Dim 
' 	VarDefn k (As Long)
' Line #91:
' 	Dim 
' 	VarDefn l (As Long)
' Line #92:
' Line #93:
' 	Dim 
' 	VarDefn strSearchString (As Variant)
' Line #94:
' 	Ld _B_var_arrOtherEp 
' 	FnUBound 0x0000 
' 	LitDI2 0x0001 
' 	UMi 
' 	Ne 
' 	IfBlock 
' Line #95:
' 	OptionBase 
' 	Ld _B_var_arrOtherEp 
' 	FnUBound 0x0000 
' 	Redim strSearchString 0x0001 (As Variant)
' Line #96:
' Line #97:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld _B_var_arrOtherEp 
' 	FnUBound 0x0000 
' 	For 
' Line #98:
' 	Ld objTTX 
' 	Ld i 
' 	ArgsLd _B_var_arrOtherEp 0x0001 
' 	LitVarSpecial (False)
' 	ArgsLd _B_var_buildEpocList 0x0003 
' 	Ld i 
' 	ArgsSt strSearchString 0x0001 
' Line #99:
' 	StartForVariable 
' 	Next 
' Line #100:
' 	EndIfBlock 
' Line #101:
' Line #102:
' 	LitDI2 0x0000 
' 	St i 
' Line #103:
' 	LitDI2 0x0000 
' 	St j 
' Line #104:
' Line #105:
' 	Dim 
' 	VarDefn iYAxisIndex (As Integer)
' Line #106:
' 	Dim 
' 	VarDefn otherEpocList (As Integer)
' Line #107:
' 	Dim 
' 	VarDefn _B_var_arrOtherEpocIndex (As Integer)
' Line #108:
' 	Ld _B_var_arrOtherEp 
' 	FnUBound 0x0000 
' 	LitDI2 0x0001 
' 	UMi 
' 	Ne 
' 	IfBlock 
' Line #109:
' 	OptionBase 
' 	Ld _B_var_arrOtherEp 
' 	FnUBound 0x0000 
' 	Redim _B_var_arrOtherEpocIndex 0x0001 (As Variant)
' Line #110:
' 	EndIfBlock 
' Line #111:
' Line #112:
' 	Dim 
' 	VarDefn varChanData (As Variant)
' Line #113:
' 	Dim 
' 	VarDefn IsEmpty (As Double)
' Line #114:
' Line #115:
' 	Dim 
' 	VarDefn yAxisSearchString (As String)
' Line #116:
' 	Dim 
' 	VarDefn otherAxisSearchString (As String)
' Line #117:
' 	Dim 
' 	VarDefn processSearch (As String)
' Line #118:
' 	Dim 
' 	VarDefn arrOtherEpFor (As String)
' Line #119:
' 	Ld _B_var_arrOtherEp 
' 	FnUBound 0x0000 
' 	LitDI2 0x0001 
' 	UMi 
' 	Ne 
' 	IfBlock 
' Line #120:
' 	OptionBase 
' 	Ld _B_var_arrOtherEp 
' 	FnUBound 0x0000 
' 	Redim processSearch 0x0001 (As Variant)
' Line #121:
' 	EndIfBlock 
' Line #122:
' Line #123:
' 	Dim 
' 	VarDefn 1 (As Integer)
' Line #124:
' 	LitDI2 0x0000 
' 	St 1 
' Line #125:
' Line #126:
' 	Ld _B_var_arrOtherEp 
' 	FnUBound 0x0000 
' 	LitDI2 0x0001 
' 	UMi 
' 	Ne 
' 	IfBlock 
' Line #127:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld vYAxisKeys 
' 	FnUBound 0x0000 
' 	For 
' Line #128:
' 	Ld yAxisEp 
' 	LitStr 0x0007 "Channel"
' 	Eq 
' 	IfBlock 
' Line #129:
' 	Ld i 
' 	ArgsLd vYAxisKeys 0x0001 
' 	St 1 
' Line #130:
' 	LitStr 0x0000 ""
' 	St yAxisSearchString 
' Line #131:
' 	ElseBlock 
' Line #132:
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
' Line #133:
' 	EndIfBlock 
' Line #134:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld buildOptionLists 
' 	FnUBound 0x0000 
' 	For 
' Line #135:
' 	Ld otherEp 
' 	LitStr 0x0007 "Channel"
' 	Eq 
' 	IfBlock 
' Line #136:
' 	Ld j 
' 	ArgsLd buildOptionLists 0x0001 
' 	St 1 
' Line #137:
' 	LitStr 0x0000 ""
' 	St otherAxisSearchString 
' Line #138:
' 	ElseBlock 
' Line #139:
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
' Line #140:
' 	EndIfBlock 
' Line #141:
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
' 	ArgsCall (Call) _B_var_processSearch 0x000A 
' Line #142:
' 	StartForVariable 
' 	Next 
' Line #143:
' 	StartForVariable 
' 	Next 
' Line #144:
' 	EndIfBlock 
' Line #145:
' Line #146:
' 	QuoteRem 0x0000 0x0051 "    Call writeAxes(theWorksheets, vXAxisKeys, vYAxisKeys, iColOffset, iRowOffset)"
' Line #147:
' Line #148:
' 	Ld objTTX 
' 	ArgsMemCall (Call) CloseTank 0x0000 
' Line #149:
' 	Ld objTTX 
' 	ArgsMemCall (Call) ReleaseServer 0x0000 
' Line #150:
' Line #151:
' 	Ld spNB 
' 	IfBlock 
' Line #152:
' 	Ld theWorksheets 
' 	Ld vYAxisKeys 
' 	Ld buildOptionLists 
' 	Ld _B_var_iColOffset 
' 	Ld _B_var_Const 
' 	Ld lMaxHistHeigh 
' 	ArgsCall (Call) ACTIVESPWLib 0x0006 
' Line #153:
' 	EndIfBlock 
' Line #154:
' Line #155:
' 	EndSub 
' Line #156:
' Line #157:
' 	FuncDefn (Function buildWorksheetArray() As Variant)
' Line #158:
' 	Dim 
' 	OptionBase 
' 	LitDI2 0x001F 
' 	VarDefn theWorksheets
' Line #159:
' 	Dim 
' 	VarDefn strWsname (As String)
' Line #160:
' 	Dim 
' 	VarDefn intWSNum (As Long)
' Line #161:
' Line #162:
' 	Dim 
' 	VarDefn i (As Integer)
' Line #163:
' Line #164:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0001 
' 	Ld Worksheets 
' 	MemLd Count 
' 	For 
' Line #165:
' 	Ld i 
' 	Ld Worksheets 
' 	ArgsMemLd Item 0x0001 
' 	MemLd Name 
' 	St strWsname 
' Line #166:
' 	Ld strWsname 
' 	LitDI2 0x0004 
' 	ArgsLd Left 0x0002 
' 	LitStr 0x0004 "Site"
' 	Eq 
' 	IfBlock 
' Line #167:
' 	Ld strWsname 
' 	Ld strWsname 
' 	FnLen 
' 	LitDI2 0x0004 
' 	Sub 
' 	ArgsLd Right 0x0002 
' 	ArgsLd IsNumeric 0x0001 
' 	IfBlock 
' Line #168:
' 	Ld strWsname 
' 	Ld strWsname 
' 	FnLen 
' 	LitDI2 0x0004 
' 	Sub 
' 	ArgsLd Right 0x0002 
' 	Coerce (Int) 
' 	St intWSNum 
' Line #169:
' 	Ld intWSNum 
' 	LitDI2 0x0021 
' 	Lt 
' 	Ld intWSNum 
' 	LitDI2 0x0000 
' 	Gt 
' 	And 
' 	IfBlock 
' Line #170:
' 	SetStmt 
' 	Ld i 
' 	Ld Worksheets 
' 	ArgsMemLd Item 0x0001 
' 	Ld intWSNum 
' 	LitDI2 0x0001 
' 	Sub 
' 	ArgsSet theWorksheets 0x0001 
' Line #171:
' 	EndIfBlock 
' Line #172:
' 	EndIfBlock 
' Line #173:
' 	EndIfBlock 
' Line #174:
' 	StartForVariable 
' 	Next 
' Line #175:
' Line #176:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	LitDI2 0x001F 
' 	For 
' Line #177:
' 	Ld i 
' 	ArgsLd theWorksheets 0x0001 
' 	ArgsLd Sheet70 0x0001 
' 	IfBlock 
' Line #178:
' 	Ld i 
' 	LitDI2 0x0000 
' 	Gt 
' 	IfBlock 
' Line #179:
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
' Line #180:
' 	ElseBlock 
' Line #181:
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
' Line #182:
' 	EndIfBlock 
' Line #183:
' 	LitStr 0x0004 "Site"
' 	Ld i 
' 	LitDI2 0x0001 
' 	Add 
' 	Coerce (Str) 
' 	Concat 
' 	Ld i 
' 	ArgsLd theWorksheets 0x0001 
' 	MemSt Name 
' Line #184:
' 	EndIfBlock 
' Line #185:
' 	StartForVariable 
' 	Next 
' Line #186:
' 	Ld theWorksheets 
' 	St buildWorksheetArray 
' Line #187:
' 	EndFunc 
' Line #188:
' Line #189:
' 	FuncDefn (Sub SubwriteAxes(rowLabels As Variant, deleteWorksheets As Variant, _B_var_iColOffset, _B_var_Const, xOffes))
' Line #190:
' 	Dim 
' 	VarDefn j (As Long)
' Line #191:
' Line #192:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld deleteWorksheets 
' 	FnUBound 0x0000 
' 	For 
' Line #193:
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
' Line #194:
' 	StartForVariable 
' 	Next 
' Line #195:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld rowLabels 
' 	FnUBound 0x0000 
' 	For 
' Line #196:
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
' Line #197:
' 	StartForVariable 
' 	Next 
' Line #198:
' Line #199:
' 	EndSub 
' Line #200:
' Line #201:
' 	FuncDefn (Sub Delete())
' Line #202:
' 	Dim 
' 	VarDefn strWsname (As String)
' Line #203:
' 	Dim 
' 	VarDefn intWSNum (As Long)
' Line #204:
' Line #205:
' 	Dim 
' 	VarDefn i (As Integer)
' Line #206:
' Line #207:
' 	Ld Worksheets 
' 	MemLd Count 
' 	St i 
' Line #208:
' Line #209:
' 	Do 
' Line #210:
' 	Ld i 
' 	LitDI2 0x0000 
' 	Eq 
' 	IfBlock 
' Line #211:
' 	ExitDo 
' Line #212:
' 	EndIfBlock 
' Line #213:
' 	Ld i 
' 	Ld Worksheets 
' 	ArgsMemLd Item 0x0001 
' 	MemLd Name 
' 	St strWsname 
' Line #214:
' 	Ld strWsname 
' 	LitDI2 0x0004 
' 	ArgsLd Left 0x0002 
' 	LitStr 0x0004 "Site"
' 	Eq 
' 	IfBlock 
' Line #215:
' 	Ld strWsname 
' 	Ld strWsname 
' 	FnLen 
' 	LitDI2 0x0004 
' 	Sub 
' 	ArgsLd Right 0x0002 
' 	ArgsLd IsNumeric 0x0001 
' 	IfBlock 
' Line #216:
' 	Ld strWsname 
' 	Ld strWsname 
' 	FnLen 
' 	LitDI2 0x0004 
' 	Sub 
' 	ArgsLd Right 0x0002 
' 	Coerce (Int) 
' 	St intWSNum 
' Line #217:
' 	Ld intWSNum 
' 	LitDI2 0x0021 
' 	Lt 
' 	Ld intWSNum 
' 	LitDI2 0x0000 
' 	Gt 
' 	And 
' 	IfBlock 
' Line #218:
' 	Ld i 
' 	Ld Worksheets 
' 	ArgsMemLd Item 0x0001 
' 	ArgsMemCall UserForm1 0x0000 
' Line #219:
' 	EndIfBlock 
' Line #220:
' 	EndIfBlock 
' Line #221:
' 	EndIfBlock 
' Line #222:
' 	Ld i 
' 	LitDI2 0x0001 
' 	Sub 
' 	St i 
' Line #223:
' 	Loop 
' Line #224:
' 	EndSub 
' Line #225:
' Line #226:
' 	FuncDefn (Sub ACTIVESPWLib(theWorksheets As Variant, rowLabels As Variant, deleteWorksheets As Variant, _B_var_iColOffset, _B_var_Const, lMaxHistHeigh))
' Line #227:
' Line #228:
' 	Dim (Const) 
' 	LitHI2 0x0406 
' 	VarDefn SAA_TOVAL
' Line #229:
' 	Dim (Const) 
' 	LitHI2 0x0407 
' 	VarDefn GraphPages
' Line #230:
' 	Dim (Const) 
' 	LitHI2 0x0301 
' 	VarDefn SLA_SELECTDIM
' Line #231:
' 	Dim (Const) 
' 	LitDI2 0x0401 
' 	VarDefn SEA_COLORCOL
' Line #232:
' 	Dim (Const) 
' 	LitDI2 0x0308 
' 	VarDefn SAA_OPTIONS
' Line #233:
' 	Dim (Const) 
' 	LitDI2 0x0403 
' 	VarDefn _B_var_GPM_SETPLOTATTR
' Line #234:
' 	Dim (Const) 
' 	LitDI2 0x0408 
' 	VarDefn SAA_FROMVAL
' Line #235:
' 	Dim (Const) 
' 	LitDI2 0x0615 
' 	VarDefn GPM_SETAXISATTRSTRING
' Line #236:
' 	Dim (Const) 
' 	LitDI2 0x0613 
' 	VarDefn SLA_CONTOURFILLTYPE
' Line #237:
' 	Dim (Const) 
' 	LitDI2 0x0358 
' 	VarDefn SAA_SELECTLINE
' Line #238:
' 	Dim (Const) 
' 	LitDI2 0x040A 
' 	VarDefn SEA_THICKNESS
' Line #239:
' 	Dim (Const) 
' 	LitDI2 0x0601 
' 	VarDefn SEA_COLOR
' Line #240:
' 	Dim (Const) 
' 	LitDI2 0x0606 
' 	VarDefn _B_var_SEA_THICKNESS
' Line #241:
' 	Dim (Const) 
' 	LitDI2 0x0410 
' 	VarDefn _B_var_SAA_SUB1OPTIONS
' Line #242:
' Line #243:
' 	Dim 
' 	VarDefn Module2 (As Object)
' Line #244:
' 	SetStmt 
' 	LitStr 0x0017 "SigmaPlot.Application.1"
' 	ArgsLd CreateObject 0x0001 
' 	Set Module2 
' Line #245:
' 	LitVarSpecial (True)
' 	Ld Module2 
' 	MemSt Application 
' Line #246:
' 	Ld Module2 
' 	MemLd Notebooks 
' 	MemLd buildTuningCurvesIntoSigmaplot 
' 	ArgsMemCall (Call) Add 0x0000 
' Line #247:
' Line #248:
' 	Dim 
' 	VarDefn i (As Integer)
' Line #249:
' 	Dim 
' 	VarDefn j (As Long)
' Line #250:
' 	Dim 
' 	VarDefn k (As Long)
' Line #251:
' Line #252:
' 	Dim 
' 	VarDefn SPApplication (As Object)
' Line #253:
' 	Dim 
' 	VarDefn spDT (As Object)
' Line #254:
' 	Dim 
' 	VarDefn DataTable (As Object)
' Line #255:
' 	Dim 
' 	VarDefn objSPWizard (As Object)
' Line #256:
' Line #257:
' Line #258:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld theWorksheets 
' 	FnUBound 0x0000 
' 	For 
' Line #259:
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
' Line #260:
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
' Line #261:
' 	Ld i 
' 	ArgsLd theWorksheets 0x0001 
' 	MemLd Name 
' 	Ld spDT 
' 	MemSt Name 
' Line #262:
' 	SetStmt 
' 	Ld spDT 
' 	MemLd Cell 
' 	Set DataTable 
' Line #263:
' Line #264:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld deleteWorksheets 
' 	FnUBound 0x0000 
' 	For 
' Line #265:
' 	Ld j 
' 	ArgsLd deleteWorksheets 0x0001 
' 	LitDI2 0x0001 
' 	Ld j 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #266:
' 	StartForVariable 
' 	Next 
' Line #267:
' Line #268:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld rowLabels 
' 	FnUBound 0x0000 
' 	For 
' Line #269:
' 	Ld j 
' 	ArgsLd rowLabels 0x0001 
' 	LitDI2 0x0000 
' 	Ld j 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #270:
' 	StartForVariable 
' 	Next 
' Line #271:
' Line #272:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld rowLabels 
' 	FnUBound 0x0000 
' 	For 
' Line #273:
' 	StartForVariable 
' 	Ld k 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld deleteWorksheets 
' 	FnUBound 0x0000 
' 	For 
' Line #274:
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
' Line #275:
' 	StartForVariable 
' 	Next 
' Line #276:
' 	StartForVariable 
' 	Next 
' Line #277:
' Line #278:
' 	LitStr 0x0011 "@rgb(255,255,255)"
' 	LitDI2 0x0002 
' 	LitDI2 0x0000 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #279:
' 	LitStr 0x000D "@rgb(0,0,255)"
' 	LitDI2 0x0002 
' 	LitDI2 0x0001 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #280:
' 	LitStr 0x000F "@rgb(0,255,255)"
' 	LitDI2 0x0002 
' 	LitDI2 0x0002 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #281:
' 	LitStr 0x000D "@rgb(0,255,0)"
' 	LitDI2 0x0002 
' 	LitDI2 0x0003 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #282:
' 	LitStr 0x000F "@rgb(255,255,0)"
' 	LitDI2 0x0002 
' 	LitDI2 0x0004 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #283:
' 	LitStr 0x000D "@rgb(255,0,0)"
' 	LitDI2 0x0002 
' 	LitDI2 0x0005 
' 	Ld DataTable 
' 	ArgsMemSt iRowOffset 0x0002 
' Line #284:
' Line #285:
' 	QuoteRem 0x0008 0x001E "Call spNB.NotebookItems.Add(2)"
' Line #286:
' 	QuoteRem 0x0008 0x0042 "Set spGRPH = spNB.NotebookItems.Item(spNB.NotebookItems.Count - 1)"
' Line #287:
' Line #288:
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd SPWNotebookComponentType 
' 	ArgsMemCall (Call) Add 0x0001 
' Line #289:
' 	Dim 
' 	OptionBase 
' 	LitDI2 0x0002 
' 	OptionBase 
' 	LitDI2 0x0003 
' 	VarDefn PlotColumnCountArray
' Line #290:
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #291:
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	LitDI2 0x0000 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #292:
' 	LitDI4 0x47FF 0x01E8 
' 	LitDI2 0x0002 
' 	LitDI2 0x0000 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #293:
' 	LitDI2 0x0001 
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #294:
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	LitDI2 0x0001 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #295:
' 	LitDI4 0x47FF 0x01E8 
' 	LitDI2 0x0002 
' 	LitDI2 0x0001 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #296:
' 	LitDI2 0x0003 
' 	LitDI2 0x0000 
' 	LitDI2 0x0002 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #297:
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	LitDI2 0x0002 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #298:
' 	LitDI4 0x47FF 0x01E8 
' 	LitDI2 0x0002 
' 	LitDI2 0x0002 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #299:
' 	LitDI2 0x0003 
' 	Ld deleteWorksheets 
' 	FnUBound 0x0000 
' 	Add 
' 	LitDI2 0x0000 
' 	LitDI2 0x0003 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #300:
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	LitDI2 0x0003 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #301:
' 	LitDI4 0x47FF 0x01E8 
' 	LitDI2 0x0002 
' 	LitDI2 0x0003 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #302:
' Line #303:
' 	Dim 
' 	VarDefn CurrentPageItem
' Line #304:
' 	OptionBase 
' 	LitDI2 0x0000 
' 	Redim CurrentPageItem 0x0001 (As Variant)
' Line #305:
' Line #306:
' 	LitDI2 0x0004 
' 	LitDI2 0x0000 
' 	ArgsSt CurrentPageItem 0x0001 
' Line #307:
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
' Line #308:
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemLd Graphs 0x0001 
' 	ArgsMemLd SelectObject 0x0001 
' 	ArgsMemCall (Call) testSigmaPlot 0x0000 
' Line #309:
' Line #310:
' 	LitStr 0x0006 "Site y"
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemLd Graphs 0x0001 
' 	ArgsMemLd SelectObject 0x0001 
' 	MemSt Name 
' Line #311:
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
' Line #312:
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
' Line #313:
' Line #314:
' 	Ld SLA_SELECTDIM 
' 	Ld SAA_OPTIONS 
' 	LitDI2 0x0003 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #315:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_GPM_SETPLOTATTR 
' 	LitDI2 0x000A 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #316:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_GPM_SETPLOTATTR 
' 	LitDI4 0x0001 0x0310 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #317:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_GPM_SETPLOTATTR 
' 	LitDI4 0x0402 0x00C0 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #318:
' 	Ld SAA_FROMVAL 
' 	Ld SAA_TOVAL 
' 	LitStr 0x0001 "0"
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #319:
' 	Ld SAA_FROMVAL 
' 	Ld GraphPages 
' 	Ld lMaxHistHeigh 
' 	Coerce (Str) 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #320:
' Line #321:
' 	Ld SLA_SELECTDIM 
' 	Ld SAA_OPTIONS 
' 	LitDI2 0x0003 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #322:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #323:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_COLOR 
' 	LitDI2 0x000A 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #324:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #325:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SEA_THICKNESS 
' 	LitHI4 0xFFFF 0x00FF 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #326:
' 	Ld SEA_COLORCOL 
' 	Ld GPM_SETAXISATTRSTRING 
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #327:
' 	Ld SEA_COLORCOL 
' 	Ld SLA_CONTOURFILLTYPE 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #328:
' 	Ld SLA_SELECTDIM 
' 	Ld SAA_SELECTLINE 
' 	LitDI2 0x0001 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #329:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0005 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #330:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SEA_THICKNESS 
' 	LitHI4 0xFFFF 0x00FF 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #331:
' 	Ld SEA_COLORCOL 
' 	Ld GPM_SETAXISATTRSTRING 
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #332:
' 	Ld SEA_COLORCOL 
' 	Ld SLA_CONTOURFILLTYPE 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #333:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #334:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SAA_SUB1OPTIONS 
' 	LitDI2 0x0512 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #335:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SAA_SUB1OPTIONS 
' 	LitDI2 0x0F31 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #336:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0001 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #337:
' Line #338:
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd SPWNotebookComponentType 
' 	ArgsMemCall (Call) Add 0x0001 
' Line #339:
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #340:
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	LitDI2 0x0000 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #341:
' 	LitDI4 0x47FF 0x01E8 
' 	LitDI2 0x0002 
' 	LitDI2 0x0000 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #342:
' 	LitDI2 0x0001 
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #343:
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	LitDI2 0x0001 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #344:
' 	LitDI4 0x47FF 0x01E8 
' 	LitDI2 0x0002 
' 	LitDI2 0x0001 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #345:
' 	LitDI2 0x0003 
' 	LitDI2 0x0000 
' 	LitDI2 0x0002 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #346:
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	LitDI2 0x0002 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #347:
' 	LitDI4 0x47FF 0x01E8 
' 	LitDI2 0x0002 
' 	LitDI2 0x0002 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #348:
' 	LitDI2 0x0003 
' 	Ld deleteWorksheets 
' 	FnUBound 0x0000 
' 	Add 
' 	LitDI2 0x0000 
' 	LitDI2 0x0003 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #349:
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	LitDI2 0x0003 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #350:
' 	LitDI4 0x47FF 0x01E8 
' 	LitDI2 0x0002 
' 	LitDI2 0x0003 
' 	ArgsSt PlotColumnCountArray 0x0002 
' Line #351:
' Line #352:
' 	OptionBase 
' 	LitDI2 0x0000 
' 	Redim CurrentPageItem 0x0001 (As Variant)
' Line #353:
' Line #354:
' 	LitDI2 0x0004 
' 	LitDI2 0x0000 
' 	ArgsSt CurrentPageItem 0x0001 
' Line #355:
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
' Line #356:
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemLd Graphs 0x0001 
' 	ArgsMemLd SelectObject 0x0001 
' 	ArgsMemCall (Call) testSigmaPlot 0x0000 
' Line #357:
' Line #358:
' 	LitStr 0x0006 "Site y"
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemLd Graphs 0x0001 
' 	ArgsMemLd SelectObject 0x0001 
' 	MemSt Name 
' Line #359:
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
' Line #360:
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
' Line #361:
' Line #362:
' 	Ld SLA_SELECTDIM 
' 	Ld SAA_OPTIONS 
' 	LitDI2 0x0003 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #363:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #364:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_COLOR 
' 	LitDI2 0x000A 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #365:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #366:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SEA_THICKNESS 
' 	LitHI4 0xFFFF 0x00FF 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #367:
' 	Ld SEA_COLORCOL 
' 	Ld GPM_SETAXISATTRSTRING 
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #368:
' 	Ld SEA_COLORCOL 
' 	Ld SLA_CONTOURFILLTYPE 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #369:
' 	Ld SLA_SELECTDIM 
' 	Ld SAA_SELECTLINE 
' 	LitDI2 0x0001 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #370:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0005 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #371:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SEA_THICKNESS 
' 	LitHI4 0xFFFF 0x00FF 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #372:
' 	Ld SEA_COLORCOL 
' 	Ld GPM_SETAXISATTRSTRING 
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #373:
' 	Ld SEA_COLORCOL 
' 	Ld SLA_CONTOURFILLTYPE 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #374:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #375:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SAA_SUB1OPTIONS 
' 	LitDI2 0x0512 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #376:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SAA_SUB1OPTIONS 
' 	LitDI2 0x0F31 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #377:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0001 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #378:
' Line #379:
' Line #380:
' Line #381:
' 	Ld i 
' 	Ld theWorksheets 
' 	FnUBound 0x0000 
' 	Lt 
' 	IfBlock 
' Line #382:
' 	LitDI2 0x0001 
' 	Ld SPApplication 
' 	MemLd SPWNotebookComponentType 
' 	ArgsMemCall (Call) Add 0x0001 
' Line #383:
' 	EndIfBlock 
' Line #384:
' 	StartForVariable 
' 	Next 
' Line #385:
' 	EndSub 
' Line #386:
' Line #387:
' 	FuncDefn (Sub _B_var_lMaxHistHeight())
' Line #388:
' 	Dim (Const) 
' 	LitHI2 0x0406 
' 	VarDefn SAA_TOVAL
' Line #389:
' 	Dim (Const) 
' 	LitHI2 0x0407 
' 	VarDefn GraphPages
' Line #390:
' 	Dim (Const) 
' 	LitHI2 0x0301 
' 	VarDefn SLA_SELECTDIM
' Line #391:
' 	Dim (Const) 
' 	LitDI2 0x0401 
' 	VarDefn SEA_COLORCOL
' Line #392:
' 	Dim (Const) 
' 	LitDI2 0x0308 
' 	VarDefn SAA_OPTIONS
' Line #393:
' 	Dim (Const) 
' 	LitDI2 0x0403 
' 	VarDefn _B_var_GPM_SETPLOTATTR
' Line #394:
' 	Dim (Const) 
' 	LitDI2 0x0408 
' 	VarDefn SAA_FROMVAL
' Line #395:
' 	Dim (Const) 
' 	LitDI2 0x0615 
' 	VarDefn GPM_SETAXISATTRSTRING
' Line #396:
' 	Dim (Const) 
' 	LitDI2 0x0613 
' 	VarDefn SLA_CONTOURFILLTYPE
' Line #397:
' 	Dim (Const) 
' 	LitDI2 0x0358 
' 	VarDefn SAA_SELECTLINE
' Line #398:
' 	Dim (Const) 
' 	LitDI2 0x040A 
' 	VarDefn SEA_THICKNESS
' Line #399:
' 	Dim (Const) 
' 	LitDI2 0x0601 
' 	VarDefn SEA_COLOR
' Line #400:
' 	Dim (Const) 
' 	LitDI2 0x0606 
' 	VarDefn _B_var_SEA_THICKNESS
' Line #401:
' 	Dim (Const) 
' 	LitDI2 0x0410 
' 	VarDefn _B_var_SAA_SUB1OPTIONS
' Line #402:
' Line #403:
' 	Dim 
' 	VarDefn Module2 (As Object)
' Line #404:
' 	SetStmt 
' 	LitStr 0x0017 "SigmaPlot.Application.1"
' 	ArgsLd CreateObject 0x0001 
' 	Set Module2 
' Line #405:
' 	LitVarSpecial (True)
' 	Ld Module2 
' 	MemSt Application 
' Line #406:
' 	LitStr 0x0006 "Site y"
' 	LitDI2 0x0000 
' 	LitDI2 0x0000 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemLd Graphs 0x0001 
' 	ArgsMemLd SelectObject 0x0001 
' 	MemSt Name 
' Line #407:
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
' Line #408:
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
' Line #409:
' Line #410:
' 	Ld SLA_SELECTDIM 
' 	Ld SAA_OPTIONS 
' 	LitDI2 0x0003 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #411:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_GPM_SETPLOTATTR 
' 	LitDI2 0x000A 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #412:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_GPM_SETPLOTATTR 
' 	LitDI4 0x0001 0x0310 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #413:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_GPM_SETPLOTATTR 
' 	LitDI4 0x0402 0x00C0 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #414:
' 	Ld SAA_FROMVAL 
' 	Ld SAA_TOVAL 
' 	LitStr 0x0001 "0"
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #415:
' 	Ld SAA_FROMVAL 
' 	Ld GraphPages 
' 	LitStr 0x0003 "150"
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #416:
' Line #417:
' 	Ld SLA_SELECTDIM 
' 	Ld SAA_OPTIONS 
' 	LitDI2 0x0003 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #418:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #419:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_COLOR 
' 	LitDI2 0x000A 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #420:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #421:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SEA_THICKNESS 
' 	LitHI4 0xFFFF 0x00FF 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #422:
' 	Ld SEA_COLORCOL 
' 	Ld GPM_SETAXISATTRSTRING 
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #423:
' 	Ld SEA_COLORCOL 
' 	Ld SLA_CONTOURFILLTYPE 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #424:
' 	Ld SLA_SELECTDIM 
' 	Ld SAA_SELECTLINE 
' 	LitDI2 0x0001 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #425:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0005 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #426:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SEA_THICKNESS 
' 	LitHI4 0xFFFF 0x00FF 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #427:
' 	Ld SEA_COLORCOL 
' 	Ld GPM_SETAXISATTRSTRING 
' 	LitDI2 0x0002 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #428:
' 	Ld SEA_COLORCOL 
' 	Ld SLA_CONTOURFILLTYPE 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #429:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0004 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #430:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SAA_SUB1OPTIONS 
' 	LitDI2 0x0512 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #431:
' 	Ld SEA_COLORCOL 
' 	Ld _B_var_SAA_SUB1OPTIONS 
' 	LitDI2 0x0F31 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #432:
' 	Ld SEA_COLORCOL 
' 	Ld SEA_THICKNESS 
' 	LitDI2 0x0001 
' 	Ld Module2 
' 	MemLd CT_GRAPHICPAGE 
' 	MemLd CreateWizardGraph 
' 	ArgsMemCall (Call) GPM_SETAXISATTR 0x0003 
' Line #433:
' Line #434:
' 	EndSub 
' Line #435:
' Line #436:
' 	FuncDefn (Function _B_var_buildEpocList(objTTX, iXAxisIndexAs, returnArr))
' Line #437:
' 	QuoteRem 0x0004 0x0030 "build list of epocs for the given axis epoc name"
' Line #438:
' Line #439:
' 	Dim 
' 	VarDefn AxisEp (As Dictionary)
' Line #440:
' 	SetStmt 
' 	New id_FFFF
' 	Set AxisEp 
' Line #441:
' Line #442:
' 	Dim 
' 	VarDefn dblStartTime (As Double)
' Line #443:
' 	Dim 
' 	VarDefn varReturn (As Variant)
' Line #444:
' Line #445:
' 	Dim 
' 	VarDefn i (As Integer)
' Line #446:
' 	Dim 
' 	VarDefn j (As Integer)
' Line #447:
' Line #448:
' 	Ld iXAxisIndexAs 
' 	LitStr 0x0007 "Channel"
' 	Eq 
' 	IfBlock 
' Line #449:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0001 
' 	LitDI2 0x0020 
' 	For 
' Line #450:
' 	Ld i 
' 	LitDI2 0x0000 
' 	Ld AxisEp 
' 	ArgsMemCall (Call) Add 0x0002 
' Line #451:
' 	StartForVariable 
' 	Next 
' Line #452:
' 	ElseBlock 
' Line #453:
' 	Do 
' Line #454:
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
' Line #455:
' 	Ld i 
' 	LitDI2 0x0000 
' 	Eq 
' 	IfBlock 
' Line #456:
' 	ExitDo 
' Line #457:
' 	EndIfBlock 
' Line #458:
' Line #459:
' 	LitDI2 0x0000 
' 	Ld i 
' 	LitDI2 0x0000 
' 	Ld objTTX 
' 	ArgsMemLd ParseEvInfoV 0x0003 
' 	St varReturn 
' Line #460:
' 	StartForVariable 
' 	Ld j 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld i 
' 	LitDI2 0x0001 
' 	Sub 
' 	Paren 
' 	For 
' Line #461:
' 	LitDI2 0x0006 
' 	Ld j 
' 	ArgsLd varReturn 0x0002 
' 	Ld AxisEp 
' 	ArgsMemLd Exists 0x0001 
' 	Not 
' 	IfBlock 
' Line #462:
' 	LitDI2 0x0006 
' 	Ld j 
' 	ArgsLd varReturn 0x0002 
' 	LitStr 0x0000 ""
' 	Ld AxisEp 
' 	ArgsMemCall (Call) Add 0x0002 
' Line #463:
' 	EndIfBlock 
' Line #464:
' 	LitDI2 0x0005 
' 	Ld j 
' 	ArgsLd varReturn 0x0002 
' 	LitDI2 0x0001 
' 	LitDI4 0x86A0 0x0001 
' 	Div 
' 	Paren 
' 	Add 
' 	St dblStartTime 
' Line #465:
' 	StartForVariable 
' 	Next 
' Line #466:
' Line #467:
' 	Ld i 
' 	LitDI2 0x01F4 
' 	Lt 
' 	IfBlock 
' Line #468:
' 	ExitDo 
' Line #469:
' 	EndIfBlock 
' Line #470:
' 	Loop 
' Line #471:
' 	EndIfBlock 
' Line #472:
' Line #473:
' Line #474:
' Line #475:
' 	Ld returnArr 
' 	IfBlock 
' Line #476:
' 	Dim 
' 	VarDefn _B_var_returnArr
' Line #477:
' 	Dim 
' 	VarDefn id_0794 (As Variant)
' Line #478:
' 	Ld AxisEp 
' 	MemLd Keys 
' 	St id_0794 
' Line #479:
' 	OptionBase 
' 	Ld id_0794 
' 	FnUBound 0x0000 
' 	Redim _B_var_returnArr 0x0001 (As Variant)
' Line #480:
' Line #481:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld id_0794 
' 	FnUBound 0x0000 
' 	For 
' Line #482:
' 	Ld id_0794 
' 	FnUBound 0x0000 
' 	Ld i 
' 	Sub 
' 	ArgsLd id_0794 0x0001 
' 	Ld i 
' 	ArgsSt _B_var_returnArr 0x0001 
' Line #483:
' 	StartForVariable 
' 	Next 
' Line #484:
' 	Ld _B_var_returnArr 
' 	St _B_var_buildEpocList 
' Line #485:
' 	ElseBlock 
' Line #486:
' 	Ld AxisEp 
' 	MemLd Keys 
' 	St _B_var_buildEpocList 
' Line #487:
' 	EndIfBlock 
' Line #488:
' Line #489:
' 	EndFunc 
' Line #490:
' Line #491:
' Line #492:
' 	FuncDefn (Function _B_var_processSearch(ByRef objTTX, ByRef _B_var_arrOtherEp, ByRef strSearchString, iOtherEpocIndex, xAxisSearchString As String, yOffset, zOffset, xOffes, 1, Le))
' Line #493:
' 	Dim 
' 	VarDefn i (As Integer)
' Line #494:
' 	Dim 
' 	VarDefn j (As Integer)
' Line #495:
' 	Dim 
' 	VarDefn _B_var_objTTX (As String)
' Line #496:
' 	Dim 
' 	VarDefn strHeading (As String)
' Line #497:
' 	Dim 
' 	VarDefn Label1 (As String)
' Line #498:
' Line #499:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld iOtherEpocIndex 
' 	ArgsLd strSearchString 0x0001 
' 	FnUBound 0x0000 
' 	For 
' Line #500:
' 	Ld iOtherEpocIndex 
' 	ArgsLd _B_var_arrOtherEp 0x0001 
' 	LitStr 0x0007 "Channel"
' 	Ne 
' 	IfBlock 
' Line #501:
' 	QuoteRem 0x000C 0x0014 "add to search string"
' Line #502:
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
' Line #503:
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
' Line #504:
' 	ElseBlock 
' Line #505:
' 	Ld xAxisSearchString 
' 	St _B_var_objTTX 
' Line #506:
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
' Line #507:
' 	Ld i 
' 	Ld iOtherEpocIndex 
' 	ArgsLd strSearchString 0x0001 
' 	IndexLd 0x0001 
' 	St 1 
' Line #508:
' 	EndIfBlock 
' Line #509:
' 	Ld iOtherEpocIndex 
' 	Ld _B_var_arrOtherEp 
' 	FnUBound 0x0000 
' 	Lt 
' 	IfBlock 
' Line #510:
' 	QuoteRem 0x000C 0x002F "there are still more epocs to add to the search"
' Line #511:
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
' 	ArgsCall (Call) _B_var_processSearch 0x000A 
' Line #512:
' 	ElseBlock 
' Line #513:
' 	QuoteRem 0x000C 0x004B "we have reached the end of the list of epocs - can actually do a search now"
' Line #514:
' 	Ld _B_var_objTTX 
' 	LitDI2 0x0005 
' 	ArgsLd Right 0x0002 
' 	LitStr 0x0005 " and "
' 	Eq 
' 	IfBlock 
' 	QuoteRem 0x003D 0x0045 "this should always be the case - should be a trailing 'and' to remove"
' Line #515:
' 	Ld _B_var_objTTX 
' 	Ld _B_var_objTTX 
' 	FnLen 
' 	LitDI2 0x0005 
' 	Sub 
' 	ArgsLd Left 0x0002 
' 	St strHeading 
' Line #516:
' 	ElseBlock 
' Line #517:
' 	Ld _B_var_objTTX 
' 	St strHeading 
' Line #518:
' 	EndIfBlock 
' Line #519:
' 	Ld strHeading 
' 	Ld objTTX 
' 	ArgsMemCall (Call) SetFilterWithDescEx 0x0001 
' Line #520:
' Line #521:
' 	Ld yOffset 
' 	LitDI2 0x0001 
' 	Eq 
' 	Ld zOffset 
' 	LitDI2 0x0001 
' 	Eq 
' 	And 
' 	IfBlock 
' Line #522:
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
' Line #523:
' 	Ld vYAxisKeys 
' 	Ld buildOptionLists 
' 	Ld _B_var_iColOffset 
' 	Ld _B_var_Const 
' 	Ld i 
' 	Ld xOffes 
' 	Mul 
' 	Paren 
' 	ArgsCall (Call) SubwriteAxes 0x0005 
' Line #524:
' 	EndIfBlock 
' Line #525:
' Line #526:
' 	Ld objTTX 
' 	Ld yOffset 
' 	Ld zOffset 
' 	Ld i 
' 	Ld xOffes 
' 	Mul 
' 	Ld 1 
' 	ArgsCall (Call) _B_var_writeResults 0x0005 
' Line #527:
' 	EndIfBlock 
' Line #528:
' 	StartForVariable 
' 	Next 
' Line #529:
' Line #530:
' 	EndFunc 
' Line #531:
' Line #532:
' 	FuncDefn (Sub _B_var_writeResults(ByRef objTTX, yOffset, zOffset, xOffes, 1))
' Line #533:
' 	Dim 
' 	VarDefn varReturn (As Variant)
' Line #534:
' 	Dim 
' 	VarDefn varChanData (As Variant)
' Line #535:
' Line #536:
' 	Dim 
' 	VarDefn dblStartTime (As Double)
' Line #537:
' 	Dim 
' 	VarDefn dblEndTime (As Double)
' Line #538:
' 	Dim 
' 	VarDefn IsEmpty (As Double)
' Line #539:
' Line #540:
' 	Dim 
' 	VarDefn i (As Long)
' Line #541:
' 	Dim 
' 	VarDefn j (As Long)
' Line #542:
' 	Dim 
' 	VarDefn k (As Long)
' Line #543:
' Line #544:
' 	Dim 
' 	VarDefn histTmp (As Long)
' Line #545:
' Line #546:
' 	LitStr 0x0004 "Swep"
' 	LitDI2 0x0000 
' 	Ld objTTX 
' 	ArgsMemLd GetEpocsExV 0x0002 
' 	St varReturn 
' Line #547:
' 	Ld varReturn 
' 	ArgsLd Dib 0x0001 
' 	IfBlock 
' Line #548:
' 	StartForVariable 
' 	Ld i 
' 	EndForVariable 
' 	LitDI2 0x0000 
' 	Ld varReturn 
' 	LitDI2 0x0002 
' 	FnUBound 0x0001 
' 	For 
' Line #549:
' 	LitDI2 0x0002 
' 	Ld i 
' 	ArgsLd varReturn 0x0002 
' 	Ld _B_var_lIgnoreFirstMsec 
' 	Add 
' 	St dblStartTime 
' Line #550:
' 	Ld dblStartTime 
' 	Ld lBinWidth 
' 	Add 
' 	Ld _B_var_lIgnoreFirstMsec 
' 	Add 
' 	St dblEndTime 
' Line #551:
' 	Ld dblStartTime 
' 	St IsEmpty 
' Line #552:
' 	Do 
' Line #553:
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
' Line #554:
' 	Ld k 
' 	LitDI2 0x0000 
' 	Eq 
' 	IfBlock 
' Line #555:
' 	ExitDo 
' Line #556:
' 	EndIfBlock 
' Line #557:
' Line #558:
' 	Ld histTmp 
' 	Coerce (Lng) 
' 	Ld k 
' 	Coerce (Lng) 
' 	Add 
' 	St histTmp 
' Line #559:
' 	Ld k 
' 	LitDI2 0x01F4 
' 	Lt 
' 	IfBlock 
' Line #560:
' 	ExitDo 
' Line #561:
' 	ElseBlock 
' Line #562:
' 	Ld k 
' 	LitDI2 0x0001 
' 	Sub 
' 	LitDI2 0x0001 
' 	LitDI2 0x0006 
' 	Ld objTTX 
' 	ArgsMemLd ParseEvInfoV 0x0003 
' 	St varChanData 
' Line #563:
' 	LitDI2 0x0000 
' 	ArgsLd varChanData 0x0001 
' 	LitDI2 0x0001 
' 	LitDI4 0x86A0 0x0001 
' 	Div 
' 	Paren 
' 	Add 
' 	St dblStartTime 
' Line #564:
' 	EndIfBlock 
' Line #565:
' 	Loop 
' Line #566:
' 	Ld IsEmpty 
' 	St dblStartTime 
' Line #567:
' 	StartForVariable 
' 	Next 
' Line #568:
' Line #569:
' 	Ld yAxisEp 
' 	LitStr 0x0007 "Channel"
' 	Eq 
' 	IfBlock 
' Line #570:
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
' Line #571:
' 	Ld otherEp 
' 	LitStr 0x0007 "Channel"
' 	Eq 
' 	ElseIfBlock 
' Line #572:
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
' Line #573:
' 	ElseBlock 
' Line #574:
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
' Line #575:
' 	EndIfBlock 
' Line #576:
' 	LitDI2 0x0000 
' 	St histTmp 
' Line #577:
' 	EndIfBlock 
' Line #578:
' Line #579:
' 	EndSub 
' _VBA_PROJECT_CUR/VBA/Sheet2 - 1166 bytes
' _VBA_PROJECT_CUR/VBA/Sheet3 - 1150 bytes
