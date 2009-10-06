Attribute VB_Name = "ShrtFunctions"

Const TDT_ConnectSuccess = 0
Const TDT_ServerConnectFail = 1
Const TDT_TankConnectFail = 2
Const TDT_BlockConnectFail = 3

Function loadConfigParams( _
        ByRef outputWorkbook As Workbook, _
        ByRef thisWorkbook As Workbook, _
        ByRef stimStartEpoc As String, _
        ByRef lBinWidth As Long, _
        ByRef lIgnoreFirstMsec As Long, _
        ByRef lNumOfChans As Long, _
        ByRef iRowOffset As Integer, _
        ByRef iColOffset As Integer)
        
    loadConfigParams = True
        
    'load the stimulus start epoc
    If Not readCopyParam(outputWorkbook, thisWorkbook, "Variables (do not edit)", "B7", "", stimStartEpoc, vbString, False) Then
        loadConfigParams = False
    End If
    
    'load the bin width for histogram generation
    If Not readCopyParam(outputWorkbook, thisWorkbook, "Settings", "B1", "", lBinWidth, vbLong, False) Then
        loadConfigParams = False
    End If
        
    'load the # of msec to ignore at the start (for filtering stimulation artifact
    If Not readCopyParam(outputWorkbook, thisWorkbook, "Settings", "B2", "", lIgnoreFirstMsec, vbLong, False) Then
        loadConfigParams = False
    End If
    
    'read number of channels to process; write to output
    If Not readCopyParam(outputWorkbook, thisWorkbook, "Settings", "B3", "", lNumOfChans, vbLong, False) Then
        loadConfigParams = False
    End If
    
    'offsets to leave space at the top and left of the chart
    If Not readCopyParam(outputWorkbook, thisWorkbook, "Variables (do not edit)", "E4", "", iRowOffset, vbInteger, False) Then
        loadConfigParams = False
    End If
    If Not readCopyParam(outputWorkbook, thisWorkbook, "Variables (do not edit)", "E5", "", iColOffset, vbInteger, False) Then
        loadConfigParams = False
    End If
End Function

'reads a value from the input workbook, typechecks it, stores it in a (byref) variable to be passed back, and writes it to the matching spot on the output workbook
Function readCopyParam(ByRef outputWorkbook As Workbook, ByRef inputWorkbook As Workbook, strWorksheetName As String, strAddress As String, strParamNameAddress As String, ByRef theVariable As Variant, intType As Integer, blnBlankAllowed As Boolean)
    Const ErrType_NoError = 0
    Const ErrType_NotNumeric = 1
    Const ErrType_NotWholeNum = 2
    Const ErrType_TooLongForInt = 3
    Const ErrType_NotBoolean = 4
    Const ErrType_BlankNotAllowed = 5
    Const ErrType_UnsupportedDataTypeCheck = 6
    'const ErrType_DataType = 1
        
    Dim vInputValue As Variant
    Dim strParamName As String
    Dim intErrType As Integer
    
    readCopyParam = False
    
    intErrType = ErrType_NoError
    
    If intType <> 2 And _
        intType <> 3 And _
        intType <> 5 And _
        intType <> 8 And _
        intType <> 11 Then
            intErrType = ErrType_UnsupportedDataTypeCheck
    Else
        vInputValue = inputWorkbook.Worksheets(strWorksheetName).Range(strAddress).Value
        
        If intType = vbInteger Or intType = vbLong Or intType = vbDouble Then
            If Not IsNumeric(vInputValue) Then
                If Not (vInputValue = "" And blnBlankAllowed) Then
                    If (vInputValue = "" And Not blnBlankAllowed) Then
                        'blank value and blank is not allowed
                        ErrType_BlankNotAllowed
                    Else
                        intErrType = ErrType_DataType
                    End If
                End If
            Else
                If intType = vbInteger Or intType = vbLong Then
                    If Not Int(vInputValue) = vInputValue Then
                        intErrType = ErrType_NotWholeNum
                    ElseIf intType = vbInteger And Abs(vInputValue) > 32767 Then
                        intErrType = ErrType_TooLongForInt
                    End If
                End If
            End If
        ElseIf intType = vbBoolean Then
            If (vInputValue = "" And blnBlankAllowed) Then
                vInputValue = vbNull 'default value is False if blank
            ElseIf CInt(CBool(vInputValue)) <> vInputValue Then
                intErrType = ErrType_NotBoolean
            End If
        End If
        
        If intType <> ErrType_NoError Then
            If strParamNameAddress = "" Then 'if no param name address provided, we assume the title is located one cell left of the value address
                If inputWorkbook.Worksheets(strWorksheetName).Range(strAddress).Column > 1 Then 'label can only be in the next column to the left if we are not in the first column
                    strParamName = inputWorkbook.Worksheets(strWorksheetName).Range(strAddress).Offset(0, -1).Value
                Else
                    strParamName = "<no name specified>"
                End If
            Else
                strParamName = inputWorkbook.Worksheets(strWorksheetName).Range(strParamNameAddress).Value
            End If
            
            Select Case intErrType
                Case ErrType_NotNumeric:
                    MsgBox "The parameter '" & strParamName & "' at location " & strWorksheetName & "." & strAddressis & " should be numeric (value: " & vInputValue & ")"
                Case ErrType_NotWholeNum:
                    MsgBox "The parameter '" & strParamName & "' at location " & strWorksheetName & "." & strAddressis & " should be a whole number (value: " & vInputValue & ")"
                Case ErrType_TooLongForInt:
                    MsgBox "The parameter '" & strParamName & "' at location " & strWorksheetName & "." & strAddressis & " is too large (value: " & vInputValue & "; must be between -32767 and 32767)"
                Case ErrType_NotBoolean:
                    MsgBox "The parameter '" & strParamName & "' at location " & strWorksheetName & "." & strAddressis & " should be 'TRUE' or 'FALSE' (value: " & vInputValue & ")"
                Case ErrType_BlankNotAllowed:
                    MsgBox "The parameter '" & strParamName & "' at location " & strWorksheetName & "." & strAddressis & " must not be blank"
            End Select
        Else
            theVariable = vInputValue
            outputWorkbook.Worksheets(strWorksheetName).Range(strAddress).Value = theVariable
            
            readCopyParam = True
        End If
    End If
End Function


Function connectToTDT(objTTX As TTankX)
    connectToTDT = False
    
    If theServer = "" Then
        theServer = Worksheets("Variables (do not edit)").Range("B1").Value
        theTank = Worksheets("Variables (do not edit)").Range("B2").Value
        theBlock = Worksheets("Variables (do not edit)").Range("B3").Value
    End If
    Select Case testSettings(objTTX, theServer, theTank, theBlock)
        Case TDT_ConnectSuccess:
            connectToTDT = True
    End Select
End Function

Function testTDTConnection(objTTX As TTankX, ActServer, ActTank, ActBlock)
    testTDTConnection = TDT_ConnectSuccess
    If objTTX.ConnectServer(ActServer, "Me") <> CLng(1) Then
        testTDTConnection = TDT_ServerConnectFail
        Exit Function
    ElseIf objTTX.OpenTank(ActTank, "R") <> CLng(1) Then
        objTTX.ReleaseServer
        testTDTConnection = TDT_TankConnectFail
        Exit Function
    ElseIf objTTX.SelectBlock(ActBlock) <> CLng(1) Then
        objTTX.CloseTank
        objTTX.ReleaseServer
        testTDTConnection = TDT_BlockConnectFail
    End If
    
End Function

