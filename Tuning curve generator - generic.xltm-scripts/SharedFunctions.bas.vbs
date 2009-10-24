Attribute VB_Name = "SharedFunctions"
Option Explicit
Global Const TDT_Unknown = -1
Global Const TDT_ConnectSuccess = 0
Global Const TDT_ServerConnectFail = 1
Global Const TDT_TankNotProvided = 2
Global Const TDT_TankInvalidMode = 3
Global Const TDT_TankConnectFail = 4
Global Const TDT_BlockNotProvided = 5
Global Const TDT_BlockConnectFail = 6

'reads a value from the input workbook, typechecks it, stores it in a (byref) variable to be passed back, and writes it to the matching spot on the output workbook
Function readCopyParam(ByRef outputWorkbook As Workbook, ByRef inputWorkbook As Workbook, strWorksheetName As String, strAddress As String, strParamNameAddress As String, ByRef theVariable As Variant, intType As Integer, blnBlankAllowed As Boolean)
        
    Dim vInputValue As Variant
    Dim strParamName As String
    
    readCopyParam = False
    
    If strParamNameAddress = "" Then 'if no param name address provided, we assume the title is located one cell left of the value address
        If inputWorkbook.Worksheets(strWorksheetName).Range(strAddress).Column > 1 Then 'label can only be in the next column to the left if we are not in the first column
            strParamName = inputWorkbook.Worksheets(strWorksheetName).Range(strAddress).Offset(0, -1).Value
        Else
            strParamName = "<no name specified>"
        End If
    Else
        strParamName = inputWorkbook.Worksheets(strWorksheetName).Range(strParamNameAddress).Value
    End If
    
    vInputValue = inputWorkbook.Worksheets(strWorksheetName).Range(strAddress).Value
    If checkDataType(vInputValue, intType, strParamName, blnBlankAllowed, strWorksheetName, strAddress) = 0 Then
            theVariable = vInputValue
            outputWorkbook.Worksheets(strWorksheetName).Range(strAddress).Value = theVariable
            readCopyParam = True
    End If
End Function

'reads a value from the input workbook, typechecks it, stores it in a (byref) variable to be passed back, and writes it to the matching spot on the output workbook
Function checkDataType(ByRef vInputValue As Variant, intType As Integer, strParamName As String, Optional blnBlankAllowed As Variant, Optional strWorksheetName, Optional strAddress As Variant) As Integer
    Const ErrType_NoError = 0
    Const ErrType_NotNumeric = 1
    Const ErrType_NotWholeNum = 2
    Const ErrType_TooLongForInt = 3
    Const ErrType_NotBoolean = 4
    Const ErrType_BlankNotAllowed = 5
    Const ErrType_UnsupportedDataTypeCheck = 6
    'const ErrType_DataType = 1
    
    If IsMissing(blnBlankAllowed) Or Not VarType(blnBlankAllowed) = vbBoolean Then
        blnBlankAllowed = False
    End If
    
    Dim intErrType As Integer
    
    intErrType = ErrType_NoError
    
    If intType <> 2 And _
        intType <> 3 And _
        intType <> 5 And _
        intType <> 8 And _
        intType <> 11 Then
            intErrType = ErrType_UnsupportedDataTypeCheck
    Else
        If intType = vbInteger Or intType = vbLong Or intType = vbDouble Then
            If Not IsNumeric(vInputValue) Then
                If Not (vInputValue = "" And blnBlankAllowed) Then
                    If (vInputValue = "" And Not blnBlankAllowed) Then
                        'blank value and blank is not allowed
                        intErrType = ErrType_BlankNotAllowed
                    Else
                        intErrType = ErrType_NotNumeric
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
            'ElseIf CInt(CBool(vInputValue)) <> CInt(vInputValue) Then
            '    intErrType = ErrType_NotBoolean
            End If
        End If
        
        If intErrType <> ErrType_NoError Then
            If IsMissing(strWorksheetName) Then strWorksheetName = "<Not Available>"
            If IsMissing(strAddress) Then strAddress = "<Not Available>"
            
            Select Case intErrType
                Case ErrType_NotNumeric:
                    MsgBox "The parameter '" & strParamName & "' at location " & strWorksheetName & "." & strAddress & " should be numeric (value: " & vInputValue & ")"
                Case ErrType_NotWholeNum:
                    MsgBox "The parameter '" & strParamName & "' at location " & strWorksheetName & "." & strAddress & " should be a whole number (value: " & vInputValue & ")"
                Case ErrType_TooLongForInt:
                    MsgBox "The parameter '" & strParamName & "' at location " & strWorksheetName & "." & strAddress & " is too large (value: " & vInputValue & "; must be between -32767 and 32767)"
                Case ErrType_NotBoolean:
                    MsgBox "The parameter '" & strParamName & "' at location " & strWorksheetName & "." & strAddress & " should be 'TRUE' or 'FALSE' (value: " & vInputValue & ")"
                Case ErrType_BlankNotAllowed:
                    MsgBox "The parameter '" & strParamName & "' at location " & strWorksheetName & "." & strAddress & " must not be blank"
            End Select
        End If
    End If
    checkDataType = intErrType
End Function


Function connectToTDT(ByRef objTTX As TTankX, Optional blnAllowReadFromFile As Variant, Optional ByRef sServer As Variant, _
        Optional ByRef sTank As Variant, Optional ByRef sBlock As Variant, Optional sTankMode As Variant) As Variant
        
    Const STAT_TANKSTATE = &H0
    Dim vReturnValues As Variant
    vReturnValues = Array(TDT_Unknown, "", "", "", "") 'connection flag, and error message (if one is present)
    Dim lTDTResponse As Long
    
    'if there isn't already a connect to TDT, crease it
    If objTTX Is Nothing Then
        Set objTTX = CreateObject("TTank.X")
    Else
        If objTTX.GetStatus(STAT_TANKSTATE) <> -1 Then 'check if a tank is open, and if so close it
            Call objTTX.CloseTank
        End If
    End If
    
    If IsMissing(sServer) Or (blnAllowReadFromFile And sServer = "") Then
        sServer = Worksheets("Variables (do not edit)").Range("B1").Value
    End If
    If sServer = "" Then
        sServer = "Local" 'if no server provided, try the local server
    End If

        lTDTResponse = objTTX.ConnectServer(sServer, "Me")
        If lTDTResponse <> CLng(1) Then
            'problem occurred making server connection
            vReturnValues(0) = TDT_ServerConnectFail
            vReturnValues(1) = objTTX.GetError
            Set objTTX = Nothing
        Else 'successful connection to TDT server
            If IsMissing(sTank) Or (blnAllowReadFromFile And sTank = "") Then
                sTank = Worksheets("Variables (do not edit)").Range("B2").Value
            End If
            If sTank = "" Then 'if no tank provided or readable, then give an error
                vReturnValues(0) = TDT_TankNotProvided
                Call objTTX.ReleaseServer
                Set objTTX = Nothing
            Else
                'check the given mode is acceptable (if one is provided)
                If IsMissing(sTankMode) Then
                    sTankMode = "R"
                End If
                If UCase(sTankMode) <> "R" And UCase(sTankMode) <> "W" And UCase(sTankMode) <> "C" And UCase(sTankMode) <> "M" Then
                    vReturnValues(0) = TDT_TankInvalidMode
                    Call objTTX.ReleaseServer
                Else
                    lTDTResponse = objTTX.OpenTank(sTank, sTankMode)
                    If lTDTResponse <> CLng(1) Then
                        vReturnValues(0) = TDT_TankConnectFail
                        vReturnValues(1) = objTTX.GetError
                        Call objTTX.ReleaseServer
                    Else
                        If IsMissing(sBlock) Or (blnAllowReadFromFile And sBlock = "") Then
                            sBlock = Worksheets("Variables (do not edit)").Range("B3").Value
                        End If
                        If sBlock = "" Then 'if no tank provided or readable, then give an error
                            vReturnValues(0) = TDT_BlockNotProvided
                            Call objTTX.CloseTank
                            Call objTTX.ReleaseServer
                            Set objTTX = Nothing
                        Else
                            If objTTX.SelectBlock(sBlock) <> CLng(1) Then
                                vReturnValues(0) = TDT_BlockConnectFail
                                vReturnValues(1) = objTTX.GetError
                                Call objTTX.CloseTank
                                Call objTTX.ReleaseServer
                                Set objTTX = Nothing
                            Else
                                vReturnValues(0) = TDT_ConnectSuccess
                            End If
                        End If
                    End If
                End If
            End If
        End If
    
    vReturnValues(2) = sServer
    vReturnValues(3) = sTank
    vReturnValues(4) = sBlock
    connectToTDT = vReturnValues
End Function


Function LoWord(wInt)
  LoWord = wInt And &HFFFF&
End Function

Function HiWord(wInt)
  HiWord = wInt \ &H10000 And &HFFFF&
End Function
Function MAKELPARAM(wLow, wHigh)
  MAKELPARAM = LoWord(wLow) Or (&H10000 * LoWord(wHigh))
End Function

Sub delayMe(secs As Long)
    Dim newHour As Integer
    Dim newMinute As Integer
    Dim newSecond As Integer
    Dim waitTime As String
    newHour = Hour(Now())
    newMinute = Minute(Now())
    newSecond = Second(Now()) + secs
    waitTime = TimeSerial(newHour, newMinute, newSecond)
    Call Application.Wait(waitTime)
End Sub

Function checkPathExists(thePath As String) As Boolean
    
    If thePath = "" Then
        checkPathExists = False
    Else
        Dim objFS As FileSystemObject
        Set objFS = CreateObject("Scripting.FileSystemObject")
        If Not objFS.FolderExists(thePath) Then
            checkPathExists = False
        Else
            checkPathExists = True
        End If
        Set objFS = Nothing
    End If

End Function
Function getFilename(theFilename As String, filenameOnTargetDrive As String) As String
    
    Dim objFS As FileSystemObject
    Set objFS = CreateObject("Scripting.FileSystemObject")
    
    If theFilename = "" Then
        getFilename = ""
    Else
        If Right(Left(theFilename, 2), 1) <> ":" Then
            Dim theDrive As String
            theDrive = objFS.GetDriveName(filenameOnTargetDrive)
            getFilename = theDrive & theFilename
        End If
        
        If Not objFS.FileExists(getFilename) Then
            getFilename = ""
        End If
    End If
    
    Set objFS = Nothing
    
End Function

Function getDirName(theDir As String, filenameOnTargetDrive As String) As String
    
    Dim objFS As FileSystemObject
    Set objFS = CreateObject("Scripting.FileSystemObject")
    
    If theDir = "" Then
        getDirName = objFS.GetParentFolderName(filenameOnTargetDrive)
    ElseIf Right(Left(theDir, 2), 1) <> ":" Then
        Dim theDrive As String
        theDrive = objFS.GetDriveName(filenameOnTargetDrive)
        getDirName = theDrive & theDir
    End If
    
    If Not objFS.FolderExists(getDirName) Then
        getDirName = ""
    End If
    
    Set objFS = Nothing
End Function

Function selectAllInList(theList As Variant)
    Dim i As Integer
    For i = 0 To (theList.ListCount - 1)
        theList.Selected(i) = True
    Next
End Function

Function deselectAllInList(theList As Variant)
    Dim i As Integer
    For i = 0 To (theList.ListCount - 1)
        theList.Selected(i) = False
    Next
End Function


'gives msg box with error if one occurred connecting to TDT. Returns True if an error occurred, false otherwise
Function connectToTDTReportError(vReturnArr As Variant, Optional vUseMessageBox As Variant) As String

    Dim strGenericErrAppend As String
    strGenericErrAppend = " (State: " & vReturnArr(0) & ":Msg " & vReturnArr(1) & ": Server '" & vReturnArr(2) & "': Tank '" & vReturnArr(3) & "': Block '" & vReturnArr(4) & "')"

    Dim strErr As String
    strErr = ""

    Select Case vReturnArr(0)
        Case TDT_Unknown:
            strErr = "TDT Connection is in an unknown state. (State: " & vReturnArr(0) & ":Msg " & vReturnArr(1)
            'connectToTDTReportError = True
        Case TDT_ConnectSuccess:
            'connectToTDTReportError = False
        Case TDT_ServerConnectFail:
            strErr = "Connection to the Tank Server failed. " & strGenericErrAppend
            'connectToTDTReportError = True
        Case TDT_TankNotProvided:
            strErr = "Insufficient tank details provided to locate tank. " & strGenericErrAppend
            'connectToTDTReportError = True
        Case TDT_TankInvalidMode:
            strErr = "Connecting with an invalid tank connection mode was attempted. " & strGenericErrAppend
            'connectToTDTReportError = True
        Case TDT_TankConnectFail:
            strErr = "The attempt to connect to the tank failed. " & strGenericErrAppend
            'connectToTDTReportError = True
        Case TDT_BlockNotProvided:
            strErr = "Insufficient block details provided to locate block. " & strGenericErrAppend
            'connectToTDTReportError = True
        Case TDT_BlockConnectFail:
            strErr = "The attempt to connect to the block failed. " & strGenericErrAppend
            'connectToTDTReportError = True
    End Select
    
    connectToTDTReportError = strErr
    
    If Not IsMissing(vUseMessageBox) Then
        If VarType(vUseMessageBox) = vbBoolean Then
            If vUseMessageBox = True Then
                MsgBox strErr
            End If
        End If
    End If
End Function


