Attribute VB_Name = "ImportFrom"
Attribute VB_Base = "0{09904E73-14F3-4B07-B696-23AB5951DB98}{3EFC56DF-A5A7-4C47-8191-0D79BEFA40AB}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit

Dim objTTX As Object
Const ConnectSuccess = 1
Const ServerConnectFail = 1
Const TankConnectFail = 2
Const BlockConnectFail = 3

Private Sub Cancel_Click()
    doImport = False
    Unload Me        'Unloads the UserForm.
End Sub

Private Sub ImportButton_Click()
    If BlockSelect1.ActiveBlock <> "" Then
        doImport = True
    
        'set global variables to the selected block information
        theServer = BlockSelect1.UseServer
        theTank = BlockSelect1.UseTank
        theBlock = BlockSelect1.ActiveBlock
        
        Worksheets("Variables (do not edit)").Range("B1").Value = BlockSelect1.UseServer
        Worksheets("Variables (do not edit)").Range("B2").Value = BlockSelect1.UseTank
        Worksheets("Variables (do not edit)").Range("B3").Value = BlockSelect1.ActiveBlock
        
        'store the selected 'axis' and other grouping data
        Dim dictOtherEp As Dictionary
        Set dictOtherEp = New Dictionary
        
        Dim dBlocks As Dictionary
        Set dBlocks = New Dictionary
        
        Dim i As Integer
        Dim j As Integer
        
        Dim iOrigOtherItemIndex As Integer
        iOrigOtherItemIndex = 9
        While Worksheets("Variables (do not edit)").Range("B" & CStr(iOrigOtherItemIndex)).Value <> ""
            Worksheets("Variables (do not edit)").Range("B" & CStr(iOrigOtherItemIndex)).Value = ""
            iOrigOtherItemIndex = iOrigOtherItemIndex + 1
        Wend
        
        For i = 0 To (OtherGroupings.ListCount - 1)
            If OtherGroupings.Selected(i) Then
                Call dictOtherEp.Add(OtherGroupings.List(i), 1)
                Worksheets("Variables (do not edit)").Range("B" & CStr(9 + j)).Value = OtherGroupings.List(i)
                j = j + 1
            End If
        Next
        Worksheets("Variables (do not edit)").Range("B" & CStr(9 + j)).Value = ""

        j = 0
    
        For i = 0 To (BlockList.ListCount - 1)
            If BlockList.Selected(i) Then
                Call dBlocks.Add(BlockList.List(i), 1)
                Worksheets("Variables (do not edit)").Range("N" & CStr(2 + j)).Value = BlockList.List(i)
                j = j + 1
            End If
        Next
        If Not dBlocks.Exists(BlockSelect1.ActiveBlock) Then
            Call dBlocks.Add(BlockSelect1.ActiveBlock, 1)
            Worksheets("Variables (do not edit)").Range("N" & CStr(2 + j)).Value = BlockSelect1.ActiveBlock
            j = j + 1
        End If
        
        Worksheets("Variables (do not edit)").Range("N" & CStr(2 + j)).Value = ""
    
        xAxisEp = XAxis.Value
        Worksheets("Variables (do not edit)").Range("B5").Value = xAxisEp
        yAxisEp = YAxis.Value
        Worksheets("Variables (do not edit)").Range("B6").Value = yAxisEp
        arrOtherEp = dictOtherEp.Keys
        
        If ReverseX.Value = True Then
            bReverseX = True
            Worksheets("Variables (do not edit)").Range("E1").Value = 1
        Else
            bReverseX = False
            Worksheets("Variables (do not edit)").Range("E1").Value = 0
        End If
        
        If ReverseY.Value = True Then
            bReverseY = True
            Worksheets("Variables (do not edit)").Range("E2").Value = 1
        Else
            bReverseY = False
            Worksheets("Variables (do not edit)").Range("E2").Value = 0
        End If
        
        Unload Me
    Else
        MsgBox ("Please select a block to import")
    End If
End Sub


Private Sub TankSelect1_TankChanged(actTank As String, ActServer As String)
    'When a different tank is selected, test if a connection can be made
    Select Case testSettings(ActServer, actTank, "")
        Case ConnectSuccess:
            'if so update the list of available blocks for the tank
            BlockSelect1.UseServer = ActServer
            BlockSelect1.UseTank = actTank
            Call BlockSelect1.Refresh
            Call buildBlockList(TankSelect1.ActiveTank)
        Case BlockConnectFail:
            'if so update the list of available blocks for the tank
            BlockSelect1.UseServer = ActServer
            BlockSelect1.UseTank = actTank
            Call BlockSelect1.Refresh
            Call buildBlockList(TankSelect1.ActiveTank)
    End Select

'    BlockSelect1.UseServer = ActServer
'    BlockSelect1.UseTank = ActTank
'    Call BlockSelect1.Refresh
'    Call buildBlockList(TankSelect1.ActiveTank)
End Sub

Private Sub UserForm_Activate()
    
    Set objTTX = CreateObject("TTank.X") 'establish connection to TDT Tank engine

    'when the form loads, if tanks etc were already selected then re-select them
    If theServer <> "" Then
        TankSelect1.UseServer = theServer
        BlockSelect1.UseServer = theServer
        If theTank <> "" Then
            TankSelect1.ActiveTank = theTank
            BlockSelect1.UseTank = theTank
            If theBlock <> "" Then
                BlockSelect1.ActiveBlock = theBlock
                Call buildOptionLists(theBlock, theTank, theServer, True)
            End If
            BlockSelect1.Refresh
        End If
        TankSelect1.Refresh
    End If
    
    'try to read parameters from the spreadsheet variables
    If theServer = "" Then
        theServer = Worksheets("Variables (do not edit)").Range("B1").Value
        theTank = Worksheets("Variables (do not edit)").Range("B2").Value
        theBlock = Worksheets("Variables (do not edit)").Range("B3").Value
        Select Case testSettings(theServer, theTank, theBlock)
            Case ConnectSuccess:
                TankSelect1.UseServer = theServer
                TankSelect1.ActiveTank = theTank
                BlockSelect1.UseServer = theServer
                BlockSelect1.UseTank = theTank
                BlockSelect1.ActiveBlock = theBlock
                TankSelect1.Refresh
                BlockSelect1.Refresh
                Call buildOptionLists(theBlock, theTank, theServer, True)
            Case BlockConnectFail:
                TankSelect1.UseServer = theServer
                TankSelect1.ActiveTank = theTank
                TankSelect1.Refresh
        End Select
    End If
    
    If bReverseX = True Or Worksheets("Variables (do not edit)").Range("E1").Value = 1 Then
        ReverseX.Value = True
    Else
        ReverseX.Value = False
    End If
    
    If bReverseY = True Or Worksheets("Variables (do not edit)").Range("E2").Value = 1 Then
        ReverseY.Value = True
    Else
        ReverseY.Value = False
    End If
    
End Sub

Private Sub BlockSelect1_BlockChanged(actBlock As String, actTank As String, ActServer As String)
    Call buildOptionLists(actBlock, actTank, ActServer, False)
End Sub

'test the connection settings to see if it is possible to connect to the server/tank/block
Function testSettings(ActServer, actTank, actBlock)
    
    If objTTX.ConnectServer(ActServer, "Me") <> CLng(1) Then
        testSettings = ServerConnectFail
        Exit Function
    ElseIf objTTX.OpenTank(actTank, "R") <> CLng(1) Then
        objTTX.ReleaseServer
        testSettings = TankConnectFail
        Exit Function
    ElseIf objTTX.SelectBlock(actBlock) <> CLng(1) Then
        objTTX.CloseTank
        objTTX.ReleaseServer
        testSettings = BlockConnectFail
    End If
    
End Function

Sub buildOptionLists(actBlock, actTank, ActServer, usePrevValues)
    'if a different block is selcted, try to connect to it
    Const EVTYPE_STRON = &H101
    
    Dim objTTX As Object
    Set objTTX = CreateObject("TTank.X")
    
    If objTTX.ConnectServer(ActServer, "Me") <> CLng(1) Then
        MsgBox ("Connecting to server " & theServer & " failed.")
        Exit Sub
    End If
    
    If objTTX.OpenTank(actTank, "R") <> CLng(1) Then
        MsgBox ("Connecting to tank " & theTank & " on server " & theServer & " failed .")
        Call objTTX.ReleaseServer
        Exit Sub
    End If
    
    If objTTX.SelectBlock(actBlock) <> CLng(1) Then
        MsgBox ("Connecting to block " & theBlock & " in tank " & theTank & " on server " & theServer & " failed.")
        Call objTTX.CloseTank
        Call objTTX.ReleaseServer
        Exit Sub
    End If
    
    'build a list of all event codes
    Dim arrEventCodes() As Long
    
    arrEventCodes = objTTX.GetEventCodes(EVTYPE_STRON)
    
    'fill the select boxes with the event lists
    Dim i As Integer
    
    Dim sOrigXAxis As String
    Dim sOrigYAxis As String
    Dim vOrigOtherGroupings As Dictionary
    Set vOrigOtherGroupings = New Dictionary
        
    If usePrevValues Then
        If xAxisEp <> "" Then
            sOrigXAxis = xAxisEp
            sOrigYAxis = yAxisEp
        ElseIf Worksheets("Variables (do not edit)").Range("B5").Value <> "" Then
            sOrigXAxis = Worksheets("Variables (do not edit)").Range("B5").Value
            sOrigYAxis = Worksheets("Variables (do not edit)").Range("B6").Value
        End If

        Dim iOrigOtherItemIndex As Integer
        iOrigOtherItemIndex = 9
        While Worksheets("Variables (do not edit)").Range("B" & CStr(iOrigOtherItemIndex)).Value <> ""
            If Not vOrigOtherGroupings.Exists(Worksheets("Variables (do not edit)").Range("B" & CStr(iOrigOtherItemIndex)).Value) Then
                Call vOrigOtherGroupings.Add(Worksheets("Variables (do not edit)").Range("B" & CStr(iOrigOtherItemIndex)).Value, 1)
            End If
            iOrigOtherItemIndex = iOrigOtherItemIndex + 1
        Wend
    Else
        sOrigXAxis = XAxis.Value
        sOrigYAxis = YAxis.Value
    
        For i = 0 To (OtherGroupings.ListCount - 1)
            If OtherGroupings.Selected(i) Then
                Call vOrigOtherGroupings.Add(OtherGroupings.List(i), 1)
            End If
        Next
    End If
    
    Dim bMatchXAxis As Boolean
    bMatchXAxis = False
    Dim bMatchYAxis As Boolean
    bMatchYAxis = False
    
    Call XAxis.Clear
    Call YAxis.Clear
    Call OtherGroupings.Clear
       
    For i = 0 To UBound(arrEventCodes)
        Call XAxis.AddItem(objTTX.CodeToString(arrEventCodes(i)), i)
        
        If bMatchXAxis = False And objTTX.CodeToString(arrEventCodes(i)) = "Frq1" Then 'if no item was selected, choose Frq1 as default
            XAxis.Value = "Frq1"
            bMatchXAxis = True
        ElseIf CStr(objTTX.CodeToString(arrEventCodes(i))) = CStr(sOrigXAxis) Then 'if item was selected before changing blocks, keep same name selected
            XAxis.Value = CStr(sOrigXAxis)
            bMatchXAxis = True
        End If
        Call YAxis.AddItem(objTTX.CodeToString(arrEventCodes(i)), i)
        If bMatchYAxis = False And objTTX.CodeToString(arrEventCodes(i)) = "Lev1" Then 'if no item previously selected, choose Lev1 as default
            YAxis.Value = "Lev1"
            bMatchYAxis = True
        ElseIf CStr(objTTX.CodeToString(arrEventCodes(i))) = CStr(sOrigYAxis) Then 'if item was previously selected, try to reselect it
            YAxis.Value = CStr(sOrigYAxis)
            bMatchYAxis = True
        End If
        Call OtherGroupings.AddItem(objTTX.CodeToString(arrEventCodes(i)), i)
        If vOrigOtherGroupings.Exists(objTTX.CodeToString(arrEventCodes(i))) Then
            OtherGroupings.Selected(i) = True
        End If
    Next
    
    'add the channel option, as it is not actually an epoch
    Call XAxis.AddItem("Channel", i)
    If CStr(sOrigXAxis) = "Channel" Then
        XAxis.Value = "Channel"
    End If
    Call YAxis.AddItem("Channel", i)
    If CStr(sOrigYAxis) = "Channel" Then
        YAxis.Value = "Channel"
    End If

    Call OtherGroupings.AddItem("Channel", i)
    If vOrigOtherGroupings.Exists("Channel") Then
        OtherGroupings.Selected(i) = True
    End If

    'if the defaults were not available, and nothing was selected, choose the first items by default
    If bMatchXAxis = False Then
        XAxis.Value = XAxis.List(0, 0)
    End If
    If bMatchYAxis = False Then
        YAxis.Value = YAxis.List(0, 0)
    End If

    Call objTTX.CloseTank
    Call objTTX.ReleaseServer
    
    Set vOrigOtherGroupings = Nothing
    
End Sub
Sub buildBlockList(tankPath)

    Dim dBlocks As Dictionary
    Set dBlocks = New Dictionary
    Dim vBlocks

    Dim objFS As FileSystemObject
    Dim objFolder As Folder
    Set objFS = CreateObject("Scripting.FileSystemObject")
'   fso.GetParentFolderName(docFullName)
    Set objFolder = objFS.GetFolder(tankPath)

    Dim Subfolders As Folders
    Dim objSubfolder As Folder

    Set Subfolders = objFolder.Subfolders
    For Each objSubfolder In Subfolders
        If objSubfolder.Name <> "TempBlk" Then
            If Not dBlocks.Exists(objSubfolder.Name) Then
                 Call dBlocks.Add(objSubfolder.Name, 0)
            End If
        End If
    Next

    Set objFolder = Nothing
    Set objSubfolder = Nothing
    Set objFS = Nothing
    
    vBlocks = dBlocks.Keys

    BlockList.Clear
    
    Dim i As Integer
    For i = 0 To UBound(vBlocks)
        Call BlockList.AddItem(vBlocks(i), i)
    Next
End Sub

