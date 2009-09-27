Attribute VB_Name = "BulkImportFrom"
Attribute VB_Base = "0{F8788D04-28D7-4742-B65F-C9D2F259D0EB}{3B5BA4A0-5260-41A3-8138-D02D8D085012}"
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


Private Sub BuildEpocList_Click()

    Dim isOneSelected As Boolean
    Dim i As Integer
    
    For i = 0 To (BlockList.ListCount - 1)
        If BlockList.Selected(i) Then
            isOneSelected = True
            Exit For
        End If
    Next
    
    If isOneSelected Then
        Dim ActBlock As String
        Dim ActTank As String
        ActTank = Left(BlockList.List(i), InStr(BlockList.List(i), ":") - 1)
        ActBlock = Right(BlockList.List(i), Len(BlockList.List(i)) - Len(ActTank) - 1)
        ActTank = bulkImportRootDir & "\" & ActTank
        
        Call buildOptionLists(ActBlock, ActTank, theServer, False)
        'BlockList.List (i)
    Else
        MsgBox "Please select a tank/block before loading attempting to load epocs"
    End If

End Sub

Private Sub Cancel_Click()
    doImport = False
    Unload Me        'Unloads the UserForm.
End Sub

Private Sub ImportButton_Click()
    Dim i As Integer
    Dim j As Integer
    
    Dim dBlocks As Dictionary
    Set dBlocks = New Dictionary
    
    j = 0
    
    For i = 0 To (BlockList.ListCount - 1)
        If BlockList.Selected(i) Then
            Call dBlocks.Add(BlockList.List(i), 1)
            Worksheets("Variables (do not edit)").Range("N" & CStr(2 + j)).Value = BlockList.List(i)
            j = j + 1
        End If
    Next
    
    Worksheets("Variables (do not edit)").Range("N" & CStr(2 + j)).Value = ""
    
    If j > 0 Then
        doImport = True

    
'    If BlockSelect1.ActiveBlock <> "" Then
'        doImport = True
    
        'set global variables to the selected block information
'        theServer = BlockSelect1.UseServer
'        theTank = BlockSelect1.UseTank
'        theBlock = BlockSelect1.ActiveBlock
        
'        Worksheets("Variables (do not edit)").Range("B1").Value = BlockSelect1.UseServer
'        Worksheets("Variables (do not edit)").Range("B2").Value = BlockSelect1.UseTank
'        Worksheets("Variables (do not edit)").Range("B3").Value = BlockSelect1.ActiveBlock
        
        'store the selected 'axis' and other grouping data
        Dim dictOtherEp As Dictionary
        Set dictOtherEp = New Dictionary
        
        Dim iOrigOtherItemIndex As Integer
        iOrigOtherItemIndex = 9
        While Worksheets("Variables (do not edit)").Range("B" & CStr(iOrigOtherItemIndex)).Value <> ""
            Worksheets("Variables (do not edit)").Range("B" & CStr(iOrigOtherItemIndex)).Value = ""
            iOrigOtherItemIndex = iOrigOtherItemIndex + 1
        Wend
        
        j = 0
        
        For i = 0 To (OtherGroupings.ListCount - 1)
            If OtherGroupings.Selected(i) Then
                Call dictOtherEp.Add(OtherGroupings.List(i), 1)
                Worksheets("Variables (do not edit)").Range("B" & CStr(9 + j)).Value = OtherGroupings.List(i)
                j = j + 1
            End If
        Next
        Worksheets("Variables (do not edit)").Range("B" & CStr(9 + j)).Value = ""
    
'        If Not dBlocks.Exists(BlockSelect1.ActiveBlock) Then
'            Call dBlocks.Add(BlockSelect1.ActiveBlock, 1)
'            Worksheets("Variables (do not edit)").Range("N" & CStr(2 + j)).Value = BlockSelect1.ActiveBlock
'            j = j + 1
'        End If
    
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


Private Sub TankSelect1_TankChanged(ActTank As String, ActServer As String)
    'When a different tank is selected, test if a connection can be made
    Select Case testSettings(ActServer, ActTank, "")
        Case ConnectSuccess:
            'if so update the list of available blocks for the tank
            BlockSelect1.UseServer = ActServer
            BlockSelect1.UseTank = ActTank
            Call BlockSelect1.Refresh
            Call buildBlockList(TankSelect1.ActiveTank)
        Case BlockConnectFail:
            'if so update the list of available blocks for the tank
            BlockSelect1.UseServer = ActServer
            BlockSelect1.UseTank = ActTank
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

    
    If objTTX.ConnectServer(theServer, "Me") = CLng(1) Then
        
        Dim objFS As FileSystemObject
        Set objFS = CreateObject("Scripting.FileSystemObject")
        
        Dim objFolder As Folder
        Set objFolder = objFS.GetFolder(bulkImportRootDir)
         
        Dim dBlockList As Dictionary
        Set dBlockList = New Dictionary
        
        Call findAllBlocks(objFS, objFolder, dBlockList)
        
        Set objFolder = Nothing
        Set objFS = Nothing
        
        If dBlockList.Count > 0 Then
            Dim vBlocks As Variant
            vBlocks = dBlockList.Keys
        
            BlockList.Clear
            
            Dim i As Integer
            For i = 0 To UBound(vBlocks)
                Call BlockList.AddItem(Right(vBlocks(i), Len(vBlocks(i)) - Len(bulkImportRootDir) - 1), i)
            Next
        End If
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

Private Sub BlockSelect1_BlockChanged(ActBlock As String, ActTank As String, ActServer As String)
    Call buildOptionLists(ActBlock, ActTank, ActServer, False)
End Sub

'test the connection settings to see if it is possible to connect to the server/tank/block
Function testSettings(ActServer, ActTank, ActBlock)
    
    If objTTX.ConnectServer(ActServer, "Me") <> CLng(1) Then
        testSettings = ServerConnectFail
        Exit Function
    ElseIf objTTX.OpenTank(ActTank, "R") <> CLng(1) Then
        objTTX.ReleaseServer
        testSettings = TankConnectFail
        Exit Function
    ElseIf objTTX.SelectBlock(ActBlock) <> CLng(1) Then
        objTTX.CloseTank
        objTTX.ReleaseServer
        testSettings = BlockConnectFail
    End If
    
End Function

Sub buildOptionLists(ActBlock, ActTank, ActServer, usePrevValues)
    'if a different block is selcted, try to connect to it
    Const EVTYPE_STRON = &H101
    
    Dim objTTX As Object
    Set objTTX = CreateObject("TTank.X")
    
    If objTTX.ConnectServer(ActServer, "Me") <> CLng(1) Then
        MsgBox ("Connecting to server " & theServer & " failed.")
        Exit Sub
    End If
    
    If objTTX.OpenTank(ActTank, "R") <> CLng(1) Then
        MsgBox ("Connecting to tank " & theTank & " on server " & theServer & " failed .")
        Call objTTX.ReleaseServer
        Exit Sub
    End If
    
    If objTTX.SelectBlock(ActBlock) <> CLng(1) Then
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

Sub findAllBlocks(ByRef objFS As FileSystemObject, ByRef objFolder As Folder, ByRef dBlockList As Dictionary)
    Dim folderList As Folders
    Dim subFolder As Folder
    
    Dim i As Integer
    Dim strBlockName As String
    
    Set folderList = objFolder.Subfolders
    For Each subFolder In folderList
        Call objTTX.OpenTank(subFolder.Path, "R")
        'If objTTX.OpenTank(subFolder.Path, "R") = CLng(1) Then
            i = 0
            Do
                strBlockName = objTTX.QueryBlockName(i)
                If strBlockName <> "" Then
                    If Not dBlockList.Exists(subFolder.Path & ":" & strBlockName) Then
                        Call dBlockList.Add(subFolder.Path & ":" & strBlockName, 0)
                    End If
                Else
                    Exit Do
                End If
                i = i + 1
            Loop
            Call findAllBlocks(objFS, subFolder, dBlockList)
'        End If
    Next
    
    Set subFolder = Nothing
    Set folderList = Nothing
End Sub
