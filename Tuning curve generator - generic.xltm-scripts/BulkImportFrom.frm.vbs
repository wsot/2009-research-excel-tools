Attribute VB_Name = "BulkImportFrom"
Attribute VB_Base = "0{201B2E64-06FC-4C73-A0CD-8D1C5234B152}{C200197C-D283-4969-A74D-124FD250C7FA}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Dim theServer As String
Dim theTank As String
Dim theBlock As String

Private Sub SelectAll_Click()
    Call selectAllInList(BlockList)
End Sub

Private Sub DeselectAll_Click()
    Call deselectAllInList(BlockList)
End Sub

'generated a list of epocs from the first selected item in the 'block list'
Private Sub BuildEpocList_Click()

    Dim isOneSelected As Boolean
    Dim i As Integer
    
    'check that at least one option is selected, and identify that option (i)
    For i = 0 To (BlockList.ListCount - 1)
        If BlockList.Selected(i) Then
            isOneSelected = True
            Exit For
        End If
    Next
    
    If isOneSelected Then
        'parse the block and tank name from the info in the list box
        Dim actBlock As String
        Dim actTank As String
        actTank = Left(BlockList.List(i), InStr(BlockList.List(i), ":") - 1)
        actBlock = Right(BlockList.List(i), Len(BlockList.List(i)) - Len(actTank) - 1)
        actTank = bulkImportRootDir & "\" & actTank
        
        'build the option lists
        Call buildOptionLists(theServer, actTank, actBlock, False)
    Else
        MsgBox "Please select a tank/block before loading attempting to load epocs"
    End If

End Sub

Private Sub Cancel_Click()
    doImport = False
    Unload Me        'Unloads the UserForm.
End Sub

Private Sub ImportButton_Click()
    
    If writeSelectedBlocksToSheet() > 0 Then 'will return 0 if no blocks selected, in which case throw an error
        Call writeSelectedEpocsToSheet
        doImport = True
        Unload Me
    Else
        MsgBox ("Please select at least one block to import")
    End If
End Sub

Sub buildOptionLists(sServer As String, sTank As String, sBlock As String, usePrevValues)
    'if a different block is selcted, try to connect to it
    Const EVTYPE_STRON = &H101 'this is a strobe-on event type
    
    Dim objTTX As TTankX
    Dim strErr As String
    'Set objTTX = CreateObject("TTank.X")
    Set objTTX = New TTankX 'don't know if this will work, but it'd be cute if it did
        
    strErr = connectToTDTReportError(connectToTDT(objTTX, False, sServer, sTank, sBlock))
    If strErr = "" Then 'if blank, then no error occurred connecting to TDT
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
            If Worksheets("Variables (do not edit)").Range("B5").Value <> "" Then
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

        Set vOrigOtherGroupings = Nothing
        
        Call objTTX.CloseTank
        Call objTTX.ReleaseServer
        Set objTTX = Nothing
        
    End If
End Sub

Private Sub UserForm_Activate()
    
    Dim objTTX As TTankX
    'Set objTTX = CreateObject("TTank.X") 'establish connection to TDT Tank engine
    Set objTTX = New TTankX
    
    theServer = Worksheets("Variables (do not edit)").Range("B1").Value
    theTank = Worksheets("Variables (do not edit)").Range("B2").Value

    Dim vConnReturn As Variant
    vConnReturn = connectToTDT(objTTX, False, theServer, theTank, theBlock)
    If Not vConnReturn(0) = TDT_ServerConnectFail Then
        Set objTTX = Nothing
        
        Dim objFS As FileSystemObject
        Set objFS = CreateObject("Scripting.FileSystemObject")
        
        Dim objFolder As Folder
        Set objFolder = objFS.GetFolder(bulkImportRootDir)
         
        Dim dBlockList As Dictionary
        Set dBlockList = New Dictionary
        
        Call findAllBlocks(objFS, objFolder, dBlockList, objTTX)
        
        Set objFolder = Nothing
        Set objFS = Nothing
        
        If dBlockList.Count > 0 Then
            Dim vBlocks As Variant
            vBlocks = dBlockList.Keys
        
            BlockList.Clear
            
            Dim i As Integer
            For i = 0 To UBound(vBlocks)
                Call BlockList.AddItem(Right(vBlocks(i), Len(vBlocks(i)) - Len(bulkImportRootDir) - 1), i)
                If InStr(1, LCase(Right(vBlocks(i), Len(vBlocks(i)) - Len(bulkImportRootDir) - 1)), "map") > 0 Then
                    BlockList.Selected(BlockList.ListCount - 1) = True
                End If
            Next
        End If
        
        Set objTTX = CreateObject("TTank.X") 'establish connection to TDT Tank engine
        Call objTTX.ConnectServer(theServer, "Me")
        
        If Worksheets("Variables (do not edit)").Range("E1").Value = 1 Then
            ReverseX.Value = True
        Else
            ReverseX.Value = False
        End If
        
        If Worksheets("Variables (do not edit)").Range("E2").Value = 1 Then
            ReverseY.Value = True
        Else
            ReverseY.Value = False
        End If
    Else
        Call connectToTDTReportError(vConnReturn, True)
        doImport = False
        Unload Me
    End If
End Sub



Sub findAllBlocks(ByRef objFS As FileSystemObject, ByRef objFolder As Folder, ByRef dBlockList As Dictionary, objTTX As TTankX)
    Dim folderList As Folders
    Dim subFolder As Folder
       
    Dim ts As TextStream
    Dim theText As String
    Dim instrRes As Integer
    Dim blnIsTank As Boolean
    
    Dim i As Integer
    Dim strBlockName As String
    
    Set folderList = objFolder.Subfolders
    For Each subFolder In folderList
        blnIsTank = False
        If objFS.FileExists(subFolder & "\desktop.ini") Then
            Set ts = objFS.OpenTextFile(subFolder & "\desktop.ini", 1, False)
            theText = ts.ReadLine
            instrRes = InStr(1, theText, "TDT data tank folder", vbTextCompare)
            If instrRes = 0 Or IsNull(instrRes) Then
                theText = ts.ReadLine
                theText = ts.ReadLine
                theText = ts.ReadLine
                theText = ts.ReadLine
                instrRes = InStr(1, theText, "TDT data tank folder", vbTextCompare)
                If instrRes <> 0 And Not IsNull(instrRes) Then
                    blnIsTank = True
                End If
            Else
                blnIsTank = True
            End If
            Call ts.Close
            
            If blnIsTank Then
                Set objTTX = CreateObject("TTank.X") 'establish connection to TDT Tank engine
                Call objTTX.ConnectServer(theServer, "Me")
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
        '        End If
                Set objTTX = Nothing
            Else
                Call findAllBlocks(objFS, subFolder, dBlockList, objTTX)
            End If
        Else
            Call findAllBlocks(objFS, subFolder, dBlockList, objTTX)
        End If
    Next
    
    Set subFolder = Nothing
    Set folderList = Nothing
End Sub


Function writeSelectedBlocksToSheet()
    Dim iIterA As Integer
    Dim iIterB As Integer

    iIterB = 0

    For iIterA = 0 To (BlockList.ListCount - 1)
        If BlockList.Selected(iIterA) Then
            Worksheets("Variables (do not edit)").Range("N" & CStr(2 + iIterB)).Value = BlockList.List(iIterA)
            iIterB = iIterB + 1
        End If
    Next
    
    Worksheets("Variables (do not edit)").Range("N" & CStr(2 + iIterB)).Value = ""

    writeSelectedBlocksToSheet = iIterB
    
End Function

Function writeSelectedAxisGroups()
        
    Dim iIterA As Integer
    Dim iIterB As Integer
    
    'store the selected 'axis' and other grouping data
    Dim dictOtherEp As Dictionary
    Set dictOtherEp = New Dictionary
    
    'remove the current list of blocks
    iIterA = 9
    While Worksheets("Variables (do not edit)").Range("B" & CStr(iIterA)).Value <> ""
        Worksheets("Variables (do not edit)").Range("B" & CStr(iIterA)).Value = ""
        iIterA = iIterA + 1
    Wend
    
    iIterA = 0
    iIterB = 0
    
    For iIterA = 0 To (OtherGroupings.ListCount - 1)
        If OtherGroupings.Selected(iIterA) Then
            Call dictOtherEp.Add(OtherGroupings.List(iIterA), 1)
            Worksheets("Variables (do not edit)").Range("B" & CStr(9 + iIterB)).Value = OtherGroupings.List(iIterA)
            iIterB = iIterB + 1
        End If
    Next
    Worksheets("Variables (do not edit)").Range("B" & CStr(9 + iIterB)).Value = ""

    Worksheets("Variables (do not edit)").Range("B5").Value = XAxis.Value
    Worksheets("Variables (do not edit)").Range("B6").Value = YAxis.Value
    
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

End Function


Function writeSelectedEpocsToSheet()

        If ReverseX.Value = True Then
            Worksheets("Variables (do not edit)").Range("E1").Value = 1
        Else
            Worksheets("Variables (do not edit)").Range("E1").Value = 0
        End If
        
        If ReverseY.Value = True Then
            Worksheets("Variables (do not edit)").Range("E2").Value = 1
        Else
            Worksheets("Variables (do not edit)").Range("E2").Value = 0
        End If
        
        Worksheets("Variables (do not edit)").Range("B5").Value = XAxis.Value
        Worksheets("Variables (do not edit)").Range("B6").Value = YAxis.Value

        Dim i As Integer
        Dim iOtherItemIndex As Integer
        iOtherItemIndex = 9
        For i = 0 To (OtherGroupings.ListCount - 1)
            If OtherGroupings.Selected(i) Then
                Worksheets("Variables (do not edit)").Range("B" & CStr(iOtherItemIndex)).Value = OtherGroupings.List(i)
                iOtherItemIndex = iOtherItemIndex + 1
            End If
        Next
        Worksheets("Variables (do not edit)").Range("B" & CStr(iOtherItemIndex)).Value = ""

End Function
