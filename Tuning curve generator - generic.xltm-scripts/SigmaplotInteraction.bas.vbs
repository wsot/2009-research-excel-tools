Attribute VB_Name = "SigmaplotInteraction"
Global dHeadingList As Dictionary
Global dHeadingsSelected As Dictionary
Global bXAxisLog As Boolean

Global plotWhichSheet As String
Global isFirstChart As Boolean
Global SigmaPlotHandle As Variant

Option Explicit
Sub findSigmplotWindow()
        Dim iRet As Long
        Dim lWindHandle As Long
        Dim oDynWrap As Variant
        Set oDynWrap = CreateObject("DynamicWrapper")
        iRet = oDynWrap.Register("user32.dll", "FindWindowA", "i=ss", "f=s", "r=l")
        lWindHandle = oDynWrap.FindWindowA(vbNullString, "SigmaPlot") 'find the SigmaPlot window
        SigmaPlotHandle = lWindHandle
        Set oDynWrap = Nothing
End Sub

Sub trySigmaplotSave(saveFilename As String, SPApp)
    plotWorkbook.ActiveSheet.Range("A25").Value = "'Interacting with SigmaPlot"
    Dim iRetries As Integer
    Dim z As Integer
    Dim filenameParts
    Dim wrongFilename As String
    wrongFilename = ""
    
    Dim objFS As FileSystemObject
    Set objFS = CreateObject("Scripting.FileSystemObject")

    If saveFilename <> "" Then
        If useSendKeys Then
            filenameParts = Split(saveFilename, "\")
    
            For iRetries = 0 To 3
                'SPApp.ActiveDocument.SaveAs (saveFilename)
                'Call SPApp.ActiveDocument.Close(True, saveFilename)
                'Call SPApp.ActiveDocument.Close(False)
                'Call spNB.Close(True, saveFilename)
    
                If objFS.FileExists(saveFilename) Then
                    Call objFS.DeleteFile(saveFilename, True)
                End If
        
                If isFirstChart Then
                    isFirstChart = False
                    Call AppActivate("SigmaPlot", 1)
                Else
                    On Error Resume Next
                    Call AppActivate("SigmaPlot", 0)
                    On Error GoTo 0
                End If
                delayMe (5)
                Call SendKeys("{F12}", 1)
                delayMe (5)
                For z = 0 To UBound(filenameParts)
                    If z < UBound(filenameParts) Then
                        Call SendKeys(filenameParts(z) & "\", 1)
                    Else
                        Call SendKeys(filenameParts(z), 1)
                    End If
                    delayMe (1)
                    Call SendKeys("{RIGHT}", 1)
                    delayMe (1)
                    Call SendKeys("{ENTER}", 1)
                    delayMe (1)
                Next
                
                'check it saved with the correct filename
                If LCase(SPApp.ActiveDocument.FullName) <> LCase(saveFilename) Then
                    If LCase(SPApp.ActiveDocument.FullName) <> "" And Left(LCase(SPApp.ActiveDocument.FullName), 8) <> "notebook" Then
                        If wrongFilename <> "" And wrongFilename <> LCase(SPApp.ActiveDocument.FullName) Then 'still not the right filename, but a different one - delete the old one
                            Call objFS.DeleteFile(wrongFilename, True)
                        End If
                        wrongFilename = LCase(SPApp.ActiveDocument.FullName)
                    End If
                    delayMe (5)
                    Call SendKeys("{ESC}", 1)
                Else
                    If wrongFilename <> "" Then 'wrote a wrong filename first - delete it
                        Call objFS.DeleteFile(wrongFilename, True)
                    End If
                    delayMe (2)
                    Call SendKeys("^+{F4}", 1)
                    delayMe (2)
                    Call SendKeys("^+{F4}", 1)
                    delayMe (2)
                    Call SendKeys("^+{F4}", 1)
                    Exit For
                End If
            Next
        Else
        
            If objFS.FileExists(saveFilename) Then
                Call objFS.DeleteFile(saveFilename, True)
            End If
            
            Dim iRet
            Dim lWindHandle
            Dim lDialogHandle
            Dim lButtonHandle
            Const WM_LBUTTONDOWN = &H201
            Const WM_LBUTTONUP = &H201
            Const WM_KEYDOWN = &H100
            Const WM_KEYUP = &H101
            
            Const WM_COMMAND = &H111
            
            Const WM_USER = &H400
            Const WMTRAY_TOGGLEQL = (WM_USER + 237)
            Const BM_CLICK = &HF5
                
            Const VK_ENTER = &HD
            Dim oDynWrap As Variant
            
            Set oDynWrap = CreateObject("DynamicWrapper")
            iRet = oDynWrap.Register("user32.dll", "FindWindowA", "i=ss", "f=s", "r=l")
            iRet = oDynWrap.Register("USER32.DLL", "PostMessageA", "i=hlll", "f=s", "r=l")
            iRet = oDynWrap.Register("USER32.DLL", "SendMessageA", "i=hlll", "f=s", "r=l")
            iRet = oDynWrap.Register("USER32.DLL", "SetForegroundWindow", "i=h", "f=s", "r=l")
            iRet = oDynWrap.Register("USER32.DLL", "FindWindowEx", "i=hhss", "f=s", "r=l")
                   
            'iRet = oDynWrap.FindWindowA("Afx:00400000:8:00010003:00000000:03F50C6B", vbNullString)
            'lWindHandle = oDynWrap.FindWindowA("Afx:00400000:8:00010017:00000000:00010460", vbNullString) 'find the SigmaPlot window
            If IsNull(SigmaPlotHandle) Then
                lWindHandle = oDynWrap.FindWindowA(vbNullString, "SigmaPlot") 'find the SigmaPlot window
            Else
                lWindHandle = SigmaPlotHandle
            End If
            '    lWindHandle = oDynWrap.FindWindowA("#32770", vbNullString)
            If IsNull(lWindHandle) Then
                MsgBox ("Trying to attach to the SigmaPlot window failed: " & lWindHandle)
            Else
                iRet = oDynWrap.PostMessageA(lWindHandle, WM_COMMAND, MAKELPARAM(57604, 1), 0&) 'send the 'save as' command
                'iRet = oDynWrap.SendMessageA(lWindHandle, WM_COMMAND, MAKELPARAM(57604, 1), 0&)
                delayMe (2)
                lDialogHandle = oDynWrap.FindWindowA("#32770", "Save As") 'get the dialog box
                If IsNull(lDialogHandle) Then
                   MsgBox ("Trying to attach to the Save dialog failed: " & lDialogHandle)
                Else
                    lButtonHandle = oDynWrap.FindWindowEx(lDialogHandle, 0&, vbNullString, "&Save") 'get the save button
                    delayMe (2)
                    If IsNull(lButtonHandle) Then
                        MsgBox ("Trying to attach to the Save button failed: " & lDialogHandle)
                    Else
                        iRet = oDynWrap.SendMessageA(lButtonHandle, BM_CLICK, 0&, 0&)
                        delayMe (2)
                        
                        wrongFilename = SPApp.ActiveDocument.FullName
                        
                        iRet = oDynWrap.SendMessageA(lWindHandle, WM_COMMAND, MAKELPARAM(780, 0), 0&) 'send the 'close all notebooks' command
                        delayMe (2)
                        
                        Call objFS.MoveFile(wrongFilename, saveFilename)
                        
                    End If
                End If
            End If
            Set oDynWrap = Nothing
        End If
    End If
    plotWorkbook.ActiveSheet.Range("A25").Value = ""
    Set objFS = Nothing
End Sub


Sub transferCandidatesToSigmaplot(saveFilename As String)
    plotWhichSheet = plotWorkbook.Worksheets("Settings").Range("B18").Value

'    Dim zOffsetSize As Long
'    Dim iColOffset As Integer
'    Dim iRowOffset As Integer

'    zOffsetSize = plotWorkbook.Worksheets("Variables (do not edit)").Range("H3").Value
'    iColOffset = plotWorkbook.Worksheets("Variables (do not edit)").Range("H5").Value
'    iRowOffset = plotWorkbook.Worksheets("Variables (do not edit)").Range("H6").Value

'    Dim xPos As Long
'    Dim yPos As Long
   
'    xPos = iColOffset + 1
'    yPos = iRowOffset
    
    Dim iRow As Integer
    iRow = 2
    
    Set dHeadingsSelected = New Dictionary
    
    Do
        If plotWorkbook.Worksheets("Likely tuned channels").Cells(iRow, 1).Value <> "" Then
            If Not dHeadingsSelected.Exists(plotWorkbook.Worksheets("Likely tuned channels").Cells(iRow, 1).Value) Then
                Call dHeadingsSelected.Add(plotWorkbook.Worksheets("Likely tuned channels").Cells(iRow, 1).Value, 0)
            End If
            iRow = iRow + 1
        Else
            Exit Do
        End If
    Loop
    
    If doImport Then
        Call TransferToSigmaplot(saveFilename)
    End If
End Sub
Sub transferAllToSigmaplot(saveFilename As String)
    plotWhichSheet = plotWorkbook.Worksheets("Settings").Range("B18").Value

    Dim zOffsetSize As Long
    Dim iColOffset As Integer
    Dim iRowOffset As Integer

    zOffsetSize = plotWorkbook.Worksheets("Variables (do not edit)").Range("H3").Value
    iColOffset = plotWorkbook.Worksheets("Variables (do not edit)").Range("H5").Value
    iRowOffset = plotWorkbook.Worksheets("Variables (do not edit)").Range("H6").Value

    Dim xPos As Long
    Dim yPos As Long
   
    xPos = iColOffset + 1
    yPos = iRowOffset
    
    Set dHeadingsSelected = New Dictionary
    
    Do
        If plotWorkbook.Worksheets(plotWhichSheet).Cells(yPos, xPos).Value <> "" Then
            If Not dHeadingsSelected.Exists(plotWorkbook.Worksheets(plotWhichSheet).Cells(yPos, xPos).Value) Then
                Call dHeadingsSelected.Add(plotWorkbook.Worksheets(plotWhichSheet).Cells(yPos, xPos).Value, 0)
            End If
            yPos = yPos + zOffsetSize
        Else
            Exit Do
        End If
    Loop
    
    If doImport Then
        Call TransferToSigmaplot(saveFilename)
    End If
End Sub
Sub TransferToSigmaplot(saveFilename As String)

    Dim xCount As Long
    Dim yCount As Long
    Dim zOffsetSize As Long
    Dim lMaxHistHeight As Long
    Dim iColOffset As Integer
    Dim iRowOffset As Integer

    xCount = plotWorkbook.Worksheets("Variables (do not edit)").Range("H1").Value
    yCount = plotWorkbook.Worksheets("Variables (do not edit)").Range("H2").Value
    zOffsetSize = plotWorkbook.Worksheets("Variables (do not edit)").Range("H3").Value
    If plotWhichSheet = "Means" Then
        lMaxHistHeight = plotWorkbook.Worksheets("Variables (do not edit)").Range("H7").Value
    Else
        lMaxHistHeight = plotWorkbook.Worksheets("Variables (do not edit)").Range("H4").Value
    End If
    iColOffset = plotWorkbook.Worksheets("Variables (do not edit)").Range("H5").Value
    iRowOffset = plotWorkbook.Worksheets("Variables (do not edit)").Range("H6").Value
    
    Const SAA_FROMVAL = &H406
    Const SAA_TOVAL = &H407
    Const GPM_SETPLOTATTR = &H301
    Const GPM_SETAXISATTR = 1025
    Const SLA_SELECTDIM = 776
    Const SAA_OPTIONS = 1027
    Const GPM_SETAXISATTRSTRING = 1032
    Const SEA_COLORCOL = 1557
    Const SEA_COLORREPEAT = 1555
    Const SLA_CONTOURFILLTYPE = 856
    Const SAA_SELECTLINE = 1034
    Const SEA_THICKNESS = 1537
    Const SEA_COLOR = 1542
    Const SAA_SUB1OPTIONS = 1040

    Dim SPApp As Object
    Set SPApp = CreateObject("SigmaPlot.Application.1")
    SPApp.Visible = True
    Call SPApp.Application.Notebooks.Add
    
    Dim i As Integer
    Dim j As Long
    Dim k As Long

    Dim xPos As Long
    Dim yPos As Long

    Dim spNB As Object
    Dim spWS As Object
    Dim spDT As Object
    Dim spGRPH As Object
    
    Dim strTitle
    
    xPos = iColOffset + 1
    yPos = iRowOffset
    
    Do
        If plotWorkbook.Worksheets(plotWhichSheet).Cells(yPos, xPos).Value <> "" Then
            If dHeadingsSelected.Exists(plotWorkbook.Worksheets(plotWhichSheet).Cells(yPos, xPos).Value) Then
                strTitle = plotWorkbook.Worksheets(plotWhichSheet).Cells(yPos, xPos).Value
                Set spNB = SPApp.Notebooks.Item(SPApp.Notebooks.Count - 1)
                Set spWS = spNB.NotebookItems.Item(spNB.NotebookItems.Count - 1)
                spWS.Name = plotWorkbook.Worksheets(plotWhichSheet).Cells(yPos, xPos).Value
                Set spDT = spWS.DataTable
                
                yPos = yPos
                            
                For j = 0 To (xCount - 1)
                    spDT.Cell(0, j) = plotWorkbook.Worksheets(plotWhichSheet).Cells(yPos + 1, xPos + j + 1).Value
                Next
                
                For j = 0 To (yCount - 1)
                    spDT.Cell(1, j) = plotWorkbook.Worksheets(plotWhichSheet).Cells(yPos + j + 2, xPos).Value
                Next
                
                
                For j = 0 To (xCount - 1)
                    For k = 0 To (yCount - 1)
                        spDT.Cell(3 + k, j) = plotWorkbook.Worksheets(plotWhichSheet).Cells(yPos + k + 2, xPos + j + 1).Value
                    Next
                Next
                
                spDT.Cell(2, 0) = "@rgb(255,255,255)"
                spDT.Cell(2, 1) = "@rgb(0,0,255)"
                spDT.Cell(2, 2) = "@rgb(0,255,255)"
                spDT.Cell(2, 3) = "@rgb(0,255,0)"
                spDT.Cell(2, 4) = "@rgb(255,255,0)"
                spDT.Cell(2, 5) = "@rgb(255,0,0)"
                    
                'Call spNB.NotebookItems.Add(2)
                'Set spGRPH = spNB.NotebookItems.Item(spNB.NotebookItems.Count - 1)
        
                Call SPApp.ActiveDocument.NotebookItems.Add(2)
                Dim ColumnsPerPlot(2, 3)
                ColumnsPerPlot(0, 0) = 0
                ColumnsPerPlot(1, 0) = 0
                ColumnsPerPlot(2, 0) = 31999999
                ColumnsPerPlot(0, 1) = 1
                ColumnsPerPlot(1, 1) = 0
                ColumnsPerPlot(2, 1) = 31999999
                ColumnsPerPlot(0, 2) = 3
                ColumnsPerPlot(1, 2) = 0
                ColumnsPerPlot(2, 2) = 31999999
                ColumnsPerPlot(0, 3) = 3 + (yCount - 1)
                ColumnsPerPlot(1, 3) = 0
                ColumnsPerPlot(2, 3) = 31999999
                
                Dim PlotColumnCountArray()
                ReDim PlotColumnCountArray(0)
                
                PlotColumnCountArray(0) = 4
                Call SPApp.ActiveDocument.CurrentPageItem.CreateWizardGraph("Contour Plot", "Filled Contour Plot", "XY Many Z", ColumnsPerPlot, PlotColumnCountArray, "Worksheet Columns", "Standard Deviation", "Degrees", 0#, 360#, , "Standard Deviation", True)
                Call SPApp.ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).SelectObject
            
                SPApp.ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Name = strTitle
                SPApp.ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Axes(0).Name = "Attenuation"
                SPApp.ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Axes(1).Name = "Frequency"
                
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, 3)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_OPTIONS, 10)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_OPTIONS, 51380225)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_OPTIONS, 12583938)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTRSTRING, SAA_FROMVAL, "0")
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTRSTRING, SAA_TOVAL, CStr(lMaxHistHeight))
                
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, 3)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, 2)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_THICKNESS, 10)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, 4)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_COLOR, &HFFFFFF)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_COLORCOL, 2)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_COLORREPEAT, 4)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_CONTOURFILLTYPE, 1)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, 5)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_COLOR, &HFFFFFF)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_COLORCOL, 2)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_COLORREPEAT, 4)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, 4)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SUB1OPTIONS, 1298)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SUB1OPTIONS, 3889)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, 1)
        
                Call SPApp.ActiveDocument.NotebookItems.Add(2)
                ColumnsPerPlot(0, 0) = 0
                ColumnsPerPlot(1, 0) = 0
                ColumnsPerPlot(2, 0) = 31999999
                ColumnsPerPlot(0, 1) = 1
                ColumnsPerPlot(1, 1) = 0
                ColumnsPerPlot(2, 1) = 31999999
                ColumnsPerPlot(0, 2) = 3
                ColumnsPerPlot(1, 2) = 0
                ColumnsPerPlot(2, 2) = 31999999
                ColumnsPerPlot(0, 3) = 3 + (yCount - 1)
                ColumnsPerPlot(1, 3) = 0
                ColumnsPerPlot(2, 3) = 31999999
                
                ReDim PlotColumnCountArray(0)
                
                PlotColumnCountArray(0) = 4
                Call SPApp.ActiveDocument.CurrentPageItem.CreateWizardGraph("Contour Plot", "Filled Contour Plot", "XY Many Z", ColumnsPerPlot, PlotColumnCountArray, "Worksheet Columns", "Standard Deviation", "Degrees", 0#, 360#, , "Standard Deviation", True)
                Call SPApp.ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).SelectObject
            
                SPApp.ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Name = "Site y"
                SPApp.ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Axes(0).Name = "Attenuation"
                SPApp.ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Axes(1).Name = "Frequency"
                        
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, 3)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, 2)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_THICKNESS, 10)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, 4)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_COLOR, &HFFFFFF)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_COLORCOL, 2)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_COLORREPEAT, 4)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_CONTOURFILLTYPE, 1)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, 5)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_COLOR, &HFFFFFF)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_COLORCOL, 2)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_COLORREPEAT, 4)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, 4)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SUB1OPTIONS, 1298)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SUB1OPTIONS, 3889)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, 1)
        
                Call spNB.NotebookItems.Add(1)
            End If
            yPos = yPos + zOffsetSize
        Else
            Exit Do
        End If
    Loop

    Call trySigmaplotSave(saveFilename, SPApp)
    
End Sub
Sub transferToSigmaplotButton()
    plotWhichSheet = plotWorkbook.Worksheets("Settings").Range("B18").Value
    Set plotWorkbook = Application.ActiveWorkbook

    Dim zOffsetSize As Long
    Dim iColOffset As Integer
    Dim iRowOffset As Integer

    zOffsetSize = plotWorkbook.Worksheets("Variables (do not edit)").Range("H3").Value
    iColOffset = plotWorkbook.Worksheets("Variables (do not edit)").Range("H5").Value
    iRowOffset = plotWorkbook.Worksheets("Variables (do not edit)").Range("H6").Value

    Dim xPos As Long
    Dim yPos As Long
   
    xPos = iColOffset + 1
    yPos = iRowOffset
    
    Set dHeadingList = New Dictionary
    
    Do
        If plotWorkbook.Worksheets(plotWhichSheet).Cells(yPos, xPos).Value <> "" Then
            If Not dHeadingList.Exists(plotWorkbook.Worksheets(plotWhichSheet).Cells(yPos, xPos).Value) Then
                Call dHeadingList.Add(plotWorkbook.Worksheets(plotWhichSheet).Cells(yPos, xPos).Value, 0)
            End If
            yPos = yPos + zOffsetSize
        Else
            Exit Do
        End If
    Loop
    
    TransferToSigmaplotFrm.Show
    If doImport Then
        Call TransferToSigmaplot("")
    End If

End Sub

