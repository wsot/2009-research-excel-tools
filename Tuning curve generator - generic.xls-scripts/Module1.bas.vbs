Attribute VB_Name = "Module1"
Option Explicit
Global doImport
Global theServer, theTank, theBlock
Global xAxisEp, yAxisEp, arrOtherEp
Global lBinWidth As Double
Global lIgnoreFirstMsec As Double
Global iRowOffset As Integer
Global iColOffset As Integer
Global bReverseX, bReverseY As Boolean

Global dHeadingList As Dictionary
Global dHeadingsSelected As Dictionary
Global bXAxisLog As Boolean

Dim vXAxisKeys As Variant
Dim vYAxisKeys As Variant


Sub buildTuningCurves()
    ImportFrom.Show
    
    If doImport Then
        Call processImport(False)
    End If
End Sub

Sub buildTuningCurvesIntoSigmaplot()
'    ImportFrom.Show
    
'    If doImport Then
'        Call processImport(True)
'    End If
Call TransferToSigmaplot
End Sub

Sub processImport(importIntoSigmaplot As Boolean)

    'load the bin width for histogram generation
    lBinWidth = Worksheets("Settings").Range("B1").Value
    
    'load the # of msec to ignore at the start (for filtering stimulation artifact
    lIgnoreFirstMsec = Worksheets("Settings").Range("B2").Value
    
    'used to store the maximum histogram peak for normalisation
    Dim lMaxHistHeight As Double
    lMaxHistHeight = 0
    
    Dim theWorksheets As Variant 'stores the created worksheets to write to
    Dim arrHistTmp() As Long 'used to store the histogram data for each channel as it is generated
    ReDim arrHistTmp(31)
    
    Dim yCount As Long
    Dim xCount As Long
    Dim zOffsetSize As Long
    
    'offsets to leave space at the top and left of the chart
    iRowOffset = 1
    iColOffset = 0

'    theWorksheets = buildWorksheetArray() 'build the worksheets for writing data
    
    'connect to the tank
    Dim objTTX
    Set objTTX = CreateObject("TTank.X")
    
    If objTTX.ConnectServer(theServer, "Me") <> CLng(1) Then
        MsgBox ("Connecting to server " & theServer & " failed.")
        Exit Sub
    End If
    
    If objTTX.OpenTank(theTank, "R") <> CLng(1) Then
        MsgBox ("Connecting to tank " & theTank & " on server " & theServer & " failed .")
        Call objTTX.ReleaseServer
        Exit Sub
    End If
    
    If objTTX.SelectBlock(theBlock) <> CLng(1) Then
        MsgBox ("Connecting to block " & theBlock & " in tank " & theTank & " on server " & theServer & " failed.")
        Call objTTX.CloseTank
        Call objTTX.ReleaseServer
        Exit Sub
    End If
    
    'index epochs - required to use filters
    Call objTTX.CreateEpocIndexing
    
    Dim dblStartTime As Double
    Dim dblEndTime As Double
    
    Dim varReturn As Variant
    
'    Dim vXAxisKeys As Variant
'    Dim vYAxisKeys As Variant
    
    vXAxisKeys = buildEpocList(objTTX, xAxisEp, bReverseX)
    vYAxisKeys = buildEpocList(objTTX, yAxisEp, bReverseY)
        
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    
    Dim arrOtherEpocKeys() As Variant
    If UBound(arrOtherEp) <> -1 Then
        ReDim arrOtherEpocKeys(UBound(arrOtherEp))
        
        For i = 0 To UBound(arrOtherEp)
            arrOtherEpocKeys(i) = buildEpocList(objTTX, arrOtherEp(i), False)
        Next
    End If
    
    i = 0
    j = 0
    
    Dim iXAxisIndex As Integer
    Dim iYAxisIndex As Integer
    Dim arrOtherEpocIndex() As Integer
    If UBound(arrOtherEp) <> -1 Then
        ReDim arrOtherEpocIndex(UBound(arrOtherEp))
    End If
        
    Dim varChanData As Variant
    Dim dblSwepStartTime As Double
    
    Dim xAxisSearchString As String
    Dim yAxisSearchString As String
    Dim otherAxisSearchString() As String
    Dim strOtherAxisSearchString As String
    If UBound(arrOtherEp) <> -1 Then
        ReDim otherAxisSearchString(UBound(arrOtherEp))
    End If

    Dim iChanNum As Integer
    iChanNum = 0

    If UBound(arrOtherEp) <> -1 Then
        For i = 0 To UBound(vXAxisKeys)
            If xAxisEp = "Channel" Then
                iChanNum = vXAxisKeys(i)
                xAxisSearchString = ""
            Else
                xAxisSearchString = xAxisEp & " = " & CStr(vXAxisKeys(i)) & " and "
            End If
            For j = 0 To UBound(vYAxisKeys)
                If yAxisEp = "Channel" Then
                    iChanNum = vYAxisKeys(j)
                    yAxisSearchString = ""
                Else
                    yAxisSearchString = yAxisEp & " = " & CStr(vYAxisKeys(j)) & " and "
                End If
                Call processSearch(objTTX, arrOtherEp, arrOtherEpocKeys, 0, xAxisSearchString & yAxisSearchString, i + 1, j + 1, UBound(vYAxisKeys) + 3, iChanNum, "", xCount, yCount, zOffsetSize, lMaxHistHeight)
            Next
        Next
    End If

'    Call writeAxes(theWorksheets, vXAxisKeys, vYAxisKeys, iColOffset, iRowOffset)

    Call objTTX.CloseTank
    Call objTTX.ReleaseServer
    
    Worksheets("Variables (do not edit)").Range("H1").Value = xCount
    Worksheets("Variables (do not edit)").Range("H2").Value = yCount
    Worksheets("Variables (do not edit)").Range("H3").Value = zOffsetSize
    Worksheets("Variables (do not edit)").Range("H4").Value = lMaxHistHeight
    Worksheets("Variables (do not edit)").Range("H5").Value = iColOffset
    Worksheets("Variables (do not edit)").Range("H6").Value = iRowOffset
    
    'If importIntoSigmaplot Then
        'Call transferToSigmaplot(xCount, yCount, zOffsetSize, iColOffset, iRowOffset, lMaxHistHeight)
    'End If
    
End Sub

Function buildWorksheetArray() As Variant
    Dim theWorksheets(31)
   Dim strWsname As String
   Dim intWSNum As Long

    Dim i As Integer

    For i = 1 To Worksheets.Count
        strWsname = Worksheets.Item(i).Name
        If Left(strWsname, 4) = "Site" Then
            If IsNumeric(Right(strWsname, Len(strWsname) - 4)) Then
                intWSNum = CInt(Right(strWsname, Len(strWsname) - 4))
                If intWSNum < 33 And intWSNum > 0 Then
                    Set theWorksheets(intWSNum - 1) = Worksheets.Item(i)
                End If
            End If
        End If
    Next

    For i = 0 To 31
        If IsEmpty(theWorksheets(i)) Then
            If i > 0 Then
                Set theWorksheets(i) = Worksheets.Add(, theWorksheets(i - 1), 1, xlWorksheet)
            Else
                Set theWorksheets(i) = Worksheets.Add(, Worksheets.Item(Worksheets.Count), 1, xlWorksheet)
            End If
            theWorksheets(i).Name = "Site" & CStr(i + 1)
        End If
    Next
    buildWorksheetArray = theWorksheets
End Function

Sub writeAxes(colLabels As Variant, rowLabels As Variant, iColOffset, iRowOffset, zOffset)
    Dim j As Long
        
    For j = 0 To UBound(rowLabels)
        Worksheets("Output").Cells(iRowOffset + j + 2 + zOffset, iColOffset + 1).Value = rowLabels(j)
    Next
    For j = 0 To UBound(colLabels)
        Worksheets("Output").Cells(iRowOffset + zOffset + 1, j + 2).Value = colLabels(j)
    Next

End Sub

Sub deleteWorksheets()
   Dim strWsname As String
   Dim intWSNum As Long

    Dim i As Integer
    
    i = Worksheets.Count
    
    Do
        If i = 0 Then
            Exit Do
        End If
        strWsname = Worksheets.Item(i).Name
        If Left(strWsname, 4) = "Site" Then
            If IsNumeric(Right(strWsname, Len(strWsname) - 4)) Then
                intWSNum = CInt(Right(strWsname, Len(strWsname) - 4))
                If intWSNum < 33 And intWSNum > 0 Then
                    Worksheets.Item(i).Delete
                End If
            End If
        End If
        i = i - 1
    Loop
End Sub

Sub TransferToSigmaplot()

    Dim xCount As Long
    Dim yCount As Long
    Dim zOffsetSize As Long
    Dim lMaxHistHeight As Long
    Dim iColOffset As Integer
    Dim iRowOffset As Integer

    xCount = Worksheets("Variables (do not edit)").Range("H1").Value
    yCount = Worksheets("Variables (do not edit)").Range("H2").Value
    zOffsetSize = Worksheets("Variables (do not edit)").Range("H3").Value
    lMaxHistHeight = Worksheets("Variables (do not edit)").Range("H4").Value
    iColOffset = Worksheets("Variables (do not edit)").Range("H5").Value
    iRowOffset = Worksheets("Variables (do not edit)").Range("H6").Value
    
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
        If Worksheets("Output").Cells(yPos, xPos).Value <> "" Then
            If dHeadingsSelected.Exists(Worksheets("Output").Cells(yPos, xPos).Value) Then
                strTitle = Worksheets("Output").Cells(yPos, xPos).Value
                Set spNB = SPApp.Notebooks.Item(SPApp.Notebooks.Count - 1)
                Set spWS = spNB.NotebookItems.Item(spNB.NotebookItems.Count - 1)
                spWS.Name = Worksheets("Output").Cells(yPos, xPos).Value
                Set spDT = spWS.DataTable
                
                yPos = yPos
                            
                For j = 0 To (xCount - 1)
                    spDT.Cell(0, j) = Worksheets("Output").Cells(yPos + 1, xPos + j + 1).Value
                Next
                
                For j = 0 To (yCount - 1)
                    spDT.Cell(1, j) = Worksheets("Output").Cells(yPos + j + 2, xPos).Value
                Next
                
                
                For j = 0 To (xCount - 1)
                    For k = 0 To (yCount - 1)
                        spDT.Cell(3 + k, j) = Worksheets("Output").Cells(yPos + k + 2, xPos + j + 1).Value
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
End Sub

Sub testSigmaPlot()
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
    SPApp.ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Name = "Site y"
    SPApp.ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Axes(0).Name = "Attenuation"
    SPApp.ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Axes(1).Name = "Frequency"
    
    Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, 3)
    Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_OPTIONS, 10)
    Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_OPTIONS, 51380225)
    Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_OPTIONS, 12583938)
    Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTRSTRING, SAA_FROMVAL, "0")
    Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTRSTRING, SAA_TOVAL, "150")
    
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

End Sub

Function buildEpocList(objTTX, AxisEp, bReverseOrder)
    'build list of epocs for the given axis epoc name
    
    Dim AxisList As Dictionary
    Set AxisList = New Dictionary
    
    Dim dblStartTime As Double
    Dim varReturn As Variant
    
    Dim i As Integer
    Dim j As Integer
    
    If AxisEp = "Channel" Then
        For i = 1 To 32
            Call AxisList.Add(i, 0)
        Next
    Else
        Do
            i = objTTX.ReadEventsV(500, AxisEp, 0, 0, dblStartTime, 0#, "ALL")
            If i = 0 Then
                Exit Do
            End If
            
            varReturn = objTTX.ParseEvInfoV(0, i, 0)
            For j = 0 To (i - 1)
                If Not AxisList.Exists(varReturn(6, j)) Then
                    Call AxisList.Add(varReturn(6, j), "")
                End If
                dblStartTime = varReturn(5, j) + (1 / 100000)
            Next
            
            If i < 500 Then
                Exit Do
            End If
        Loop
    End If
    
    
    
    If bReverseOrder Then
        Dim returnArr()
        Dim tempArr As Variant
        tempArr = AxisList.Keys
        ReDim returnArr(UBound(tempArr))

        For i = 0 To UBound(tempArr)
            returnArr(i) = tempArr(UBound(tempArr) - i)
        Next
        buildEpocList = returnArr
    Else
        buildEpocList = AxisList.Keys
    End If

End Function


Function processSearch(ByRef objTTX, ByRef arrOtherEp, ByRef arrOtherEpocKeys, iOtherEpocNum, strSearchString As String, xOffset, yOffset, zOffset, iChanNum, strTitle, ByRef xCount, ByRef yCount, ByRef zOffsetSize, ByRef lMaxHistHeight)
    Dim i As Integer
    Dim j As Integer
    Dim strAddedSearchString As String
    Dim strFilter As String
    Dim strAddedTitle As String
    
    For i = 0 To UBound(arrOtherEpocKeys(iOtherEpocNum))
        If arrOtherEp(iOtherEpocNum) <> "Channel" Then
            'add to search string
            strAddedSearchString = strSearchString & arrOtherEp(iOtherEpocNum) & " = " & CStr(arrOtherEpocKeys(iOtherEpocNum)(i)) & " and "
            strAddedTitle = strTitle & arrOtherEp(iOtherEpocNum) & " = " & CStr(arrOtherEpocKeys(iOtherEpocNum)(i)) & ", "
        Else
            strAddedSearchString = strSearchString
            strAddedTitle = strTitle & "Channel = " & CStr(arrOtherEpocKeys(iOtherEpocNum)(i)) & ", "
            iChanNum = arrOtherEpocKeys(iOtherEpocNum)(i)
        End If
        If iOtherEpocNum < UBound(arrOtherEp) Then
            'there are still more epocs to add to the search
            Call processSearch(objTTX, arrOtherEp, arrOtherEpocKeys, iOtherEpocNum + 1, strAddedSearchString, xOffset, yOffset, (zOffset * UBound(arrOtherEpocKeys(iOtherEpocNum))) + i, iChanNum, strAddedTitle, xCount, yCount, zOffsetSize, lMaxHistHeight)
        Else
            'we have reached the end of the list of epocs - can actually do a search now
            If Right(strAddedSearchString, 5) = " and " Then 'this should always be the case - should be a trailing 'and' to remove
                strFilter = Left(strAddedSearchString, Len(strAddedSearchString) - 5)
            Else
                strFilter = strAddedSearchString
            End If
            Call objTTX.SetFilterWithDescEx(strFilter)
            
            If xOffset = 1 And yOffset = 1 Then
                Worksheets("Output").Cells(iRowOffset + (i * zOffset), iColOffset + 1).Value = Left(strAddedTitle, Len(strAddedTitle) - 2)
                Call writeAxes(vXAxisKeys, vYAxisKeys, iColOffset, iRowOffset, (i * zOffset))
            End If

            Call writeResults(objTTX, xOffset, yOffset, i * zOffset, iChanNum, lMaxHistHeight)
            If xOffset > xCount Then
                xCount = xOffset
            End If
            If yOffset > yCount Then
                yCount = yOffset
            End If
            zOffsetSize = zOffset
        End If
    Next

End Function

Sub writeResults(ByRef objTTX, xOffset, yOffset, zOffset, iChanNum, ByRef lMaxHistHeight)
    Dim varReturn As Variant
    Dim varChanData As Variant
    
    Dim dblStartTime As Double
    Dim dblEndTime As Double
    Dim dblSwepStartTime As Double
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Dim histTmp As Long

    varReturn = objTTX.GetEpocsExV("Swep", 0)
    If IsArray(varReturn) Then
        For i = 0 To UBound(varReturn, 2)
            dblStartTime = varReturn(2, i) + lIgnoreFirstMsec
            dblEndTime = dblStartTime + lBinWidth + lIgnoreFirstMsec
            dblSwepStartTime = dblStartTime
            Do
                k = objTTX.ReadEventsV(500, "CSPK", iChanNum, 0, dblStartTime, dblEndTime, "JUSTTIMES")
                If k = 0 Then
                    Exit Do
                End If
    
                histTmp = CLng(histTmp) + CLng(k)
                If k < 500 Then
                    Exit Do
                Else
                    varChanData = objTTX.ParseEvInfoV(k - 1, 1, 6)
                    dblStartTime = varChanData(0) + (1 / 100000)
                End If
            Loop
            dblStartTime = dblSwepStartTime
        Next
        
        If xAxisEp = "Channel" Then
            Worksheets("Output").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Value = histTmp
        ElseIf yAxisEp = "Channel" Then
            Worksheets("Output").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Value = histTmp
        Else
            Worksheets("Output").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Value = histTmp
        End If
        If histTmp > lMaxHistHeight Then
            lMaxHistHeight = histTmp
        End If
        histTmp = 0
    End If
    
End Sub

Sub transferToSigmaplotButton()
    Dim zOffsetSize As Long
    Dim iColOffset As Integer
    Dim iRowOffset As Integer

    zOffsetSize = Worksheets("Variables (do not edit)").Range("H3").Value
    iColOffset = Worksheets("Variables (do not edit)").Range("H5").Value
    iRowOffset = Worksheets("Variables (do not edit)").Range("H6").Value

    Dim xPos As Long
    Dim yPos As Long
   
    xPos = iColOffset + 1
    yPos = iRowOffset
    
    Set dHeadingList = New Dictionary
    
    Do
        If Worksheets("Output").Cells(yPos, xPos).Value <> "" Then
            If Not dHeadingList.Exists(Worksheets("Output").Cells(yPos, xPos).Value) Then
                Call dHeadingList.Add(Worksheets("Output").Cells(yPos, xPos).Value, 0)
            End If
            yPos = yPos + zOffsetSize
        Else
            Exit Do
        End If
    Loop
    
    TransferToSigmaplotFrm.Show
    If doImport Then
        Call TransferToSigmaplot
    End If

End Sub
