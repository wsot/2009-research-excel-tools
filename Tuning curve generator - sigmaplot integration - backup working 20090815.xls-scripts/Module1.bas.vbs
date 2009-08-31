Attribute VB_Name = "Module1"
Option Explicit
Global doImport
Global theServer, theTank, theBlock

Sub buildTuningCurves()
    ImportFrom.Show
    
    If doImport Then
        Call processImport(False)
    End If
End Sub

Sub buildTuningCurvesIntoSigmaplot()
    ImportFrom.Show
    
    If doImport Then
        Call processImport(True)
    End If
End Sub

Sub processImport(importIntoSigmaplot As Boolean)
    Dim lBinWidth As Double
    lBinWidth = Worksheets("Settings").Range("B1").Value
    
    Dim lMaxHistHeight As Double
    lMaxHistHeight = 0
    
    Dim theWorksheets As Variant
    'Dim chanHistTmp(32) As Long 'used as a temporary store to build a histogram across multiple 'Swep's before outputting the data
    Dim arrHistTmp(31) As Long
    
    Const iRowOffset = 1
    Const iColOffset = 0

    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long

    theWorksheets = buildWorksheetArray()
    
    Dim objttx
    Set objttx = CreateObject("TTank.X")
    
    If objttx.ConnectServer(theServer, "Me") <> CLng(1) Then
        MsgBox ("Connecting to server " & theServer & " failed.")
        Exit Sub
    End If
    
    If objttx.OpenTank(theTank, "R") <> CLng(1) Then
        MsgBox ("Connecting to tank " & theTank & " on server " & theServer & " failed .")
        Call objttx.ReleaseServer
        Exit Sub
    End If
    
    If objttx.SelectBlock(theBlock) <> CLng(1) Then
        MsgBox ("Connecting to block " & theBlock & " in tank " & theTank & " on server " & theServer & " failed.")
        Call objttx.CloseTank
        Call objttx.ReleaseServer
        Exit Sub
    End If
    
    Call objttx.CreateEpocIndexing
    
    Dim freqList As Dictionary
    Dim ampList As Dictionary
    
    Set freqList = New Dictionary
    Set ampList = New Dictionary

    Dim dblStartTime As Double
    Dim dblEndTime As Double
    Dim varReturn As Variant
    Dim varAmp As Variant
    
    
    Do
        i = objttx.ReadEventsV(500, "Frq1", 0, 0, dblStartTime, 0#, "ALL")
        If i = 0 Then
            Exit Do
        End If
        
        varReturn = objttx.ParseEvInfoV(0, i, 0)
        For j = 0 To (i - 1)
            If Not freqList.Exists(varReturn(6, j)) Then
                Call freqList.Add(varReturn(6, j), "")
            End If
            dblStartTime = varReturn(5, j) + (1 / 100000)
            varAmp = objttx.QryEpocAtV("Lev1", varReturn(5, j), 0)
            If Not ampList.Exists(varAmp) Then
                Call ampList.Add(varAmp, "")
            End If
            
        Next
        
        If i < 500 Then
            Exit Do
        End If
    Loop
    
    i = 0
    j = 0
    
'    Dim freqAmpArray()
'    Dim freqAmpArray(freqList.Count - 1, ampList.Count - 1, 32)
    Dim iFreqIndex As Integer
    Dim iAmpIndex As Integer

    Dim vFreqKeys As Variant
    Dim vAmpKeys As Variant
                
        
    vFreqKeys = freqList.Keys
    vAmpKeys = ampList.Keys
        
    Dim varChanData As Variant
    Dim dblSwepStartTime As Double

    For iFreqIndex = 0 To UBound(vFreqKeys)
        For iAmpIndex = 0 To UBound(vAmpKeys)
            Call objttx.SetFilterWithDescEx("Frq1 = " & CStr(vFreqKeys(iFreqIndex)) & " and Lev1 = " & CStr(vAmpKeys(iAmpIndex)))
            varReturn = objttx.GetEpocsExV("Swep", 0)
            For i = 0 To UBound(varReturn, 2)
                dblStartTime = varReturn(2, i)
                dblEndTime = dblStartTime + lBinWidth
                'dblEndTime = varReturn(3, i)
                dblSwepStartTime = dblStartTime
                For j = 1 To 32
                    Do
                        k = objttx.ReadEventsV(500, "CSPK", j, 0, dblStartTime, dblEndTime, "JUSTTIMES")
                        If k = 0 Then
                            Exit Do
                        End If
                        
                        arrHistTmp(j - 1) = CLng(arrHistTmp(j - 1)) + CLng(k)
                        
'                        varChanData = objttx.ParseEvInfoV(0, k, 6)
'                        For l = 0 To (k - 1)
'                            Worksheets.Item("Settings").Cells(iAmpIndex + 3, 1).Value = j
'                            Worksheets.Item("Settings").Cells(iAmpIndex + 3, 2).Value = varChanData(0)
'                        Next
                        
                        If k < 500 Then
                            Exit Do
                        Else
                            varChanData = objttx.ParseEvInfoV(k - 1, 1, 6)
                            dblStartTime = varChanData(0) + (1 / 100000)
                        End If
                    Loop
                    dblStartTime = dblSwepStartTime
                Next
            Next
            For j = 1 To 32
                theWorksheets(j - 1).Cells((UBound(vAmpKeys) + iRowOffset + 2) - iAmpIndex, iFreqIndex + iColOffset + 2).Value = arrHistTmp(j - 1)
                If arrHistTmp(j - 1) > lMaxHistHeight Then
                    lMaxHistHeight = arrHistTmp(j - 1)
                End If
                
                arrHistTmp(j - 1) = 0
            Next
        Next
    Next

    Call writeAxes(theWorksheets, vFreqKeys, vAmpKeys, iColOffset, iRowOffset)

    Call objttx.CloseTank
    Call objttx.ReleaseServer
    
    If importIntoSigmaplot Then
        Call transferToSigmaplot(theWorksheets, vFreqKeys, vAmpKeys, iColOffset, iRowOffset, lMaxHistHeight)
    End If
    
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

Sub writeAxes(theWorksheets As Variant, colLabels As Variant, rowLabels As Variant, iColOffset, iRowOffset)
    Dim i As Long
    Dim j As Long
        
    For i = 0 To UBound(theWorksheets)
        For j = 0 To UBound(rowLabels)
            theWorksheets(i).Cells((UBound(rowLabels) + iRowOffset + 2) - j, iColOffset + 1).Value = rowLabels(j)
        Next
        For j = 0 To UBound(colLabels)
            theWorksheets(i).Cells(2, j + 2).Value = colLabels(j)
        Next
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

Sub transferToSigmaplot(theWorksheets As Variant, colLabels As Variant, rowLabels As Variant, iColOffset, iRowOffset, lMaxHistHeight)

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

    Dim spNB As Object
    Dim spWS As Object
    Dim spDT As Object
    Dim spGRPH As Object
    
    
    For i = 0 To UBound(theWorksheets)
        Set spNB = SPApp.Notebooks.Item(SPApp.Notebooks.Count - 1)
        Set spWS = spNB.NotebookItems.Item(spNB.NotebookItems.Count - 1)
        spWS.Name = theWorksheets(i).Name
        Set spDT = spWS.DataTable
        
        For j = 0 To UBound(rowLabels)
            spDT.Cell(1, j) = rowLabels(j)
        Next
        
        For j = 0 To UBound(colLabels)
            spDT.Cell(0, j) = colLabels(j)
        Next
        
        For j = 0 To UBound(colLabels)
            For k = 0 To UBound(rowLabels)
                spDT.Cell(3 + k, j) = theWorksheets(i).Cells((UBound(rowLabels) + iRowOffset + 2) - k, j + iColOffset + 2).Value
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
        ColumnsPerPlot(0, 3) = 3 + UBound(rowLabels)
        ColumnsPerPlot(1, 3) = 0
        ColumnsPerPlot(2, 3) = 31999999
        
        Dim PlotColumnCountArray()
        ReDim PlotColumnCountArray(0)
        
        PlotColumnCountArray(0) = 4
        Call SPApp.ActiveDocument.CurrentPageItem.CreateWizardGraph("Contour Plot", "Filled Contour Plot", "XY Many Z", ColumnsPerPlot, PlotColumnCountArray, "Worksheet Columns", "Standard Deviation", "Degrees", 0#, 360#, , "Standard Deviation", True)
        Call SPApp.ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).SelectObject
    
        SPApp.ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Name = "Site y"
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

        If i < UBound(theWorksheets) Then
            Call spNB.NotebookItems.Add(1)
        End If
    Next
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
