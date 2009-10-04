Attribute VB_Name = "NeuralDataProcess"
Option Explicit
Global objTTX As TTankX
Global dblTotalWidthSecs As Double
Global dblBinWidthSecs As Double
Global dblStartOffsetSecs As Double
Global dictOnlyIncludeChannels As Dictionary

Dim theServer As String
Dim theTank As String
Dim theBlock As String
'Dim dAtten As Dictionary
'Dim dOldAtten As Dictionary
    
'Const initialEpocName = "TriS"
'Const stimEpocName = "SweS"

Dim blnBuildCharts As Boolean

Const ConnectSuccess = 0
Const ServerConnectFail = 1
Const TankConnectFail = 2
Const BlockConnectFail = 2

Sub ExtractNeuralDataWithCharts()
    blnBuildCharts = True
    Worksheets("Variables (do not edit)").Range("B6").Value = True
    Call ExtractNeuralData
End Sub
Sub ExtractNeuralDataWithoutCharts()
    blnBuildCharts = False
    Worksheets("Variables (do not edit)").Range("B6").Value = False
    Call ExtractNeuralData
End Sub


Sub ExtractNeuralData()
    
    Application.Calculation = xlCalculationManual
    
    Set objTTX = CreateObject("TTank.X") 'establish connection to TDT Tank engine
    
    If Not connectToTDT Then
        MsgBox "Connection to TDT could not be established."
        Set objTTX = Nothing
        Exit Sub
    End If
    
'Don't need any of the 'actual volume' calculations because we are not comparing between frequencies - only need to use raw values to check same number of stim with same property
'    Set dAtten = New Dictionary
'    Set dOldAtten = New Dictionary

'    Call loadAttenList(dAtten, "Attenuations")
'    Call loadAttenList(dOldAtten, "Attenuations (incorrect)")
    
    Call loadIncludeChannelList
    
    Call getParsingVariables
    
    Call Worksheets("Neural Data").UsedRange.Delete
'    Dim lChartDelete As Long
    While Worksheets("Neural Data").ChartObjects.Count > 0
    'For lChartDelete = 1 To Worksheets("Neural Data").ChartObjects.Count
        Call Worksheets("Neural Data").ChartObjects(1).Delete
    'Next
    Wend
    
    Call parseNeuralData
    
'    Set dAtten = Nothing
'    Set dOldAtten = Nothing
    
    Set objTTX = Nothing
    
    Application.Calculation = xlCalculationAutomatic
    
End Sub

Function connectToTDT()
    connectToTDT = False
    
    If theServer = "" Then
        theServer = Worksheets("Variables (do not edit)").Range("B1").Value
        theTank = Worksheets("Variables (do not edit)").Range("B2").Value
        theBlock = Worksheets("Variables (do not edit)").Range("B3").Value
    End If
    Select Case testSettings(theServer, theTank, theBlock)
        Case ConnectSuccess:
            connectToTDT = True
    End Select
End Function

Function testSettings(ActServer, ActTank, ActBlock)
    testSettings = ConnectSuccess
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

Function getParsingVariables()
    dblTotalWidthSecs = CDbl(Worksheets("Settings").Range("B20").Value)
    dblBinWidthSecs = CDbl(Worksheets("Settings").Range("B21").Value)
    dblStartOffsetSecs = CDbl(Worksheets("Settings").Range("B22").Value)
End Function

Function parseNeuralData()
    Dim neuroWS As Worksheet
    Set neuroWS = Worksheets("Neural Data")
    
    Dim trialDataWS As Worksheet
    Set trialDataWS = Worksheets("Output")
    
    Dim iTrialNum As Integer
    Dim iTrialNumTDT As Integer

    Dim lStim1Freq As Long
    Dim strStim1Filter As String

    Call objTTX.CreateEpocIndexing

    Dim returnVal As Variant

    iTrialNum = 1
    While trialDataWS.Cells(iTrialNum + 1, 1) <> "" 'iterate through all trials
        iTrialNumTDT = CInt(trialDataWS.Range("B" & (iTrialNum + 1)).Value)
        lStim1Freq = CLng(stripTrailingHz(trialDataWS.Range("F" & (iTrialNum + 1)).Value))
        strStim1Filter = "TriS = " & iTrialNumTDT & " AND AFrq = " & lStim1Freq
        
        Call objTTX.ResetFilters
        Call objTTX.SetFilterWithDescEx(strStim1Filter)
        
        returnVal = objTTX.GetEpocsExV("SweS", 0)
'        returnVal = objTTX.GetEpocsExV("Attn", 0)
'        If Not IsArray(returnVal) Then
'            returnVal = objTTX.GetEpocsExV("Ampl", 0)
'            isAtten = False
'        Else
'            isAtten = True
'        End If

        If Not IsArray(returnVal) Then
            MsgBox "Could not obtain Sweeps for search string: " & strStim1Filter
        Else
            Call readInTrialNeuralData(returnVal, neuroWS, trialDataWS, iTrialNum, lStim1Freq)
        End If
        
        Call objTTX.ResetFilters
        iTrialNum = iTrialNum + 1
    Wend
End Function

Function stripTrailingHz(strInput) As String
        'acoustic trial - drop the last 2 letters to remove the Hz
        If LCase(Right(strInput, 2)) = "hz" Then
            stripTrailingHz = Left(strInput, Len(strInput) - 2)
        Else
            stripTrailingHz = strInput
        End If
End Function

Function readInTrialNeuralData(stimEpocs As Variant, neuroWS As Worksheet, trialDataWS As Worksheet, iTrialNum As Integer, lStim1Freq As Long)
    Dim intChartGap As Integer
    
    If blnBuildCharts Then
        intChartGap = 21
    Else
        intChartGap = 0
    End If

    Dim returnVal As Variant
    Dim isAtten As Boolean 'true if the read value is an attenuation, false if it is an (incorrect) absolute amplitude (which needs to be corrected based on 'Attenuations (incorrect)' and 'Attenuations'
    Dim iStimNum As Long
    Dim k As Long
    
    Dim histoSums() As Variant
    Dim histoSquares() As Variant
    Dim histoN As Long
    Dim histoBinCount As Long
    
    histoN = 0
    histoBinCount = CInt(dblTotalWidthSecs / dblBinWidthSecs)
    Call setHistoArraySizes(histoSums, histoSquares, histoBinCount)
    
    'ReDim histoSums(histoBinCount)
    'ReDim histoSquares(histoBinCount)
    'Global dblTotalWidthSecs As Double
    'Global dblBinWidthSecs As Double
    'Global dblStartOffsetSecs As Double
    
    Dim stimAmp(2) As Integer 'this is used to store the individual frequencies for matching
    Dim stimAmpCounts(2) As Integer 'this is used to count the frequency of each amplitude of a given stimulation, to ensure even numbers between in-trial and pre-trial
    Dim stimAmpStep As Integer
    
    returnVal = objTTX.QryEpocAtV("Attn", stimEpocs(1, 0), 0) 'returnVal/stimEpocs offset 5 is time of event
    If IsEmpty(returnVal) Then
        isAtten = False
    Else
        isAtten = True
    End If
    
    For iStimNum = 0 To 8 'only want to look at the first 9 stims, because after than the shock will be on, which could screw up the neural data
        If isAtten Then
            returnVal = objTTX.QryEpocAtV("Attn", stimEpocs(1, iStimNum), 0) 'get the attenuation epoc at the stim time
        Else
            returnVal = objTTX.QryEpocAtV("Ampl", stimEpocs(1, iStimNum), 0) 'get the amplitude epoc at the stim time (which we don't actually need to correct because we are not looking at differences...)
        End If
        If IsEmpty(returnVal) Then
            MsgBox "SweS epoc occurred without paired Attn or Ampl epoc at time:" & stimEpocs(1, iStimNum)
        Else
            For stimAmpStep = 0 To 2
                If CInt(returnVal) = stimAmp(stimAmpStep) Then
                    stimAmpCounts(stimAmpStep) = stimAmpCounts(stimAmpStep) + 1
                    Exit For
                ElseIf stimAmp(stimAmpStep) = 0 Then
                    stimAmpCounts(stimAmpStep) = stimAmpCounts(stimAmpStep) + 1
                    stimAmp(stimAmpStep) = Int(returnVal)
                    Exit For
                End If
            Next
            
            histoN = histoN + 1
            Call buildHistogramForStimMethod1(stimEpocs(1, iStimNum), histoSums, histoSquares, False, Null, Null, histoBinCount)
        End If
    Next
    'once it has gotten to this point, it has the histogram data for all channels, and all bins in the histoSums and histoSquares arrays
    
    
    Dim iHistoBin As Integer
    
    'write out all the headings
    neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 1, 1).Value = "Trial " & iTrialNum
    neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 2, 1).Value = "Channel"
    neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 1, 3).Value = "Freq:"
    neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 2, 3).Value = lStim1Freq
    neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 1, 5).Value = "Mean:"
    neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 1, 7 + histoBinCount).Value = "StdDev:"
    For iHistoBin = 0 To histoBinCount
            'totals
            neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 2, 5 + iHistoBin).Value = _
                CStr(iHistoBin * dblBinWidthSecs) ' & "-" & CStr((iHistoBin + 1) * dblBinWidthSecs)
            'mean
            neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 2, 7 + histoBinCount + iHistoBin).Value = _
                CStr(iHistoBin * dblBinWidthSecs) ' & "-" & CStr((iHistoBin + 1) * dblBinWidthSecs)
            'stddev
            neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 2, 9 + histoBinCount * 2 + iHistoBin).Value = _
                CStr(iHistoBin * dblBinWidthSecs) ' & "-" & CStr((iHistoBin + 1) * dblBinWidthSecs)
    Next
    
    Dim myChart As ChartObject
    Dim chartOffset As Long
    Dim chartHeight As Long
    
    If blnBuildCharts Then
        chartOffset = neuroWS.Range(neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + dictOnlyIncludeChannels.Count * 2 + 4, 1), neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count + 5 + intChartGap) + dictOnlyIncludeChannels.Count * 2 + 3 + 21, 1)).Top
        chartHeight = neuroWS.Range(neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + dictOnlyIncludeChannels.Count * 2 + 4, 1), neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count + 5 + intChartGap) + dictOnlyIncludeChannels.Count * 2 + 3 + 21, 1)).Height
    End If

    Dim vChanKey As Variant
    'step through each channel
    For Each vChanKey In dictOnlyIncludeChannels.Keys
        neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + ((dictOnlyIncludeChannels(vChanKey) - 1) * 2) + 1 + 2, 1).Value = vChanKey
        For iHistoBin = 0 To histoBinCount
            'totals
            neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + ((dictOnlyIncludeChannels(vChanKey) - 1) * 2) + 1 + 2, 5 + iHistoBin).Value = histoSums(dictOnlyIncludeChannels(vChanKey) - 1)(iHistoBin)
            'mean
            neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + ((dictOnlyIncludeChannels(vChanKey) - 1) * 2) + 1 + 2, 7 + histoBinCount * 2 + iHistoBin).Value = histoSums(dictOnlyIncludeChannels(vChanKey) - 1)(iHistoBin) / histoN
            'stddev
            neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + ((dictOnlyIncludeChannels(vChanKey) - 1) * 2) + 1 + 2, 9 + histoBinCount + iHistoBin).Value = (histoSquares(dictOnlyIncludeChannels(vChanKey) - 1)(iHistoBin) - ((histoSums(dictOnlyIncludeChannels(vChanKey) - 1)(iHistoBin) ^ 2) / histoN) / (histoN - 1)) ^ 0.5
            'top of chart will be: (iTrialNum - 1) * (dictOnlyIncludeChannels.Count + 4) + dictOnlyIncludeChannels.Count + 3
           Next
            
            If blnBuildCharts Then
                Set myChart = neuroWS.ChartObjects.Add(((dictOnlyIncludeChannels(vChanKey) - 1) * 500) + 1, chartOffset, 500, chartHeight)
                myChart.Chart.ChartType = xlColumnClustered
                'myChart.Chart.SeriesCollection.NewSeries
                Call myChart.Chart.SetSourceData(neuroWS.Range(neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + ((dictOnlyIncludeChannels(vChanKey) - 1) * 2) + 1 + 2, 5), neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + ((dictOnlyIncludeChannels(vChanKey) - 1) * 2) + 1 + 2, 5 + histoBinCount)))
                myChart.Chart.ChartGroups(1).GapWidth = 0
                'myChart.Chart.Border.Weight = 0.25
                myChart.Chart.SeriesCollection(1).Name = "Channel " & vChanKey
                myChart.Chart.SeriesCollection(1).XValues = neuroWS.Range(neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 2, 5), neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 2, 5 + histoBinCount))
                myChart.Chart.SeriesCollection(1).Format.Line.Style = msoLineSingle
                myChart.Chart.SeriesCollection(1).Format.Line.Weight = 0.25
                myChart.Chart.SeriesCollection(1).Format.Line.Visible = msoTrue
                myChart.Chart.Legend.Delete
    
                myChart.Chart.ChartTitle.Characters.Font.Size = 12
            End If
    Next
    
End Function

Function buildHistogramForStimMethod1(ByVal dblStartTime As Double, ByRef histoSums As Variant, ByRef histoSquares As Variant, useFromEpocList As Boolean, ByRef stimAmp As Variant, ByRef stimAmpCount As Variant, lHistoBinCount As Long)
    Dim iChanNum As Integer
    Dim iEvtCount As Integer
    Dim iEvtNum As Integer
    Dim lBinNum As Long
    
    Dim dblEndTime As Double
    Dim dblFinalEndTime As Double
    Dim varData As Variant
    
    Dim nCount() As Long
    ReDim nCount(dictOnlyIncludeChannels.Count - 1)
    
    Dim dblInitialStartTime As Double
    dblInitialStartTime = dblStartTime
       
    dblStartTime = dblInitialStartTime + dblStartOffsetSecs
    dblEndTime = dblStartTime + dblBinWidthSecs
    For lBinNum = 0 To lHistoBinCount
            Do
                iEvtCount = objTTX.ReadEventsV(500, "CSPK", 0, 0, dblStartTime, dblEndTime, "ALL")
                If iEvtCount = 0 Then
                    Exit Do
                End If
            
                varData = objTTX.ParseEvInfoV(0, iEvtCount, 4)
            
                For iEvtNum = 0 To iEvtCount - 1
                    'count the number of events for each channel in the current bin
                    nCount(dictOnlyIncludeChannels(varData(iEvtNum)) - 1) = nCount(dictOnlyIncludeChannels(varData(iEvtNum)) - 1) + 1
                Next
    
                'if the full 500 was retrieved, there may be more to fetch, so try to fetch them
                If iEvtCount < 500 Then
                    Exit Do
                Else
                    'get the time of the last event, and search forward from that - there is a risk this could miss events where the time is identical, however. That said, never got more than 500 event yet
                    MsgBox "Obtained 500+ events!"
                    varData = objTTX.ParseEvInfoV(iEvtCount - 1, 1, 6)
                    dblStartTime = varData(0) + (1 / 100000)
                End If
            Loop
            
            'update the totals with the obtained number of events
            For iChanNum = 0 To UBound(nCount)
                histoSums(iChanNum)(lBinNum) = histoSums(iChanNum)(lBinNum) + nCount(iChanNum)
                histoSquares(iChanNum)(lBinNum) = histoSquares(iChanNum)(lBinNum) + (nCount(iChanNum) ^ 2)
            Next
            ReDim nCount(dictOnlyIncludeChannels.Count - 1) 'clear the storage array

            dblStartTime = dblEndTime
            dblEndTime = dblStartTime + dblBinWidthSecs
        Next

End Function

'load the list of channels to include from the spreadsheet - if none specified then all channels (up to the number provided in B23) included
Function loadIncludeChannelList()
    Dim iCounter As Integer
    Dim iChanCount As Integer
    iChanCount = Worksheets("Settings").Range("B23").Value
    
    Set dictOnlyIncludeChannels = New Dictionary
    
    If Worksheets("Settings").Range("B25") = "" Then
        For iCounter = 1 To iChanCount
            Call dictOnlyIncludeChannels.Add(iCounter, iCounter)
        Next
    Else
        Dim arrElements As Variant
        arrElements = Split(Worksheets("Settings").Range("B25"), ",", , vbTextCompare)
        For iCounter = 0 To UBound(arrElements)
            If Not dictOnlyIncludeChannels.Exists(arrElements(iCounter)) Then
                Call dictOnlyIncludeChannels.Add(arrElements(iCounter), iCounter)
            End If
        Next
    End If
    
End Function

'creates arrays the right size for the histogram data
Function setHistoArraySizes(ByRef histoSums As Variant, ByRef histoSquares As Variant, ByRef histoBinCount As Long)
    Dim i As Long
    
    Dim arrDoubles() As Double
        
    ReDim histoSums(dictOnlyIncludeChannels.Count - 1)
    ReDim histoSquares(dictOnlyIncludeChannels.Count - 1)
    
    'ReDim arrVariants(dictOnlyIncludeChannels.Count - 1)
    
    ReDim arrDoubles(histoBinCount)
    
    For i = 0 To dictOnlyIncludeChannels.Count - 1
        histoSums(i) = arrDoubles
        histoSquares(i) = arrDoubles
    Next
End Function

