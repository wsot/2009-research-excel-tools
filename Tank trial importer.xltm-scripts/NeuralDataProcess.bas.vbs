Attribute VB_Name = "NeuralDataProcess"
Option Explicit
Dim objTTX As TTankX
'Dim dblTotalWidthSecs As Double
'Dim dblBinWidthSecs As Double
'Dim dblStartOffsetSecs As Double
Dim dictOnlyIncludeChannels As Dictionary
Dim undrivenAction As Integer

Dim theServer As String
Dim theTank As String
Dim theBlock As String
'Dim dAtten As Dictionary
'Dim dOldAtten As Dictionary
    
'Const initialEpocName = "TriS"
'Const stimEpocName = "SweS"
Const snipEpocName = "CSPK"

Dim blnBuildCharts As Boolean
Dim unmatchedStimCell As Range
Dim undrivenChanCell As Range
Dim drivenChanOnsetDetectedCell As Range
Dim drivenChanDifferenceDetectedCell As Range

Const ConnectSuccess = 0
Const ServerConnectFail = 1
Const TankConnectFail = 2
Const BlockConnectFail = 3

Const UndrivenNoAction = 0
Const UndrivenExclude = 1
Const UndrivenMark = 2

Const DriveDetect_Undriven = 0
Const DriveDetect_OnsetDetected = 1
Const DriveDetect_ActDiffDetected = 2

Const DriveDetect_ActivityDifferenceThreshold = 2#


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
    Set unmatchedStimCell = Worksheets("Settings").Cells(28, 2)
    Set undrivenChanCell = Worksheets("Settings").Cells(29, 2)
    Set drivenChanOnsetDetectedCell = Worksheets("Settings").Cells(30, 2)
    Set drivenChanDifferenceDetectedCell = Worksheets("Settings").Cells(31, 2)
    
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
    
    Dim dblTotalWidthSecs As Double
    Dim dblBinWidthSecs As Double
    Dim dblStartOffsetSecs As Double
    
    Call getParsingVariables(dblTotalWidthSecs, dblBinWidthSecs, dblStartOffsetSecs)
    
    Call Worksheets("Neural Data").UsedRange.Clear
    Call Worksheets("Neural Data").UsedRange.ClearFormats
'    Dim lChartDelete As Long
    While Worksheets("Neural Data").ChartObjects.Count > 0
    'For lChartDelete = 1 To Worksheets("Neural Data").ChartObjects.Count
        Call Worksheets("Neural Data").ChartObjects(1).Delete
    'Next
    Wend
    
    If Not containsSnips Then
        Set objTTX = Nothing
        Worksheets("Variables (do not edit)").Range("B7").Value = False
        Exit Sub
    End If
    
    Call parseNeuralData(dblTotalWidthSecs, dblBinWidthSecs, dblStartOffsetSecs)
    Worksheets("Variables (do not edit)").Range("B7").Value = True
'    Set dAtten = Nothing
'    Set dOldAtten = Nothing
    
    Set objTTX = Nothing
    
    'Application.Calculation = xlCalculationAutomatic
    
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

Function getParsingVariables(ByRef dblTotalWidthSecs As Double, ByRef dblBinWidthSecs As Double, ByRef dblStartOffsetSecs As Double)
    dblTotalWidthSecs = CDbl(Worksheets("Settings").Range("B20").Value)
    dblBinWidthSecs = CDbl(Worksheets("Settings").Range("B21").Value)
    dblStartOffsetSecs = CDbl(Worksheets("Settings").Range("B22").Value)
    If CBool(Worksheets("Settings").Range("B34").Value) Then
        undrivenAction = UndrivenExclude
    Else
        undrivenAction = UndrivenNoAction
    End If
End Function

Function parseNeuralData(dblTotalWidthSecs As Double, dblBinWidthSecs As Double, dblStartOffsetSecs As Double)
    Dim iTrialNum As Integer

    Dim neuroWS As Worksheet
    Set neuroWS = Worksheets("Neural Data")
    
    Dim trialDataWS As Worksheet
    Set trialDataWS = Worksheets("Output")
    
    Call objTTX.CreateEpocIndexing

    iTrialNum = 1
    While trialDataWS.Cells(iTrialNum + 1, 1) <> "" 'iterate through all trials
'        iTrialNumTDT = CInt(trialDataWS.Range("B" & (iTrialNum + 1)).Value)
'        lStim1Freq = CLng(stripTrailingHz(trialDataWS.Range("F" & (iTrialNum + 1)).Value))
'        strStim1Filter = "TriS = " & iTrialNumTDT & " AND AFrq = " & lStim1Freq
        
'        Call objTTX.ResetFilters
'        Call objTTX.SetFilterWithDescEx(strStim1Filter)
        
'        returnVal = objTTX.GetEpocsExV("SweS", 0)

'        If Not IsArray(returnVal) Then
'            MsgBox "Could not obtain Sweeps for search string: " & strStim1Filter
'        Else
'            Call readInTrialNeuralData(returnVal, neuroWS, trialDataWS, iTrialNum, lStim1Freq)
'        End If
        
'        Call objTTX.ResetFilters
        'find first trial actual start time
        Call readTrialNeuralData(iTrialNum, neuroWS, trialDataWS, dblTotalWidthSecs, dblBinWidthSecs, dblStartOffsetSecs)
        
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

Function readTrialNeuralData(iTrialNum As Integer, neuroWS As Worksheet, trialDataWS As Worksheet, dblTotalWidthSecs As Double, dblBinWidthSecs As Double, dblStartOffsetSecs As Double)
    Dim iTrialNumTDT As Integer

    Dim lStim1Freq As Long
    Dim strStim1Filter As String
    
    'get the trial number with reference to TDT for the current trial
    iTrialNumTDT = CInt(trialDataWS.Range("B" & (iTrialNum + 1)).Value)
    'get the frequency of the 'continued' stimulus for this trial from the existing trial data
    lStim1Freq = CLng(stripTrailingHz(trialDataWS.Range("F" & (iTrialNum + 1)).Value))
    'build a filter to get this frequency in the given trial
    strStim1Filter = "TriS = " & iTrialNumTDT & " AND AFrq = " & lStim1Freq
        
    Call objTTX.ResetFilters
    Call objTTX.SetFilterWithDescEx(strStim1Filter)
        
    'find the sweep times within the 'alternating' period of the given for this frequency
    Dim stimEpocs As Variant
    stimEpocs = objTTX.GetEpocsExV("SweS", 0)

    'if none found, something is very wrong
    If Not IsArray(stimEpocs) Then
        MsgBox "Could not obtain Sweeps for search string: " & strStim1Filter
        Exit Function
    End If
    
    'intChartGap is used to give extra space in the output for charts
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
    
    'dim variables used to store output of the histogram generation
    Dim histoSums() As Variant
    Dim histoSquares() As Variant
    Dim histoN As Long
    Dim histoBinCount As Long
    Dim histoMaxTotal As Long
    Dim histoMaxMean As Double
    
    'create a linked list for links to charts, later used to update the scales
    Dim chartList As clsLinkedList
    Set chartList = New clsLinkedList
    
    histoN = 0
    'calculate the total number of bin (always has one extra at the end, but not having that seems to lead to array overrun problems...)
    histoBinCount = CInt(dblTotalWidthSecs / dblBinWidthSecs)
    'set the arrays to fit the data to go into them
    Call setHistoArraySizes(histoSums, histoSquares, histoBinCount)
    Call outputHeaders(neuroWS, intChartGap, histoBinCount, iTrialNum, lStim1Freq, dblTotalWidthSecs, dblBinWidthSecs, dblStartOffsetSecs) 'write out the headings for the current block of histograms
    
    'arrays used to store frequency/values of amplitudes
    Dim stimAmp(2) As Integer 'this is used to store the individual frequencies for matching
    Dim stimAmpCounts(2) As Integer 'this is used to count the frequency of each amplitude of a given stimulation, to ensure even numbers between in-trial and pre-trial
    Dim stimAmpStep As Integer
    
    returnVal = objTTX.QryEpocAtV("Attn", stimEpocs(1, 0), 0) 'returnVal/stimEpocs offset 5 is time of event
    If IsEmpty(returnVal) Then
        isAtten = False
    Else
        isAtten = True
    End If
    
    Dim dDrivenChanList As Dictionary
    Set dDrivenChanList = New Dictionary
    Call identifyDrivenChannels(stimEpocs, dDrivenChanList)
    
    'Dim lHistoBin As Long
    
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
            Call buildHistogramForStimMethod1(stimEpocs(1, iStimNum), histoSums, histoSquares, histoBinCount, dblTotalWidthSecs, dblBinWidthSecs, dblStartOffsetSecs)
        End If
    Next
    'once it has gotten to this point, it has the histogram data for all channels, and all bins in the histoSums and histoSquares arrays
    
    Call renderAmpList(stimAmpCounts, stimAmp, intChartGap, iTrialNum, neuroWS)
     
    Call outputResults(neuroWS, intChartGap, histoBinCount, iTrialNum, histoSums, histoSquares, histoN, 0, chartList, histoMaxTotal, histoMaxMean, dDrivenChanList)
    
    histoN = 0
    Call setHistoArraySizes(histoSums, histoSquares, histoBinCount) 'flush the histo data
    Call objTTX.ResetFilters
    Call objTTX.SetFilterWithDescEx("TriS = " & iTrialNumTDT)
        
    Dim vTrialsList As Variant
    Dim dblTrialStart As Double
    vTrialsList = objTTX.GetEpocsExV("TriS", 0)
    dblTrialStart = vTrialsList(1, 0)
    
    Dim iMatchesLeft As Integer
    iMatchesLeft = 9 'check we match all the stim
    Dim iPrevStimCount As Integer
    If isAtten Then
        iPrevStimCount = objTTX.ReadEventsV(10000, "Attn", 0, 0, dblTrialStart - 60, dblTrialStart, "ALL") 'look for previous 60s for the stimulus
    Else
        iPrevStimCount = objTTX.ReadEventsV(10000, "Ampl", 0, 0, dblTrialStart - 60, dblTrialStart, "ALL") 'look for previous 60s for the stimulus
    End If
    
    If iPrevStimCount = 0 Then
        MsgBox "Couldn't find any previous stim??"
        Exit Function
    End If
    
    returnVal = objTTX.ParseEvInfoV(0, iPrevStimCount, 0)
    
    For iStimNum = (iPrevStimCount - 1) To 0 Step -1
        For stimAmpStep = 0 To 2
            If CInt(returnVal(6, iStimNum)) = stimAmp(stimAmpStep) Then
                If stimAmpCounts(stimAmpStep) > 0 Then 'check if we want a stim of this amplitude
                    'yes - let's process it =)
                    stimAmpCounts(stimAmpStep) = stimAmpCounts(stimAmpStep) - 1
                    iMatchesLeft = iMatchesLeft - 1
                    histoN = histoN + 1
                    Call buildHistogramForStimMethod1(returnVal(5, iStimNum), histoSums, histoSquares, histoBinCount, dblTotalWidthSecs, dblBinWidthSecs, dblStartOffsetSecs)
                End If
                Exit For
            End If
        Next
        If iMatchesLeft = 0 Then
            Exit For
        End If
    Next

    For stimAmpStep = 0 To 2
        neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 9 + stimAmpStep * 4, 2).Value = stimAmpCounts(stimAmpStep)
'        If stimAmpCounts(stimAmpStep) > 0 Then 'check if we didn't find all instances of this stim
'            neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 9 + stimAmpStep * 4, 2).Interior.Color = unmatchedStimCell.Interior.Color
'            neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 9 + stimAmpStep * 4, 2).Font.Color = unmatchedStimCell.Font.Color
'        End If
    Next
    
    If iMatchesLeft = 0 Then
        neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 4, 1).Value = "Pre-trial span (s):"
        neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 4, 2).Value = Round(dblTrialStart - returnVal(5, iStimNum), 2)
    Else
        neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 4, 1).Value = "No match made"
        neuroWS.Range(neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 1, 1), neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + dictOnlyIncludeChannels.Count * 2 + 3, histoBinCount * 3 + 5)).Interior.Color = unmatchedStimCell.Interior.Color
        neuroWS.Range(neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 1, 1), neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + dictOnlyIncludeChannels.Count * 2 + 3, histoBinCount * 3 + 5)).Font.Color = unmatchedStimCell.Font.Color

'        neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 4, 1).Interior.Color = unmatchedStimCell.Interior.Color
'        neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 4, 1).Font.Color = unmatchedStimCell.Font.Color
    End If
    
    Call outputResults(neuroWS, intChartGap, histoBinCount, iTrialNum, histoSums, histoSquares, histoN, 1, chartList, histoMaxTotal, histoMaxMean, dDrivenChanList)
    
    If blnBuildCharts Then
        Call setChartScales(chartList, histoMaxTotal, histoMaxMean)
    End If
    
    Set dDrivenChanList = Nothing
    
End Function

Function buildHistogramForStimMethod1(ByVal dblStartTime As Double, ByRef histoSums As Variant, ByRef histoSquares As Variant, lHistoBinCount As Long, dblTotalWidthSecs As Double, dblBinWidthSecs As Double, dblStartOffsetSecs As Double)
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
                iEvtCount = objTTX.ReadEventsV(10000, "CSPK", 0, 0, dblStartTime, dblEndTime, "ALL")
                If iEvtCount = 0 Then
                    Exit Do
                End If
            
                varData = objTTX.ParseEvInfoV(0, iEvtCount, 4)
            
                For iEvtNum = 0 To iEvtCount - 1
                    'count the number of events for each channel in the current bin
                    nCount(dictOnlyIncludeChannels(varData(iEvtNum)) - 1) = nCount(dictOnlyIncludeChannels(varData(iEvtNum)) - 1) + 1
                Next
    
                'if the full 10000 was retrieved, there may be more to fetch, so try to fetch them
                If iEvtCount < 10000 Then
                    Exit Do
                Else
                    'get the time of the last event, and search forward from that - there is a risk this could miss events where the time is identical, however. That said, never got more than 10000 event yet
                    MsgBox "Obtained 10000+ events!"
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

Function outputHeaders(neuroWS As Worksheet, intChartGap As Integer, histoBinCount As Long, iTrialNum As Integer, lStim1Freq As Long, dblTotalWidthSecs As Double, dblBinWidthSecs As Double, dblStartOffsetSecs As Double)
    Dim lHistoBin As Long
    
    'write out all the headings
    neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 1, 1).Value = "Trial " & iTrialNum
    neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 2, 3).Value = "Driven?"
    neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 2, 4).Value = "Channel"
    neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 1, 3).Value = "Freq:"
    neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 1, 4).Value = lStim1Freq
    neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 1, 5).Value = "Total:"
    neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 1, 7 + histoBinCount).Value = "Mean:"
    neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 1, 9 + histoBinCount * 2).Value = "StdDev:"
    For lHistoBin = 0 To histoBinCount
            'totals
            neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 2, 5 + lHistoBin).Value = _
                CStr(lHistoBin * dblBinWidthSecs) ' & "-" & CStr((lHistoBin + 1) * dblBinWidthSecs)
            'mean
            neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 2, 7 + histoBinCount + lHistoBin).Value = _
                CStr(lHistoBin * dblBinWidthSecs) ' & "-" & CStr((lHistoBin + 1) * dblBinWidthSecs)
            'stddev
            neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 2, 9 + histoBinCount * 2 + lHistoBin).Value = _
                CStr(lHistoBin * dblBinWidthSecs) ' & "-" & CStr((lHistoBin + 1) * dblBinWidthSecs)
    Next
End Function

Function outputResults(neuroWS As Worksheet, intChartGap As Integer, histoBinCount As Long, iTrialNum As Integer, histoSums As Variant, histoSquares As Variant, histoN As Long, iOffset As Integer, ByRef chartList As clsLinkedList, ByRef histoMaxTotal As Long, ByRef histoMaxMean As Double, dDrivenChanList As Dictionary)
    Dim myChart As ChartObject
    Dim iChartNum As Integer
    Dim lChartTopPos As Long
    Dim lChartHeight As Long
    Dim lHistoBin As Long
    
    Dim iChartOffset As Integer
    Dim vBarColour As Variant
    Dim sTitleAdjustment As String
    iChartNum = 1
    
    iChartOffset = iOffset * intChartGap
    Select Case iOffset
        Case 0:
            sTitleAdjustment = " alternating"
            vBarColour = RGB(247, 150, 70)
        Case 1:
            sTitleAdjustment = " repeated"
            vBarColour = RGB(85, 142, 213)
    End Select
    
    If blnBuildCharts Then
        lChartTopPos = neuroWS.Range(neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + (dictOnlyIncludeChannels.Count * 2) + 4 + iChartOffset, 1), neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + (dictOnlyIncludeChannels.Count * 2) + 3 + 21 + iChartOffset, 1)).Top
        lChartHeight = neuroWS.Range(neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + (dictOnlyIncludeChannels.Count * 2) + 4 + iChartOffset, 1), neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + (dictOnlyIncludeChannels.Count * 2) + 3 + 21 + iChartOffset, 1)).Height
        'neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + (dictOnlyIncludeChannels.Count * 2) + 4, 1)
    End If

    Dim vChanKey As Variant
    'step through each channel
    For Each vChanKey In dictOnlyIncludeChannels.Keys
        If iOffset = 0 Then
            neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + ((dictOnlyIncludeChannels(vChanKey) - 1) * 2) + 1 + 2, 4).Value = vChanKey
        End If
        For lHistoBin = 0 To histoBinCount
            'totals
            neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + ((dictOnlyIncludeChannels(vChanKey) - 1) * 2) + 1 + 2 + iOffset, 5 + lHistoBin).Value = histoSums(dictOnlyIncludeChannels(vChanKey) - 1)(lHistoBin)
            If histoSums(dictOnlyIncludeChannels(vChanKey) - 1)(lHistoBin) > histoMaxTotal Then histoMaxTotal = histoSums(dictOnlyIncludeChannels(vChanKey) - 1)(lHistoBin)
            'mean
            neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + ((dictOnlyIncludeChannels(vChanKey) - 1) * 2) + 1 + 2 + iOffset, 7 + histoBinCount + lHistoBin).Value = histoSums(dictOnlyIncludeChannels(vChanKey) - 1)(lHistoBin) / histoN
            If (histoSums(dictOnlyIncludeChannels(vChanKey) - 1)(lHistoBin) / histoN) > histoMaxMean Then histoMaxMean = (histoSums(dictOnlyIncludeChannels(vChanKey) - 1)(lHistoBin) / histoN)
            'stddev
            neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + ((dictOnlyIncludeChannels(vChanKey) - 1) * 2) + 1 + 2 + iOffset, 9 + histoBinCount * 2 + lHistoBin).Value = (histoSquares(dictOnlyIncludeChannels(vChanKey) - 1)(lHistoBin) - ((histoSums(dictOnlyIncludeChannels(vChanKey) - 1)(lHistoBin) ^ 2) / histoN) / (histoN - 1)) ^ 0.5
            'top of chart will be: (iTrialNum - 1) * (dictOnlyIncludeChannels.Count + 4) + dictOnlyIncludeChannels.Count + 3
        Next
        
        neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + ((dictOnlyIncludeChannels(vChanKey) - 1) * 2) + 1 + 2 + iOffset, 3).Value = dDrivenChanList(vChanKey)
        
        Select Case dDrivenChanList(vChanKey)
        Case DriveDetect_Undriven:
            neuroWS.Range(neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + ((dictOnlyIncludeChannels(vChanKey) - 1) * 2) + 1 + 2 + iOffset, 5), neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + ((dictOnlyIncludeChannels(vChanKey) - 1) * 2) + 1 + 2 + iOffset, 9 + histoBinCount * 3)).Interior.Color = undrivenChanCell.Interior.Color
            neuroWS.Range(neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + ((dictOnlyIncludeChannels(vChanKey) - 1) * 2) + 1 + 2 + iOffset, 5), neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + ((dictOnlyIncludeChannels(vChanKey) - 1) * 2) + 1 + 2 + iOffset, 9 + histoBinCount * 3)).Font.Color = undrivenChanCell.Font.Color
        Case DriveDetect_OnsetDetected:
            neuroWS.Range(neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + ((dictOnlyIncludeChannels(vChanKey) - 1) * 2) + 1 + 2 + iOffset, 5), neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + ((dictOnlyIncludeChannels(vChanKey) - 1) * 2) + 1 + 2 + iOffset, 9 + histoBinCount * 3)).Interior.Color = drivenChanOnsetDetectedCell.Interior.Color
            neuroWS.Range(neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + ((dictOnlyIncludeChannels(vChanKey) - 1) * 2) + 1 + 2 + iOffset, 5), neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + ((dictOnlyIncludeChannels(vChanKey) - 1) * 2) + 1 + 2 + iOffset, 9 + histoBinCount * 3)).Font.Color = drivenChanOnsetDetectedCell.Font.Color
        Case DriveDetect_ActDiffDetected:
            neuroWS.Range(neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + ((dictOnlyIncludeChannels(vChanKey) - 1) * 2) + 1 + 2 + iOffset, 5), neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + ((dictOnlyIncludeChannels(vChanKey) - 1) * 2) + 1 + 2 + iOffset, 9 + histoBinCount * 3)).Interior.Color = drivenChanDifferenceDetectedCell.Interior.Color
            neuroWS.Range(neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + ((dictOnlyIncludeChannels(vChanKey) - 1) * 2) + 1 + 2 + iOffset, 5), neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + ((dictOnlyIncludeChannels(vChanKey) - 1) * 2) + 1 + 2 + iOffset, 9 + histoBinCount * 3)).Font.Color = drivenChanDifferenceDetectedCell.Font.Color
        End Select
        
        If (undrivenAction <> UndrivenExclude) Or (undrivenAction = UndrivenExclude And Not (dDrivenChanList(vChanKey) = DriveDetect_Undriven)) Then 'check if this channel should be excluded from chart generation
            If blnBuildCharts Then
                Set myChart = neuroWS.ChartObjects.Add(((iChartNum - 1) * 500) + 1, lChartTopPos, 500, lChartHeight)
                
                Call chartList.Append(myChart)
                
                myChart.Chart.ChartType = xlColumnClustered
                'myChart.Chart.SeriesCollection.NewSeries
                Call myChart.Chart.SetSourceData(neuroWS.Range(neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + ((dictOnlyIncludeChannels(vChanKey) - 1) * 2) + 1 + 2 + iOffset, 5), neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + ((dictOnlyIncludeChannels(vChanKey) - 1) * 2) + 1 + 2 + iOffset, 5 + histoBinCount)))
                myChart.Chart.ChartGroups(1).GapWidth = 0
                'myChart.Chart.Border.Weight = 0.25
                myChart.Chart.SeriesCollection(1).Name = "Chan " & vChanKey & " " & sTitleAdjustment
                myChart.Chart.SeriesCollection(1).XValues = neuroWS.Range(neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 2, 5), neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 2, 5 + histoBinCount))
                myChart.Chart.SeriesCollection(1).Format.Line.Style = msoLineSingle
                myChart.Chart.SeriesCollection(1).Format.Line.Weight = 0.25
                myChart.Chart.SeriesCollection(1).Format.Line.Visible = msoTrue
                'myChart.Chart.SeriesCollection(1).Format.Fill.Type = msoFillSolid
                myChart.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = vBarColour
                myChart.Chart.Legend.Delete
                myChart.Chart.ChartTitle.Characters.Font.Size = 12
                
'                If (undrivenAction = UndrivenMark) And (dDrivenChanList(vChanKey) = DriveDetect_Undriven) Then  'check if this channel should be highlighted as undriven
'                    myChart.Chart.ChartArea.Format.Fill.BackColor = undrivenChanCell.Interior.Color
                    'undrivenChanCell.Interior.Color
                    'drivenChanOnsetDetectedCell.Interior.Color
                    'drivenChanDifferenceDetectedCell.Interior.Color
'                Else
                    Select Case dDrivenChanList(vChanKey)
                    Case DriveDetect_Undriven:
                        myChart.Chart.ChartArea.Format.Fill.ForeColor.RGB = undrivenChanCell.Interior.Color
                    Case DriveDetect_OnsetDetected:
                        'myChart.Chart.ChartArea.Format.Fill.BackColor.RGB = RGB(200, 200, 250)
                        myChart.Chart.ChartArea.Format.Fill.ForeColor.RGB = drivenChanOnsetDetectedCell.Interior.Color
                    Case DriveDetect_ActDiffDetected:
                        myChart.Chart.ChartArea.Format.Fill.ForeColor.RGB = drivenChanDifferenceDetectedCell.Interior.Color
                    End Select
'                End If
                iChartNum = iChartNum + 1
            End If
        End If
    Next
End Function

Function renderAmpList(stimAmpCounts As Variant, stimAmp As Variant, intChartGap As Integer, iTrialNum As Integer, neuroWS As Worksheet)
    Dim iAmpOffset As Integer
    For iAmpOffset = 0 To 2
        neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 7 + iAmpOffset * 4, 1).Value = "Amp/Attn:"
        neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 7 + iAmpOffset * 4, 2).Value = stimAmp(iAmpOffset)
        neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 8 + iAmpOffset * 4, 1).Value = "Count:"
        neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 8 + iAmpOffset * 4, 2).Value = stimAmpCounts(iAmpOffset)
        neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 9 + iAmpOffset * 4, 1).Value = "Unmatched:"
        neuroWS.Cells((iTrialNum - 1) * (dictOnlyIncludeChannels.Count * 2 + 5 + intChartGap * 2) + 9 + iAmpOffset * 4, 2).Value = stimAmpCounts(iAmpOffset)
    Next
End Function


Function containsSnips()
    containsSnips = False
    Dim lCounter As Long
    Dim vEvtCodes As Variant
    Dim sEvtName As String
    vEvtCodes = objTTX.GetEventCodes(0)
    
    If IsArray(vEvtCodes) Then
        For lCounter = LBound(vEvtCodes) To UBound(vEvtCodes)
            sEvtName = objTTX.CodeToString(vEvtCodes(lCounter))
            If LCase(sEvtName) = LCase(snipEpocName) Then
                containsSnips = True
                Exit For
            End If
        Next
    End If
    
End Function

Function setChartScales(chartList As clsLinkedList, histoMaxTotal As Long, histoMaxMean As Double)
    Dim iChartCount As Integer
    Dim iOffset As Integer
    Dim theChart As ChartObject
    
    iChartCount = chartList.Count
    For iOffset = 1 To iChartCount
        Set theChart = chartList.Item(iOffset)
        theChart.Chart.Axes(xlValue).MinimumScale = 0
        theChart.Chart.Axes(xlValue).MaximumScale = histoMaxTotal
    Next
End Function

Function identifyDrivenChannels(stimEpocs As Variant, dDrivenChanList As Dictionary)
    'ByVal dblStartTime As Double, ByRef histoSums As Variant, ByRef histoSquares As Variant, lHistoBinCount As Long

    Dim dblTotalWidthSecs As Double
    Dim dblBinWidthSecs As Double
    Dim dblStartOffsetSecs As Double
    
    Dim histoSums() As Variant
    Dim histoSquares() As Variant
    Dim histoN As Long
    Dim histoBinCount As Long
    
    Dim iStimNum As Integer
    Dim returnVal As Variant
    
    'FIRST: check for onset spike
    
    'create 3 bins, each 0.01 wide, to check for an onset spike
    dblTotalWidthSecs = 0.05
    dblBinWidthSecs = 0.01
    dblStartOffsetSecs = 0#
    
    histoBinCount = CInt(dblTotalWidthSecs / dblBinWidthSecs)
    Call setHistoArraySizes(histoSums, histoSquares, histoBinCount)
    
    For iStimNum = 0 To 8 'only want to look at the first 9 stims, because after than the shock will be on, which could screw up the neural data
        Call buildHistogramForStimMethod1(stimEpocs(1, iStimNum), histoSums, histoSquares, histoBinCount, dblTotalWidthSecs, dblBinWidthSecs, dblStartOffsetSecs)
    Next
    
    Dim vChanKey As Variant
    Dim iChanOffset As Integer
    'step through each channel
    For Each vChanKey In dictOnlyIncludeChannels.Keys
        iChanOffset = dictOnlyIncludeChannels(vChanKey) - 1
        'do the actual check - check if the first 10ms bin is greater than each of the four subsequent bins
        If histoSums(iChanOffset)(0) > histoSums(iChanOffset)(1) And _
            histoSums(iChanOffset)(0) > histoSums(iChanOffset)(2) And _
            histoSums(iChanOffset)(0) > histoSums(iChanOffset)(3) And _
            histoSums(iChanOffset)(0) > histoSums(iChanOffset)(4) And _
            histoSums(iChanOffset)(0) > histoSums(iChanOffset)(5) Then
                Call dDrivenChanList.Add(vChanKey, DriveDetect_OnsetDetected)
        End If
    Next
            
    'SECOND: check for higher overall activity in stim period than non-stim period
    
    'create 4 bins, each 0.1 wide, to check for greater activity during the 'in-tone' period than in the 'no-tone' period
    dblTotalWidthSecs = 0.4
    dblBinWidthSecs = 0.1
    dblStartOffsetSecs = 0#
    
    histoBinCount = CInt(dblTotalWidthSecs / dblBinWidthSecs)
    'flush the arrays
    Call setHistoArraySizes(histoSums, histoSquares, histoBinCount)
    
    For iStimNum = 0 To 8 'only want to look at the first 9 stims, because after than the shock will be on, which could screw up the neural data
        Call buildHistogramForStimMethod1(stimEpocs(1, iStimNum), histoSums, histoSquares, histoBinCount, dblTotalWidthSecs, dblBinWidthSecs, dblStartOffsetSecs)
    Next
    
    'step through each channel
    For Each vChanKey In dictOnlyIncludeChannels.Keys
        iChanOffset = dictOnlyIncludeChannels(vChanKey) - 1
        'do the actual check - check if the first 10ms bin is greater than each of the four subsequent bins
        If histoSums(iChanOffset)(0) > (histoSums(iChanOffset)(4) * DriveDetect_ActivityDifferenceThreshold) Then
                If Not dDrivenChanList.Exists(vChanKey) Then
                    Call dDrivenChanList.Add(vChanKey, DriveDetect_ActDiffDetected)
                End If
        End If
    Next
    
    For Each vChanKey In dictOnlyIncludeChannels.Keys
        If Not dDrivenChanList.Exists(vChanKey) Then
            Call dDrivenChanList.Add(vChanKey, DriveDetect_Undriven)
        End If
    Next
    
End Function


