Attribute VB_Name = "GenerateChanHistogram"
Option Explicit

Const iChartHeight = 4
Const iChartWidth = 120

Function generateChanHistograms( _
        objTTX As TTankX, _
        cfWS As Worksheet, _
        outputWS As Worksheet, _
        xAxisEp As String, _
        yAxisEp As String, _
        vXAxisKeys As Variant, _
        vYAxisKeys As Variant, _
        lNumOfChans As Long, _
        stimStartEpoc As String, _
        Optional vChannelMapper As Variant, _
        Optional dChannelsToArrayMapping As Variant _
    )
    Dim dblBinWidthSecsForHisto As Double
    dblBinWidthSecsForHisto = 0.001
    
    Dim bBuildAllCharts As Boolean
    bBuildAllCharts = True
    Dim bChartsInElectrodeArrangement As Boolean
    bChartsInElectrodeArrangement = True
    
    Dim myChart As ChartObject
    
    Dim dblTotalWidthSecs As Double
    Dim dblStartOffsetSecs As Double
    
    Dim histoSums() As Variant
    Dim histoSquares() As Variant
    Dim histoN As Long
    Dim lHistoBinCount As Long

    Dim iStimNum As Integer
    
    'create bins based of provided configuration parameters to check for an onset spike
    'dblIgnoreFirstMsec = dblIgnoreFirstMsec
    'dblBinWidthSecs =
    dblTotalWidthSecs = dblBinWidth * 2
        
    lHistoBinCount = calcBinCount(dblTotalWidthSecs - dblIgnoreFirstMsec, dblBinWidthSecsForHisto)
    
    Dim lChanNum As Long
    lChanNum = 1
    Dim lCF As Long

    Dim sStableSearchString As String
    Dim sThisSearchString As String
    Dim vStimEpocs As Variant
    Dim aStimTimes() As Double
    Dim lArrIndex As Long
    Dim lStimIter As Long
    Dim lChartNum As Long
    Dim lChartTopPos As Long
    Dim lChartHeight As Long
    Dim lChartLeftPos As Long
    Dim lRowNum As Long
    lChartNum = 0
    
    outputWS.Cells(1, 1) = "Channel"
    
    Dim lMaxAmp As Long
    For lArrIndex = 0 To UBound(vYAxisKeys)
        If vYAxisKeys(lArrIndex) > lMaxAmp Then
            lMaxAmp = vYAxisKeys(lArrIndex)
        End If
    Next
    
    If bChartsInElectrodeArrangement Then
        For lArrIndex = 0 To (lHistoBinCount - 1)
            outputWS.Cells(1, lArrIndex + 2) = dblIgnoreFirstMsec + (dblBinWidthSecsForHisto * lArrIndex)
        Next
    End If
    
    Dim lFreqMidpoint As Long
    lFreqMidpoint = vXAxisKeys(CInt(UBound(vXAxisKeys) / 2))
    
    While cfWS.Cells(lChanNum + 1, 1).Value <> ""
        lCF = cfWS.Cells(lChanNum + 1, 2).Value
        
        If bChartsInElectrodeArrangement Then
            lRowNum = lChanNum + 1
        Else
            lRowNum = lChartNum * iChartHeight + 2 + lChanNum + 2
        End If
        
        outputWS.Cells(lRowNum, 1) = lChanNum
        
        Call setHistoArraySizes(histoSums, histoSquares, lHistoBinCount, lNumOfChans)
        
        If lCF <> 0 Then
            sThisSearchString = yAxisEp & " = " & lMaxAmp & " and (" & xAxisEp & " = " & CStr(lCF) & " or " & xAxisEp & " = " & CStr(lCF - 1000) & " or " & xAxisEp & " = " & CStr(lCF + 1000) & ")"
        Else
            sThisSearchString = yAxisEp & " = " & lMaxAmp & " and (" & xAxisEp & " = " & CStr(lFreqMidpoint) & " or " & xAxisEp & " = " & CStr(lFreqMidpoint - 1000) & " or " & xAxisEp & " = " & CStr(lFreqMidpoint + 1000) & ")"
        End If
        
        If bBuildAllCharts Or lCF <> 0 Then
            Call objTTX.ResetFilters
            Call objTTX.SetFilterWithDescEx(sThisSearchString)
            vStimEpocs = objTTX.GetEpocsExV(stimStartEpoc, 0)
            If Not IsEmpty(vStimEpocs) Then
                ReDim aStimTimes(UBound(vStimEpocs, 2))
                For lStimIter = 0 To UBound(vStimEpocs, 2)
                    aStimTimes(lStimIter) = vStimEpocs(1, lStimIter)
                Next
                
                For iStimNum = 0 To UBound(aStimTimes)
                    Call buildHistogramForStim(objTTX, aStimTimes(iStimNum) + dblIgnoreFirstMsec, histoSums, histoSquares, dblTotalWidthSecs, dblBinWidthSecsForHisto, vChannelMapper, dChannelsToArrayMapping, 0)
                Next
            End If
            
            For lArrIndex = 0 To (lHistoBinCount - 1)
                If Not bChartsInElectrodeArrangement Then
                    outputWS.Cells(lRowNum, lArrIndex + 2) = dblIgnoreFirstMsec + (dblBinWidthSecsForHisto * lArrIndex)
                    outputWS.Cells(lRowNum + 1, lArrIndex + 2) = histoSums(vChannelMapper.revLookup(lChanNum) - 1)(lArrIndex) / UBound(aStimTimes)
                Else
                    outputWS.Cells(lRowNum, lArrIndex + 2) = histoSums(vChannelMapper.revLookup(lChanNum) - 1)(lArrIndex) / UBound(aStimTimes)
                End If
            Next
            
            If bChartsInElectrodeArrangement Then
                If lChanNum < 17 Then
                    lChartTopPos = outputWS.Range(outputWS.Cells(lNumOfChans + 1 + (lChanNum - 1) * iChartHeight + 1, 1), outputWS.Cells(lNumOfChans + 1 + lChanNum * iChartHeight, 1)).Top
                    lChartHeight = outputWS.Range(outputWS.Cells(lNumOfChans + 1 + (lChanNum - 1) * iChartHeight + 1, 1), outputWS.Cells(lNumOfChans + 1 + lChanNum * iChartHeight, 1)).Height
                    lChartLeftPos = outputWS.Range(outputWS.Cells(lNumOfChans + 1 + (lChanNum - 1) * iChartHeight + 1, 1), outputWS.Cells(lNumOfChans + 1 + lChanNum * iChartHeight, 1)).Left
                Else
                    lChartTopPos = outputWS.Range(outputWS.Cells(lNumOfChans + 1 + (lChanNum - 17) * iChartHeight + 1, 1), outputWS.Cells(lNumOfChans + 1 + (lChanNum - 16) * iChartHeight, 1)).Top
                    lChartHeight = outputWS.Range(outputWS.Cells(lNumOfChans + 1 + (lChanNum - 17) * iChartHeight + 1, 1), outputWS.Cells(lNumOfChans + 1 + (lChanNum - 16) * iChartHeight, 1)).Height
                    lChartLeftPos = outputWS.Range(outputWS.Cells(lNumOfChans + 1 + (lChanNum - 17) * iChartHeight + 1, 1), outputWS.Cells(lNumOfChans + 1 + (lChanNum - 16) * iChartHeight, 1)).Left + iChartWidth
                End If
            Else
                lChartTopPos = outputWS.Range(outputWS.Cells(lRowNum + 2, 2), outputWS.Cells(lRowNum + iChartHeight, 2)).Top
                lChartHeight = outputWS.Range(outputWS.Cells(lRowNum + 2, 2), outputWS.Cells(lRowNum + iChartHeight, 2)).Height
                lChartLeftPos = outputWS.Range(outputWS.Cells(lRowNum + 2, 2), outputWS.Cells(lRowNum + iChartHeight, 2)).Left
            End If
            
            Set myChart = outputWS.ChartObjects.Add(lChartLeftPos, lChartTopPos, iChartWidth, lChartHeight)
            myChart.Chart.ChartType = xlColumnClustered
            If Not bChartsInElectrodeArrangement Then
                Call myChart.Chart.SetSourceData(outputWS.Range(outputWS.Cells(lRowNum + 1, 2), outputWS.Cells(lRowNum + 1, 2 + lHistoBinCount)))
            Else
                Call myChart.Chart.SetSourceData(outputWS.Range(outputWS.Cells(lRowNum, 2), outputWS.Cells(lRowNum, 2 + lHistoBinCount)))
            End If
            myChart.Chart.ChartGroups(1).GapWidth = 0
            If lCF <> 0 Then
                myChart.Chart.SeriesCollection(1).Name = "Channel " & lChanNum & " (" & lCF & ")"
            Else
                myChart.Chart.SeriesCollection(1).Name = "Channel " & lChanNum
            End If
'            myChart.Chart.SeriesCollection(1).XValues = outputWS.Range(outputWS.Cells(1, 2), outputWS.Cells(1, 2 + lHistoBinCount))
            myChart.Chart.Axes(xlCategory).Delete
            myChart.Chart.Axes(xlValue).Delete
            myChart.Chart.SeriesCollection(1).XValues = outputWS.Range(outputWS.Cells(1, 2), outputWS.Cells(1, 2 + lHistoBinCount))
            myChart.Chart.SeriesCollection(1).Format.Line.Style = msoLineSingle
            myChart.Chart.SeriesCollection(1).Format.Line.Weight = 0.25
            myChart.Chart.SeriesCollection(1).Format.Line.Visible = msoFalse
            If lCF = 0 Then
                myChart.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(230, 185, 184)
            Else
                myChart.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(54, 156, 73)
            End If
            myChart.Chart.Legend.Delete
            myChart.Chart.ChartTitle.Characters.Font.Size = 7
            
            lChartNum = lChartNum + 1
        End If

        lChanNum = lChanNum + 1
    Wend
End Function




