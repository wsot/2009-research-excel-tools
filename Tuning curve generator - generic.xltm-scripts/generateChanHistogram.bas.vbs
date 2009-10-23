Attribute VB_Name = "GenerateChanHistogram"
Option Explicit

Function generateChanHistograms( _
        objTTX As TTankX, _
        cfWS As Worksheet, _
        outputWS As Worksheet, _
        xAxisEp As String, _
        lNumOfChans As Long, _
        stimStartEpoc As String, _
        Optional vChannelMapper As Variant, _
        Optional dChannelsToArrayMapping As Variant _
    )
    Dim dblBinWidthSecsForHisto As Double
    dblBinWidthSecsForHisto = 0.001
    
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
    dblTotalWidthSecs = dblBinWidth
        
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
    lChartNum = 0
    
    outputWS.Cells(1, 1) = "Channel"
    
    While cfWS.Cells(lChanNum + 1, 1).Value <> ""
        lCF = cfWS.Cells(lChanNum + 1, 2).Value
        outputWS.Cells(lChartNum * 21 + lChanNum + 2, 1) = lChanNum
        
        If lCF <> 0 Then
            Call setHistoArraySizes(histoSums, histoSquares, lHistoBinCount, lNumOfChans)
            sThisSearchString = xAxisEp & " = " & CStr(lCF) & " or " & xAxisEp & " = " & CStr(lCF - 1000) & " or " & xAxisEp & " = " & CStr(lCF + 1000)
            Call objTTX.ResetFilters
            Call objTTX.SetFilterWithDescEx(sThisSearchString)
            vStimEpocs = objTTX.GetEpocsExV(stimStartEpoc, 0)
            If Not IsEmpty(vStimEpocs) Then
                ReDim aStimTimes(UBound(vStimEpocs, 2))
                For lStimIter = 0 To UBound(vStimEpocs, 2)
                    aStimTimes(lStimIter) = vStimEpocs(1, lStimIter)
                Next
                
                For iStimNum = 0 To UBound(aStimTimes)
                    Call buildHistogramForStim(objTTX, aStimTimes(iStimNum) + dblIgnoreFirstMsec, histoSums, histoSquares, dblTotalWidthSecs, dblBinWidthSecsForHisto, vChannelMapper, dChannelsToArrayMapping)
                Next
            End If
            
            For lArrIndex = 0 To (lHistoBinCount - 1)
                outputWS.Cells(lChartNum * 21 + lChanNum + 1, lArrIndex + 2) = dblIgnoreFirstMsec + (dblBinWidthSecsForHisto * lArrIndex)
                outputWS.Cells(lChartNum * 21 + lChanNum + 2, lArrIndex + 2) = histoSums(vChannelMapper.revLookup(lChanNum) - 1)(lArrIndex) / UBound(aStimTimes)
            Next
            
            lChartTopPos = outputWS.Range(outputWS.Cells(lChartNum * 21 + lChanNum + 3, 2), outputWS.Cells((lChartNum + 1) * 21 + lChanNum, 2)).Top
            lChartHeight = outputWS.Range(outputWS.Cells(lChartNum * 21 + lChanNum + 3, 2), outputWS.Cells((lChartNum + 1) * 21 + lChanNum, 2)).Height
            Set myChart = outputWS.ChartObjects.Add(1, lChartTopPos, 500, lChartHeight)
            myChart.Chart.ChartType = xlColumnClustered
            Call myChart.Chart.SetSourceData(outputWS.Range(outputWS.Cells(lChartNum * 21 + lChanNum + 2, 2), outputWS.Cells(lChartNum * 21 + lChanNum + 2, 2 + lHistoBinCount)))
            myChart.Chart.ChartGroups(1).GapWidth = 0
            myChart.Chart.SeriesCollection(1).Name = "Channel " & lChanNum & " (" & lCF & ")"
            myChart.Chart.SeriesCollection(1).XValues = outputWS.Range(outputWS.Cells(lChartNum * 21 + lChanNum + 1, 2), outputWS.Cells(lChartNum * 21 + lChanNum + 1, 2 + lHistoBinCount))
            myChart.Chart.SeriesCollection(1).Format.Line.Style = msoLineSingle
            myChart.Chart.SeriesCollection(1).Format.Line.Weight = 0.25
            myChart.Chart.SeriesCollection(1).Format.Line.Visible = msoTrue
            myChart.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(247, 150, 70)
            myChart.Chart.Legend.Delete
            myChart.Chart.ChartTitle.Characters.Font.Size = 12
            
            lChartNum = lChartNum + 1
        End If
        lChanNum = lChanNum + 1
    Wend
End Function

