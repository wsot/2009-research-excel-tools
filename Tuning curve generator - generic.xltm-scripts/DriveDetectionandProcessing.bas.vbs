Attribute VB_Name = "DriveDetectionandProcessing"
Option Explicit

Global Const DDParam_Gen_MinSpikesPerEpocInFirstN = 0 'Absolute minimum spikes per epoc in intital time window duration
Global Const DDParam_Gen_FirstNDur = 1 'Duration of initial time window (s)
Global Const DDParam_Onset_IgnoreTime = 2 'Bin Width (s) for 'onset detection' (First bit compared to subsequent bins)
Global Const DDParam_Onset_BinWidth = 3 'Bin Width (s) for 'onset detection' (First bit compared to subsequent bins)
Global Const DDParam_Onset_ReqMultiple = 4 'Onset spike must be at least x times larger than following spikes for onset spike detection
Global Const DDParam_Onset_NumComparBins = 5 'Number of subsequent bins to compare onset spike to
Global Const DDParam_Onset_MinSpikesPerEpocInComparBins = 6 'Minimum total spikes (per epoc) in comparison bins (the number of bins specified above)
Global Const DDParam_Diff_Threshold = 7 'During-tone vs outside-tone activity difference threshold (ratio inside/outside) for inclusion as 'driven'
Global Const DDParam_Diff_StimDur = 8 'Tone duration (s)
Global Const DDParam_Diff_ITI__ = 9 'Inter-tone interval (s; including the duration of the tone)


Function identifyDrivenChannels(arrStimTimes As Variant, vDriveDetectionParams As Variant, Optional dChannelList As Variant) As Variant 'return dictionary of channel numbers that are driven

    Dim dblTotalWidthSecs As Double
    Dim dblBinWidthSecs As Double
    Dim dblStartOffsetSecs As Double
    
    Dim histoSums() As Variant
    Dim histoSquares() As Variant
    Dim histoN As Long
    Dim lHistoBinCount As Long
    
    Dim iStimNum As Integer
    Dim returnVal As Variant
    
    'FIRST: check for onset spike
    
    'create 3 bins, each 0.01 wide, to check for an onset spike
    dblBinWidthSecs = vDriveDetectionParams(DDParam_Onset_BinWidth)
    dblStartOffsetSecs = vDriveDetectionParams(DDParam_Onset_IgnoreTime)
    dblTotalWidthSecs = dblBinWidthSecs * (vDriveDetectionParams(DDParam_Onset_NumComparBins) + 1)
    
    lHistoBinCount = calcBinCount(dblTotalWidthSecs, dblBinWidthSecs)
    Call setHistoArraySizes(histoSums, histoSquares, lHistoBinCount)
    
    For iStimNum = 0 To UBound(arrStimTimes)
        Call buildHistogramForStim(arrStimTimes(iStimNum) + dblStartOffsetSecs, histoSums, histoSquares, lHistoBinCount, dblTotalWidthSecs, dblBinWidthSecs, dblStartOffsetSecs)
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
            histoSums(iChanOffset)(0) > histoSums(iChanOffset)(5) And _
            (histoSums(iChanOffset)(1) + histoSums(iChanOffset)(2)) >= DriveDetect_MinIn2nd3rdForOnset Then
                Call dDrivenChanList.Add(vChanKey, DriveDetect_OnsetDetected)
        End If
    Next
            
    'SECOND: check for higher overall activity in stim period than non-stim period
    
    'create 4 bins, each 0.1 wide, to check for greater activity during the 'in-tone' period than in the 'no-tone' period
    dblTotalWidthSecs = 0.4
    dblBinWidthSecs = 0.1
    dblStartOffsetSecs = 0#
    
    lHistoBinCount = calcBinCount(dblTotalWidthSecs, dblBinWidthSecs)
    'flush the arrays
    Call setHistoArraySizes(histoSums, histoSquares, lHistoBinCount)
    
    For iStimNum = 0 To 8 'only want to look at the first 9 stims, because after than the shock will be on, which could screw up the neural data
        Call buildHistogramForStimMethod1(stimEpocs(1, iStimNum), histoSums, histoSquares, lHistoBinCount, dblTotalWidthSecs, dblBinWidthSecs, dblStartOffsetSecs)
    Next
    
    'step through each channel
    For Each vChanKey In dictOnlyIncludeChannels.Keys
        iChanOffset = dictOnlyIncludeChannels(vChanKey) - 1
        'do the actual check - check if the first 10ms bin is greater than each of the four subsequent bins
        If (histoSums(iChanOffset)(0) > (histoSums(iChanOffset)(4) * DriveDetect_ActivityDifferenceThreshold)) And (histoSums(iChanOffset)(0) > DriveDetect_AbsoluteMinimumSpikesInFirstBin) Then
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

'creates arrays the right size for the histogram data
Function setHistoArraySizes(ByRef histoSums As Variant, ByRef histoSquares As Variant, ByRef lHistoBinCount As Long)
    Dim i As Long
    
    Dim arrDoubles() As Double
        
    ReDim histoSums(dictOnlyIncludeChannels.Count - 1)
    ReDim histoSquares(dictOnlyIncludeChannels.Count - 1)
    
    'ReDim arrVariants(dictOnlyIncludeChannels.Count - 1)
    
    ReDim arrDoubles(lHistoBinCount)
    
    For i = 0 To dictOnlyIncludeChannels.Count - 1
        histoSums(i) = arrDoubles
        histoSquares(i) = arrDoubles
    Next
End Function

Function buildHistogramForStim( _
        ByVal dblStartTime As Double, _
        ByRef histoSums As Variant, _
        ByRef histoSquares As Variant, _
        ByRef dblTotalWidthSecs As Double, _
        ByRef dblBinWidthSecs As Double, _
        Optional ByRef dChannelRemapping As Variant, _
        Optional ByRef dChannelsToArrayMapping As Variant _
        )
    
    Dim lHistoBinCount As Long
    lHistoBinCount = calcBinCount(dblTotalWidthSecs, dblBinWidthSecs)
    
    Dim iChanNum As Integer 'used for iteration in generating histos for each channel
    Dim lEvtCount As Long 'count of total events retrieved with the current search
    Dim lEvtNum As Long 'iterator for stepping through retreived records
    Dim lBinNum As Long 'the current bin being collected for
    
    Dim dblEndTime As Double 'the end time to search to
    Dim varData As Variant 'data returned from ParseEvInfoV
           
    'Dim dCount As Dictionary
    'Set dCount = New Dictionary
    Dim arrCount() As Long
    ReDim arrCount(200) 'because in reality redimming, especially with preserve, is a very expensive operation we're better off just starting off with a bigger number
    Dim intArrCountUpperLimit As Integer
    intArrCountUpperLimit = UBound(arrCount)


    'check if channel remapping is required
    'for the remapping table, the first value (key) needs to be the TDT CHANNEL RECORDED, and the second value the DESIRED NEW LABEL
    Dim blnRemapChannels As Boolean
    If Not IsMissing(dChannelRemapping) Then
        blnRemapChannels = True
    Else
        blnRemapChannels = False
    End If

    Dim blnRemapToArray As Boolean
    If Not IsMissing(dChannelsToArrayMapping) Then
        blnRemapToArray = True
    Else
        blnRemapToArray = False
    End If

    Dim iWriteToChan As Integer

    dblEndTime = dblStartTime + dblBinWidthSecs
    For lBinNum = 0 To lHistoBinCount
            Do
                lEvtCount = objTTX.ReadEventsV(100000, "CSPK", 0, 0, dblStartTime, dblEndTime, "ALL")
                If lEvtCount = 0 Then
                    Exit Do
                End If
            
                varData = objTTX.ParseEvInfoV(0, lEvtCount, 4)
            
                For lEvtNum = 0 To lEvtCount - 1
                    'count the number of events for each channel in the current bin
                    If (varData(lEvtNum) - 1) > intArrCountUpperLimit Then 'maybe cheaper than actually doing a ubound every time? (but who really knows...)
                        ReDim Preserve arrCount(varData(lEvtNum) - 1)
                        intArrCountUpperLimit = UBound(arrCount)
                    End If
                    arrCount(varData(lEvtNum) - 1) = arrCount(varData(lEvtNum) - 1) + 1
                Next
    
                'if the full 10000 was retrieved, there may be more to fetch, so try to fetch them
                If lEvtCount < 100000 Then
                    Exit Do
                Else
                    'get the time of the last event, and search forward from that - there is a risk this could miss events where the time is identical, however. That said, never got more than 10000 event yet
                    MsgBox "Obtained 100000+ events!"
                    varData = objTTX.ParseEvInfoV(lEvtCount - 1, 1, 6)
                    dblStartTime = varData(0) + (1 / 100000)
                End If
            Loop
            
            'update the totals with the obtained number of events
            For iChanNum = 1 To UBound(arrCount) + 1
                iWriteToChan = iChanNum
                If blnRemapChannels Then
                    If dChannelRemapping.Exists(iWriteToChan) Then
                        iWriteToChan = dChannelRemapping(iWriteToChan)
                    Else
                        iWriteToChan = 0
                    End If
                End If
                
                If blnRemapToArray Then
                    If dChannelsToArrayMapping.Exists(iWriteToChan) Then
                        iWriteToChan = dChannelsToArrayMapping(iWriteToChan) + 1 '-1 because the channel number has 1 subtracted from it to make it an array index later
                    Else
                        iWriteToChan = 0
                    End If
                End If
                
                If iWriteToChan > (UBound(histoSums, 1) + 1) Then 'if the channel number is higher than the array can actually support, then need to scrap it
                    iWriteToChan = 0
                End If
                
                If iWriteToChan <> 0 Then 'if iWriteToChan is 0, then the value will be ignored
                    histoSums(iWriteToChan - 1)(lBinNum) = histoSums(iWriteToChan)(lBinNum) + nCount(iChanNum)
                    histoSquares(iWriteToChan - 1)(lBinNum) = histoSquares(iWriteToChan)(lBinNum) + (nCount(iChanNum) ^ 2)
                End If
            Next
            ReDim nCount(intArrCountUpperLimit) 'clear the storage array, but keep it with the same number of channels

            dblStartTime = dblEndTime
            dblEndTime = dblStartTime + dblBinWidthSecs
        Next

End Function
Function calcBinCount(dblTotalWidthSecs As Double, dblBinWidthSecs As Double) As Long
    calcBinCount = CLng(dblTotalWidthSecs / dblBinWidthSecs)
End Function
