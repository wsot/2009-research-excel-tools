Attribute VB_Name = "DriveDetectionandProcessing"
Option Explicit

'Const useHistoGenType = 1

Const DriveDetect_Undriven = 0
Const DriveDetect_MinimumSpikesCrossed = 1
Const DriveDetect_OnsetDetected = 2
Const DriveDetect_ActDiffDetected = 4

'adds responding channel numbers to dDrivenChanList
'this will screw up if the vChannelCount is actually less than the REAL number of channels in the system
Function identifyDrivenChannels( _
        objTTX As TTankX, _
        arrStimTimes As Variant, _
        oDriveDetectionParams As DriveDetection, _
        ByRef dDrivenChanList As Variant, _
        Optional vChannelCount As Variant, _
        Optional vChannelMapper As Variant, _
        Optional dChannelsToArrayMapping As Variant _
    )

    If IsMissing(vChannelCount) Or Not IsNumeric(vChannelCount) Then
        If Not IsMissing(dChannelsToArrayMapping) Then
            If IsObject(dChannelsToArrayMapping) Then
                If Not (dChannelsToArrayMapping Is Nothing) Then
                    vChannelCount = dChannelsToArrayMapping.Count
                End If
            End If
        End If
    ElseIf Not Int(vChannelCount) = vChannelCount Then
        If Not IsMissing(dChannelsToArrayMapping) Then
            If IsObject(dChannelsToArrayMapping) Then
                If Not (dChannelsToArrayMapping Is Nothing) Then
                    vChannelCount = dChannelsToArrayMapping.Count
                End If
            End If
        End If
    End If

    If dDrivenChanList Is Nothing Then
        Set dDrivenChanList = New Dictionary
    End If

    Dim dblTotalWidthSecs As Double
    Dim dblBinWidthSecs As Double
    Dim dblStartOffsetSecs As Double
    
    Dim histoSums() As Variant
    Dim histoSquares() As Variant
    Dim histoN As Long
    Dim lHistoBinCount As Long
    
    Dim iStimNum As Integer
    Dim returnVal As Variant

    'FIRST: check for initital spikes per epoc requirement
    
    'create bins based of provided configuration parameters to check for an onset spike
    dblBinWidthSecs = oDriveDetectionParams.Gen_FirstNDur
    dblTotalWidthSecs = dblBinWidthSecs
    
    lHistoBinCount = calcBinCount(dblTotalWidthSecs, dblBinWidthSecs)
    Call setHistoArraySizes(histoSums, histoSquares, lHistoBinCount, CLng(vChannelCount))
    
    For iStimNum = 0 To UBound(arrStimTimes)
        Call buildHistogramForStim(objTTX, arrStimTimes(iStimNum) + oDriveDetectionParams.IgnoreInitialTime, histoSums, histoSquares, dblTotalWidthSecs, dblBinWidthSecs, vChannelMapper, dChannelsToArrayMapping)
    Next
    
    Dim blnChanIsDriven As Boolean
    Dim dblSpikePerEpoc As Double
    
    Dim lArrIndx As Long
    Dim lComparisonBin As Long

    'step through each channel
    For lArrIndx = 0 To (UBound(histoSums))
        blnChanIsDriven = True
        'do the actual check - check if the first 10ms bin is greater than each of the four subsequent bins
        If Not histoSums(lArrIndx)(0) / (UBound(arrStimTimes) + 1) > oDriveDetectionParams.Gen_MinSpikesPerEpocInFirstN Then
                blnChanIsDriven = False
        End If
        If blnChanIsDriven Then
            If Not dDrivenChanList.Exists(lArrIndx + 1) Then
                Call dDrivenChanList.Add(lArrIndx + 1, DriveDetect_MinimumSpikesCrossed)
            Else
                dDrivenChanList(lArrIndx + 1) = dDrivenChanList(lArrIndx + 1) Or DriveDetect_MinimumSpikesCrossed
            End If
        End If
    Next

    'SECOND: check for onset spike
    
    'create bins based of provided configuration parameters to check for an onset spike
    dblBinWidthSecs = oDriveDetectionParams.Onset_BinWidth
    dblTotalWidthSecs = dblBinWidthSecs * (oDriveDetectionParams.Onset_NumComparBins + 1)
    
    lHistoBinCount = calcBinCount(dblTotalWidthSecs, dblBinWidthSecs)
    Call setHistoArraySizes(histoSums, histoSquares, lHistoBinCount, CLng(vChannelCount))
    
    For iStimNum = 0 To UBound(arrStimTimes)
        Call buildHistogramForStim(objTTX, arrStimTimes(iStimNum) + oDriveDetectionParams.IgnoreInitialTime, histoSums, histoSquares, dblTotalWidthSecs, dblBinWidthSecs, vChannelMapper, dChannelsToArrayMapping)
    Next
    
    'step through each channel
    For lArrIndx = 0 To (UBound(histoSums))
        If dDrivenChanList.Exists(lArrIndx + 1) Then 'if this doesn't exist, the intital onset drive has not been detected, so can not be counted to have drive
            blnChanIsDriven = True
            dblSpikePerEpoc = 0#
            'do the actual check - check if the first 10ms bin is greater than each of the four subsequent bins
            For lComparisonBin = 1 To 1 + oDriveDetectionParams.Onset_NumComparBins
                If Not histoSums(lArrIndx)(0) > (histoSums(lArrIndx)(lComparisonBin) * oDriveDetectionParams.Onset_ReqMultiple) Then
                    blnChanIsDriven = False
                    Exit For
                End If
                dblSpikePerEpoc = dblSpikePerEpoc + histoSums(lArrIndx)(lComparisonBin)
            Next
            If blnChanIsDriven Then
                If (dblSpikePerEpoc / oDriveDetectionParams.Onset_NumComparBins) > oDriveDetectionParams.Onset_MinSpikesPerEpocInComparBins Then
                    dDrivenChanList(lArrIndx + 1) = dDrivenChanList(lArrIndx + 1) Or DriveDetect_OnsetDetected
                Else
                    blnChanIsDriven = False
                End If
            End If
        End If
    Next
    
    'THIRD: check for higher overall activity in stim period than non-stim period
    
    'create 4 bins, each 0.1 wide, to check for greater activity during the 'in-tone' period than in the 'no-tone' period
    dblTotalWidthSecs = oDriveDetectionParams.Diff_ITI
    dblBinWidthSecs = oDriveDetectionParams.Diff_StimDur
    lHistoBinCount = calcBinCount(dblTotalWidthSecs, dblBinWidthSecs)
    
    'flush the arrays
    Call setHistoArraySizes(histoSums, histoSquares, lHistoBinCount, CLng(vChannelCount))
    
    For iStimNum = 0 To UBound(arrStimTimes)
        Call buildHistogramForStim(objTTX, arrStimTimes(iStimNum) + oDriveDetectionParams.IgnoreInitialTime, histoSums, histoSquares, dblTotalWidthSecs, dblBinWidthSecs, vChannelMapper, dChannelsToArrayMapping)
    Next
    
    'step through each channel
    For lArrIndx = 0 To (UBound(histoSums))
        If dDrivenChanList.Exists(lArrIndx + 1) Then 'if this doesn't exist, the intital onset drive has not been detected, so can not be counted to have drive
            If Not (dDrivenChanList(lArrIndx + 1) And DriveDetect_OnsetDetected) Then
                If histoSums(lArrIndx)(0) > (histoSums(lArrIndx)(1) * oDriveDetectionParams.Diff_Threshold) Then
                    dDrivenChanList(lArrIndx + 1) = dDrivenChanList(lArrIndx + 1) Or DriveDetect_ActDiffDetected
                End If
            End If
        End If
    Next
    
        
    'step through each channel
    For lArrIndx = 0 To (UBound(histoSums))
        If dDrivenChanList.Exists(lArrIndx + 1) Then 'if this doesn't exist, the intital onset drive has not been detected, so can not be counted to have drive
            If dDrivenChanList(lArrIndx + 1) = DriveDetect_MinimumSpikesCrossed Then 'if only detected by threshold crossing, so should be removed
                Call dDrivenChanList.Remove(lArrIndx + 1)
            End If
        End If
    Next
    
    'If Not IsMissing(dChannelsToArrayMapping) And IsObject(dChannelsToArrayMapping) And Not (dChannelsToArrayMapping Is Nothing) Then
        'need to reverse-convert index numbers to channel numbers, because they may not be the same
    'End If
    Set identifyDrivenChannels = dDrivenChanList
    
End Function

'creates arrays the right size for the histogram data.
Function setHistoArraySizes( _
        ByRef histoSums As Variant, _
        ByRef histoSquares As Variant, _
        lHistoBinCount As Long, _
        lChanCount As Long _
    )
    
    Dim i As Long
    
    Dim arrDoubles() As Double
        
    ReDim histoSums(lChanCount - 1)
    ReDim histoSquares(lChanCount - 1)
        
    ReDim arrDoubles(lHistoBinCount)
    
    For i = 0 To lChanCount - 1
        histoSums(i) = arrDoubles
        histoSquares(i) = arrDoubles
    Next
End Function

Function buildHistogramForStim( _
        objTTX As TTankX, _
        ByVal dblStartTime As Double, _
        ByRef histoSums As Variant, _
        ByRef histoSquares As Variant, _
        ByRef dblTotalWidthSecs As Double, _
        ByRef dblBinWidthSecs As Double, _
        Optional vChannelMapper As Variant, _
        Optional dChannelsToArrayMapping As Variant, _
        Optional vHistoGenType As Variant _
        )
    
    Dim iHistoGenType As Integer
    iHistoGenType = 0
    If Not IsMissing(vHistoGenType) Then
        If IsNumeric(vHistoGenType) Then
            iHistoGenType = vHistoGenType
        End If
    End If
    
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
    ReDim arrCount(31) 'because in reality redimming, especially with preserve, is a very expensive operation we're better off just starting off with a bigger number
    Dim intArrCountUpperLimit As Integer
    intArrCountUpperLimit = UBound(arrCount)


    'check if channel remapping is required
    'for the remapping table, the first value (key) needs to be the TDT CHANNEL RECORDED, and the second value the DESIRED NEW LABEL
    Dim blnRemapChannels As Boolean
    If Not IsMissing(vChannelMapper) Then
        If IsObject(vChannelMapper) Then
            If Not (vChannelMapper Is Nothing) Then
                blnRemapChannels = True
            End If
        End If
    Else
        blnRemapChannels = False
    End If

    Dim blnRemapToArray As Boolean
    If Not IsMissing(dChannelsToArrayMapping) Then
        If IsObject(dChannelsToArrayMapping) Then
            If Not (dChannelsToArrayMapping Is Nothing) Then
                blnRemapToArray = True
            End If
        End If
    Else
        blnRemapToArray = False
    End If

    Dim iWriteToChan As Integer

    dblEndTime = dblStartTime + dblBinWidthSecs
    For lBinNum = 0 To lHistoBinCount
            Select Case iHistoGenType
                Case 0:
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
                Case 1:
                    For iChanNum = 1 To UBound(arrCount) + 3
                        lEvtCount = objTTX.ReadEventsV(100000, "CSPK", iChanNum, 0, dblStartTime, dblEndTime, "ALL")
                        If Not lEvtCount = 0 Then
                            If (iChanNum - 1) > intArrCountUpperLimit Then 'maybe cheaper than actually doing a ubound every time? (but who really knows...)
                                ReDim Preserve arrCount(iChanNum - 1)
                                intArrCountUpperLimit = UBound(arrCount)
                            End If
                            arrCount(iChanNum - 1) = arrCount(iChanNum - 1) + lEvtCount
                        
                            'if the full 10000 was retrieved, there may be more to fetch, so try to fetch them
                            If Not lEvtCount < 100000 Then
                                'get the time of the last event, and search forward from that - there is a risk this could miss events where the time is identical, however. That said, never got more than 10000 event yet
                                MsgBox "Obtained 100000+ events!"
                                varData = objTTX.ParseEvInfoV(lEvtCount - 1, 1, 6)
                                dblStartTime = varData(0) + (1 / 100000)
                            End If
                        End If
                    Next
            End Select
            
            'update the totals with the obtained number of events
            For iChanNum = 1 To UBound(arrCount) + 1
                iWriteToChan = iChanNum
                If blnRemapChannels Then
                    iWriteToChan = vChannelMapper.fwdLookup(CLng(iWriteToChan))
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
                    histoSums(iWriteToChan - 1)(lBinNum) = histoSums(iWriteToChan - 1)(lBinNum) + arrCount(iChanNum - 1)
                    histoSquares(iWriteToChan - 1)(lBinNum) = histoSquares(iWriteToChan - 1)(lBinNum) + (arrCount(iChanNum - 1) ^ 2)
                End If
            Next
            ReDim arrCount(intArrCountUpperLimit) 'clear the storage array, but keep it with the same number of channels

            dblStartTime = dblEndTime
            dblEndTime = dblStartTime + dblBinWidthSecs
        Next



End Function
Function calcBinCount(dblTotalWidthSecs As Double, dblBinWidthSecs As Double) As Long
    calcBinCount = CLng(dblTotalWidthSecs / dblBinWidthSecs)
End Function




