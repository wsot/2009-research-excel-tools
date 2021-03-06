Attribute VB_Name = "Main"
Option Explicit

Global Const useSendKeys = False

'USED FOR INTERACTION WITH FORMS
Global doImport
Dim theServer As String
Dim theTank As String
Dim theBlock As String

Dim xAxisEp As String
Dim yAxisEp As String
Dim arrOtherEp As Variant
Dim stimStartEpoc As String

Global bulkImportRootDir As String

Dim bReverseX As Boolean
Dim bReverseY As Boolean
'END SHARED WITH FORMS

Global dblBinWidth As Double
Global dblIgnoreFirstMsec As Double

Dim iRowOffset As Integer
Dim iColOffset As Integer

Global thisWorkbook As Workbook
Global outputWorkbook As Workbook
Global plotWorkbook As Workbook

Global writeTuningsToFile As Boolean
Global CFTextStream As TextStream

Global Const marginForGoodTuning = 1#

Global rngProcTimings As Range
Global rngProcTimingOffset As Integer
Global rngProcStartTime As Double

Dim successfullyProcessedOffset As Integer
Dim vXAxisKeys As Variant
Dim vYAxisKeys As Variant

Function loadConfigParams( _
        ByRef outputWorkbook As Workbook, _
        ByRef thisWorkbook As Workbook, _
        ByRef stimStartEpoc As String, _
        ByRef dblBinWidth As Double, _
        ByRef dblIgnoreFirstMsec As Double, _
        ByRef lNumOfChans As Long, _
        ByRef lFirstChan As Long, _
        ByRef iRowOffset As Integer, _
        ByRef iColOffset As Integer, _
        ByRef arrOtherEp As Variant, _
        ByRef xAxisEp As String, _
        ByRef yAxisEp As String, _
        ByRef bReverseX As Boolean, _
        ByRef bReverseY As Boolean, _
        ByRef oDriveDetParams As Variant, _
        ByRef vChannelMapper As Variant)
        
    loadConfigParams = True
        
    'load the stimulus start epoc
    If Not readCopyParam(outputWorkbook, thisWorkbook, "Variables (do not edit)", "B7", "", stimStartEpoc, vbString, False) Then
        loadConfigParams = False
    End If
    
    'load the bin width for histogram generation
    If Not readCopyParam(outputWorkbook, thisWorkbook, "Settings", "B1", "", dblBinWidth, vbDouble, False) Then
        loadConfigParams = False
    End If
        
    'load the # of msec to ignore at the start (for filtering stimulation artifact
    If Not readCopyParam(outputWorkbook, thisWorkbook, "Settings", "B2", "", dblIgnoreFirstMsec, vbDouble, False) Then
        loadConfigParams = False
    End If
    
    'read number of channels to process; write to output
    If Not readCopyParam(outputWorkbook, thisWorkbook, "Settings", "B3", "", lNumOfChans, vbLong, False) Then
        loadConfigParams = False
    End If
    
    'read number of channels to process; write to output
    If Not readCopyParam(outputWorkbook, thisWorkbook, "Settings", "B4", "", lFirstChan, vbLong, False) Then
        loadConfigParams = False
    End If
    
    'offsets to leave space at the top and left of the chart
    If Not readCopyParam(outputWorkbook, thisWorkbook, "Variables (do not edit)", "E4", "", iRowOffset, vbInteger, False) Then
        loadConfigParams = False
    End If
    If Not readCopyParam(outputWorkbook, thisWorkbook, "Variables (do not edit)", "E5", "", iColOffset, vbInteger, False) Then
        loadConfigParams = False
    End If
            
    Dim i As Integer
    Dim arrOtherEpBuilder() As String
    i = 0
    While thisWorkbook.Worksheets("Variables (do not edit)").Range("B" & (i + 9)).Value <> ""
        ReDim Preserve arrOtherEpBuilder(i)
        arrOtherEpBuilder(i) = thisWorkbook.Worksheets("Variables (do not edit)").Range("B" & (i + 9)).Value
        i = i + 1
    Wend
    arrOtherEp = arrOtherEpBuilder
    
    If Not readCopyParam(outputWorkbook, thisWorkbook, "Variables (do not edit)", "B5", "", xAxisEp, vbString, False) Then
        loadConfigParams = False
    End If
    
    If Not readCopyParam(outputWorkbook, thisWorkbook, "Variables (do not edit)", "B6", "", yAxisEp, vbString, False) Then
        loadConfigParams = False
    End If
    
    If Not readCopyParam(outputWorkbook, thisWorkbook, "Variables (do not edit)", "E1", "", bReverseX, vbBoolean, False) Then
        loadConfigParams = False
    End If
    
    If Not readCopyParam(outputWorkbook, thisWorkbook, "Variables (do not edit)", "E2", "", bReverseY, vbBoolean, False) Then
        loadConfigParams = False
    End If
    
    Set oDriveDetParams = New DriveDetection
    
    If Not oDriveDetParams.readDriveDetection(thisWorkbook.Worksheets("Settings"), "A27", outputWorkbook.Worksheets("Settings")) Then
        loadConfigParams = False
    End If
    
    Set vChannelMapper = New ChannelMapper
    If Not vChannelMapper.readMappingListsFromDirName(theTank & "\" & theBlock, lNumOfChans, outputWorkbook.Worksheets("Channel Mappings").Range("A2"), outputWorkbook.Worksheets("Channel Mappings").Range("B2")) Then
'    If Not vChannelMapper.readMappingLists(thisWorkbook.Worksheets("Channel Mappings").Range("A2"), thisWorkbook.Worksheets("Channel Mappings").Range("B2"), lNumOfChans) Then
'        loadConfigParams = False
    End If
    
End Function

'Tries to detect the CF of each driven channel
Function findCF( _
    objTTX As TTankX, _
    lNumOfChans As Long, _
    dDrivenChanList As Variant, _
    inputWS As Worksheet, _
    varsWS As Worksheet, _
    outputWS As Worksheet, _
    vChannelMapper As Variant _
    ) As Boolean

    If Not IsEmpty(CFTextStream) Then
        If Not CFTextStream Is Nothing Then
            Call CFTextStream.WriteLine("Channel" & Chr(9) & "CF (main)" & Chr(9) & "CF (secondary)")
        End If
    End If

    Dim blnReturnVal As Boolean
    blnReturnVal = True

    Dim xCount As Long
    Dim yCount As Long
    Dim zOffsetSize As Long
    Dim lMaxHistHeight As Long
    Dim iColOffset As Integer
    Dim iRowOffset As Integer
    Dim xPos As Integer
    Dim yPos As Integer

    xCount = varsWS.Range("H1").Value
    yCount = varsWS.Range("H2").Value
    zOffsetSize = varsWS.Range("H3").Value
    iColOffset = varsWS.Range("H5").Value
    iRowOffset = varsWS.Range("H6").Value
    
    Dim lThisKey As Long
    Dim lChanNum As Long

    outputWS.Cells(1, 1).Value = "Channel"
    outputWS.Cells(1, 2).Value = "CF (main)"
    outputWS.Cells(1, 3).Value = "Threshold (main)"
    outputWS.Cells(1, 5).Value = "CF (secondary)"
    outputWS.Cells(1, 5).Value = "Threshold (secondary)"
    For lThisKey = 1 To lNumOfChans
        outputWS.Cells(lThisKey + 1, 1).Value = lThisKey
    Next
    
    Dim i As Long
    Dim j As Long
    
    xPos = iColOffset + 1
    yPos = iRowOffset
    
    Dim lPeakFreq As Long
    Dim lSecondPeakFreq As Long
    Dim bIsPeak As Boolean
    
    Dim dblMean() As Double
    ReDim dblMean(yCount - 1)
    Dim frqVals() As Variant
    ReDim frqVals(xCount - 2) 'Cant't check first freq and last freq for CF because can't get side values
    
    Dim processChannel As Boolean
    
    While inputWS.Cells(yPos, xPos).Value <> ""

        lChanNum = CLng(Right(inputWS.Cells(yPos, xPos).Value, 2))
        If dDrivenChanList.Exists(lChanNum) Then
            processChannel = True
        Else
            processChannel = False
        End If
        
        If processChannel Then
            ReDim dblMean(yCount - 1)
            For i = 0 To (yCount - 1)
                For j = 0 To (xCount - 1)
                    dblMean(i) = dblMean(i) + ((inputWS.Cells(yPos + 2 + i, xPos + j + 1).Value - dblMean(i)) / (j + 1))
                Next
            Next
            
            For j = 1 To (xCount - 2)
                frqVals(j) = 0
                For i = 0 To (yCount - 1)
                    If Not inputWS.Cells(yPos + 2, xPos + j + 1).Value - dblMean(i) < 0 Then
                        frqVals(j) = frqVals(j) + _
                            (inputWS.Cells(yPos + i + 2, xPos + j + 1).Value - dblMean(i)) ^ 2 + _
                            (inputWS.Cells(yPos + i + 2, xPos + j).Value - dblMean(i)) + _
                            (inputWS.Cells(yPos + i + 2, xPos + j + 2).Value - dblMean(i))
                        If i > 0 Then
                            frqVals(j) = frqVals(j) + (inputWS.Cells(yPos + i + 1, xPos + j + 1).Value - dblMean(i - 1))
                        Else
                            frqVals(j) = frqVals(j) + (inputWS.Cells(yPos + i + 2, xPos + j + 1).Value - dblMean(i))
                        End If
                        If i < (yCount - 1) Then
                            frqVals(j) = frqVals(j) + (inputWS.Cells(yPos + i + 3, xPos + j + 1).Value - dblMean(i + 1))
                        Else
                            frqVals(j) = frqVals(j) + (inputWS.Cells(yPos + i + 2, xPos + j + 1).Value - dblMean(i))
                        End If
                    End If
                Next
            Next
                
            For j = 1 To (xCount - 2)
                'check if this is higher than/equal to the two adjacent points - i.e. a peak
                If j = 1 Then
                    If frqVals(j) >= frqVals(j + 1) Then bIsPeak = True
                ElseIf j = (xCount - 2) Then
                    If frqVals(j) >= frqVals(j - 1) Then bIsPeak = True
                ElseIf frqVals(j) >= frqVals(j + 1) And frqVals(j) >= frqVals(j - 1) Then
                    bIsPeak = True
                Else
                    bIsPeak = False
                End If
                
                If bIsPeak Then
                    If frqVals(j) > frqVals(lPeakFreq) Then
                        lPeakFreq = j
                    End If
                End If
            Next
            
            For j = 1 To (xCount - 2)
                If j = 1 Then
                    If frqVals(j) >= frqVals(j + 1) Then bIsPeak = True
                ElseIf j = (xCount - 2) Then
                    If frqVals(j) >= frqVals(j - 1) Then bIsPeak = True
                ElseIf frqVals(j) >= frqVals(j + 1) And frqVals(j) >= frqVals(j - 1) Then
                    bIsPeak = True
                Else
                    bIsPeak = False
                End If
                
                If bIsPeak Then
                    If frqVals(j) > frqVals(lSecondPeakFreq) Then
                        If Abs(j - lPeakFreq) > 2 Then 'check it is not immediately adjacent to (or is) the main peak
                            lSecondPeakFreq = j
                        End If
                    End If
                End If
            Next
            
            If lPeakFreq <> 0 Then
                outputWS.Cells(lChanNum + 1, 2).Value = inputWS.Cells(yPos + 1, xPos + lPeakFreq + 1).Value
                If lSecondPeakFreq <> 0 Then
                    outputWS.Cells(lChanNum + 1, 4).Value = inputWS.Cells(yPos + 1, xPos + lSecondPeakFreq + 1).Value
                End If
            End If
            lPeakFreq = 0
            lSecondPeakFreq = 0
        End If
        yPos = yPos + zOffsetSize
    Wend
    
    If Not IsEmpty(CFTextStream) Then
        If Not CFTextStream Is Nothing Then
            lChanNum = 1
            While outputWS.Cells(lChanNum + 1, 1).Value <> ""
                Call CFTextStream.WriteLine(lChanNum & Chr(9) & outputWS.Cells(lChanNum + 1, 2).Value & Chr(9) & outputWS.Cells(lChanNum + 1, 4).Value)
                lChanNum = lChanNum + 1
            Wend
        End If
    End If
    
End Function

'to check for drive, does a sum of all values with a single value on the Y axis (the first value...)
'it will reflect drive with whatever grouping filter is currently in place when it is called (i.e. it does not reset filters)
'Any channel that does not have drive is fully excluded (including when they are on an X or Y axis)
'DOESN'T ACTUALLY DO ANYWHERE NEAR ALL THAT RIGHT NOW!!
Function checkChannelsForDrive( _
    objTTX As TTankX, _
    xAxisEp As String, _
    vXAxisKeys As Variant, _
    yAxisEp As String, _
    vYAxisKeys As Variant, _
    stimStartEpoc As String, _
    oDriveDetectionParams As DriveDetection, _
    lNumOfChans As Long, _
    lFirstChan As Long, _
    dDrivenChanList As Variant, _
    Optional outputWS As Worksheet, _
    Optional vChannelMapper As Variant) As Boolean

    Const fixAsValidAfterXAdjacentDetections = 3 'once this many sequential detections have turned up the channel it is 'locked' as driven
    Dim blnReturnVal As Boolean
    blnReturnVal = True

    Dim dFinalDrivenChanList As Dictionary
    Set dFinalDrivenChanList = New Dictionary
    Dim dTmpDrivenChanList As Dictionary

    Dim sStableSearchString As String
    Dim sThisSearchString As String
    
'    Dim dDrivenChanList As Dictionary
    
    Dim vStimEpocs As Variant
    Dim aStimTimes() As Double
    
    Dim vStrKeyArray As Variant
    Dim lThisKey As Long
    Dim lStrKeyIndex As Integer
    
    Dim lStimIter As Long
    
    Dim blnOutputToWorksheet As Boolean
    blnOutputToWorksheet = False
    If Not IsMissing(outputWS) Then
        If IsObject(outputWS) Then
            If Not outputWS Is Nothing Then
                outputWS.Cells(1, 1).Value = "Channel"
                outputWS.Cells(1, 3).Value = xAxisEp
                blnOutputToWorksheet = True
                For lThisKey = 1 To lNumOfChans
                    outputWS.Cells(lThisKey + 1, 1).Value = lFirstChan + lThisKey - 1
                Next
            End If
        End If
    End If
    Dim tmpArr As Variant
    
    If Not xAxisEp = "Channel" Then
        sStableSearchString = yAxisEp & " = " & CStr(vYAxisKeys(0))
        Dim i
        For i = 0 To UBound(vXAxisKeys)
            Set dTmpDrivenChanList = Nothing
            sThisSearchString = sStableSearchString & " and " & xAxisEp & " = " & CStr(vXAxisKeys(i))
            Call objTTX.ResetFilters
            Call objTTX.SetFilterWithDescEx(sThisSearchString)
            vStimEpocs = objTTX.GetEpocsExV(stimStartEpoc, 0)
            If Not IsEmpty(vStimEpocs) Then
                Set dTmpDrivenChanList = Nothing
                ReDim aStimTimes(UBound(vStimEpocs, 2))
                For lStimIter = 0 To UBound(vStimEpocs, 2)
                    aStimTimes(lStimIter) = vStimEpocs(1, lStimIter)
                Next

                Call identifyDrivenChannels(objTTX, aStimTimes, oDriveDetectionParams, dTmpDrivenChanList, lNumOfChans, vChannelMapper)
                
                'check if there are previously identified entries not found this round
                vStrKeyArray = dFinalDrivenChanList.Keys
                For lStrKeyIndex = LBound(vStrKeyArray) To UBound(vStrKeyArray)
                    lThisKey = vStrKeyArray(lStrKeyIndex)
                    If Not dTmpDrivenChanList.Exists(lThisKey) Then 'wasn't detected on this pass
                        If dFinalDrivenChanList(lThisKey)(1) < fixAsValidAfterXAdjacentDetections Then 'if didn't detect and isn't already above the 'keeping' threshold, drop the channel
                            'If dFinalDrivenChanList(lThisKey)(1) < 1 Then
'                                Set dFinalDrivenChanList(lThisKey) = Nothing
                                Call dFinalDrivenChanList.Remove(lThisKey)
                            'End If
                           'dFinalDrivenChanList.Remove (lThisKey)
                           'dFinalDrivenChanList(lThisKey) = Nothing
                           'dFinalDrivenChanList.Remove (lThisKey)
                        End If
                    End If
                Next
                
                If blnOutputToWorksheet Then
                    outputWS.Cells(1, 4 + i).Value = CStr(vXAxisKeys(i))
                End If
                
                vStrKeyArray = dTmpDrivenChanList.Keys
                If UBound(vStrKeyArray) > -1 Then 'check the array is not empty
                    For lStrKeyIndex = LBound(vStrKeyArray) To UBound(vStrKeyArray)
                        lThisKey = vStrKeyArray(lStrKeyIndex)
                        If Not dFinalDrivenChanList.Exists(lThisKey) Then
                            Call dFinalDrivenChanList.Add(lThisKey, Array(dTmpDrivenChanList(lThisKey), 1))
                        Else
                            tmpArr = dFinalDrivenChanList(lThisKey)
                            tmpArr(0) = tmpArr(0) And dTmpDrivenChanList(lThisKey)
                            tmpArr(1) = tmpArr(1) + 1
                            dFinalDrivenChanList(lThisKey) = tmpArr
                        End If
                    Next
                    
                    If blnOutputToWorksheet Then
                        vStrKeyArray = dTmpDrivenChanList.Keys
                        For lStrKeyIndex = LBound(vStrKeyArray) To UBound(vStrKeyArray)
                            lThisKey = vStrKeyArray(lStrKeyIndex)
                            outputWS.Cells(lThisKey + 1, 1).Value = lThisKey
                            outputWS.Cells(lThisKey + 1, 4 + i).Value = dTmpDrivenChanList(lThisKey)
                        Next
                    End If
                End If
            Else
                blnReturnVal = False
            End If
        Next
    End If
    
        
    'check if there are channels insufficiently complete to keep
    vStrKeyArray = dFinalDrivenChanList.Keys
    For lStrKeyIndex = LBound(vStrKeyArray) To UBound(vStrKeyArray)
        lThisKey = vStrKeyArray(lStrKeyIndex)
        
            If dFinalDrivenChanList(lThisKey)(1) < fixAsValidAfterXAdjacentDetections Then 'if didn't detect and isn't already above the 'keeping' threshold, drop the channel
                'If dFinalDrivenChanList(lThisKey)(1) < 1 Then
'                                Set dFinalDrivenChanList(lThisKey) = Nothing
                    Call dFinalDrivenChanList.Remove(lThisKey)
                'End If
               'dFinalDrivenChanList.Remove (lThisKey)
               'dFinalDrivenChanList(lThisKey) = Nothing
               'dFinalDrivenChanList.Remove (lThisKey)
            End If
    Next
    
    If blnOutputToWorksheet Then
        vStrKeyArray = dFinalDrivenChanList.Keys
        For lStrKeyIndex = LBound(vStrKeyArray) To UBound(vStrKeyArray)
            lThisKey = vStrKeyArray(lStrKeyIndex)
            outputWS.Cells(lThisKey + 1, 1).Value = lThisKey
            outputWS.Cells(lThisKey + 1, 2).Value = dFinalDrivenChanList(lThisKey)
        Next
    End If
    Set dDrivenChanList = dFinalDrivenChanList
    checkChannelsForDrive = blnReturnVal

End Function

'detects the 'noise floor' for each channel - i.e. the mean spike count per second of the non-acoustic period
Function detectNoiseFloor(objTTX As TTankX, stimStartEpoc As String, oDriveDetectionParams As DriveDetection, lNumOfChans As Long, lFirstChan As Long, dNoiseFloorList As Variant, Optional outputWS As Worksheet, Optional vChannelMapper As Variant) As Boolean

    Dim blnReturnVal As Boolean
    blnReturnVal = True

    Set dNoiseFloorList = New Dictionary
   
    Dim vStimEpocs As Variant
    Dim aStimTimes() As Double
    
    Dim dblMeanSpikes As Double
    Dim dblStdDevSpikes As Double
    
    Dim lStimIter As Long
    Dim iStimNum As Long
    
'    Call objTTX.ResetFilters
    vStimEpocs = objTTX.GetEpocsExV(stimStartEpoc, 0)
    If Not IsEmpty(vStimEpocs) Then
        ReDim aStimTimes(UBound(vStimEpocs, 2))
        For lStimIter = 0 To UBound(vStimEpocs, 2)
            aStimTimes(lStimIter) = vStimEpocs(1, lStimIter)
        Next
            'Call identifyDrivenChannels(objTTX, aStimTimes, oDriveDetectionParams, dTmpDrivenChanList, lNumOfChans)
        
            Dim dblTotalWidthSecs As Double
            Dim dblBinWidthSecs As Double
            Dim dblStartOffsetSecs As Double
            
            Dim histoSums() As Variant
            Dim histoSquares() As Variant
            Dim histoN As Long
            Dim lHistoBinCount As Long
            
            Dim returnVal As Variant
            
            'create bins based of provided configuration parameters to check for an onset spike
            dblBinWidthSecs = oDriveDetectionParams.Diff_ITI - oDriveDetectionParams.Diff_StimDur
            dblTotalWidthSecs = dblBinWidthSecs
            
            lHistoBinCount = calcBinCount(dblTotalWidthSecs, dblBinWidthSecs)
            Call setHistoArraySizes(histoSums, histoSquares, lHistoBinCount, lNumOfChans)
            
            For iStimNum = 0 To UBound(aStimTimes)
                Call buildHistogramForStim(objTTX, aStimTimes(iStimNum) + oDriveDetectionParams.Diff_StimDur, histoSums, histoSquares, dblTotalWidthSecs, dblBinWidthSecs, vChannelMapper)
            Next
            
            Dim dblSpikePerEpoc As Double
            
            Dim lArrIndx As Long
            Dim lDstChan As Long
            Dim lComparisonBin As Long
        
            'step through each channel
            For lArrIndx = 0 To (UBound(histoSums))
                dblMeanSpikes = histoSums(lArrIndx)(0) / (UBound(aStimTimes) + 1)
                dblStdDevSpikes = (histoSquares(lArrIndx)(0) - ((dblMeanSpikes ^ 2) / (UBound(aStimTimes) + 1))) / (UBound(aStimTimes) + 1)
'                dblStdDevSpikes = (histoSquares(lArrIndx)(0) - ((histoSums(lArrIndx)(0) ^ 2) / (UBound(aStimTimes) + 1))) / (UBound(aStimTimes) + 1)
                Call dNoiseFloorList.Add(lArrIndx + 1, Array((dblMeanSpikes + dblStdDevSpikes) / (oDriveDetectionParams.Diff_ITI - oDriveDetectionParams.Diff_StimDur), dblMeanSpikes, dblStdDevSpikes, (oDriveDetectionParams.Diff_ITI - oDriveDetectionParams.Diff_StimDur)))
                'Call dNoiseFloorList.Add(lArrIndx + 1, Array((dblMeanSpikes + dblStdDevSpikes) / (oDriveDetectionParams.Diff_ITI - oDriveDetectionParams.Diff_StimDur) * oDriveDetectionParams.Diff_StimDur, dblMeanSpikes, dblStdDevSpikes, (oDriveDetectionParams.Diff_ITI - oDriveDetectionParams.Diff_StimDur)))
                If Not IsMissing(outputWS) Then
                    If IsObject(outputWS) Then
                        If Not outputWS Is Nothing Then
'                            If Not IsMissing(vChannelMapper) Then
'                                lDstChan = vChannelMapper.fwdLookup(lArrIndx + 1)
'                            Else
                                lDstChan = lArrIndx + 1
'                            End If
                            outputWS.Cells(1, 1).Value = "Channel"
                            outputWS.Cells(1, 2).Value = "Sum"
                            outputWS.Cells(1, 3).Value = "Sum of squares"
                            outputWS.Cells(1, 4).Value = "Mean"
                            outputWS.Cells(1, 5).Value = "StdDev"
                            outputWS.Cells(1, 6).Value = "Threshold"
                            
                                                    
                            outputWS.Cells(lDstChan + 1, 1).Value = lDstChan
                            outputWS.Cells(lDstChan + 1, 2).Value = histoSums(lArrIndx)(0)
                            outputWS.Cells(lDstChan + 1, 3).Value = histoSquares(lArrIndx)(0)
                            outputWS.Cells(lDstChan + 1, 4).Value = dblMeanSpikes
                            outputWS.Cells(lDstChan + 1, 5).Value = dblStdDevSpikes
                            outputWS.Cells(lDstChan + 1, 6).Value = (dblMeanSpikes + dblStdDevSpikes) / (oDriveDetectionParams.Diff_ITI - oDriveDetectionParams.Diff_StimDur)
                        End If
                    End If
                End If

                
            Next
        
    End If
    
    detectNoiseFloor = blnReturnVal

End Function

Sub bulkBuildTuningCurves()

    isFirstChart = True
'        Dim thisWorkbook As Workbook
    Set thisWorkbook = Application.ActiveWorkbook
    Application.Calculation = xlCalculationManual

'    If IsEmpty(theTank) Then
    theServer = thisWorkbook.Worksheets("Variables (do not edit)").Range("B1").Value
'        theTank = thisWorkbook.Worksheets("Variables (do not edit)").Range("B2").Value
'        theBlock = thisWorkbook.Worksheets("Variables (do not edit)").Range("B3").Value
'    End If

    bulkImportRootDir = thisWorkbook.Worksheets("Settings").Range("B21").Value
    If bulkImportRootDir = "" Then
        MsgBox "If bulk importing, a root data folder must be specified"
        Exit Sub
    ElseIf Not checkPathExists(bulkImportRootDir) Then
        MsgBox "The bulk import path does not exist: " & bulkImportRootDir
        Exit Sub
    End If

    BulkImportFrom.Show
    
    If doImport Then
        
        If thisWorkbook.Worksheets("Variables (do not edit)").Range("E9").Value = "" Then
            successfullyProcessedOffset = 49
        Else
            successfullyProcessedOffset = thisWorkbook.Worksheets("Variables (do not edit)").Range("E9").Value
        End If
        
        Dim objFS As FileSystemObject
        Set objFS = New FileSystemObject
        
        Dim specifiedOutputDir As String
        Dim outputDir As String
        Dim outputFilename As String
        specifiedOutputDir = thisWorkbook.Worksheets("Settings").Range("B12").Value
        'outputDir = getDirName(specifiedOutputDir, theTank)
                        
        Dim templatePath As String
        templatePath = thisWorkbook.Worksheets("Settings").Range("B16").Value
        
        Dim outputFilePrefix As String
        outputFilePrefix = thisWorkbook.Worksheets("Settings").Range("B11").Value
        
        Dim blnAutoclose As Boolean
        blnAutoclose = thisWorkbook.Worksheets("Settings").Range("B10").Value
        
        Dim writeTuningsToFile As Boolean
        writeTuningsToFile = thisWorkbook.Worksheets("Settings").Range("B47").Value
        
        Dim blnAutosave As Boolean
        If blnAutoclose Then
            blnAutosave = True
        Else
            blnAutosave = thisWorkbook.Worksheets("Settings").Range("B9").Value
        End If
        
        Dim blnAutoPlot As Boolean
        blnAutoPlot = thisWorkbook.Worksheets("Settings").Range("B5").Value
       
        Dim blnPlotOnlyCandidates As Boolean
        blnPlotOnlyCandidates = thisWorkbook.Worksheets("Settings").Range("B6").Value
       
        Dim blnPlotOnlyDriven As Boolean
        Dim dDrivenChans As Dictionary
        blnPlotOnlyDriven = thisWorkbook.Worksheets("Settings").Range("B24").Value
                
        Dim blnSubtractNoiseFloor As Boolean
        Dim dNoiseFloorList As Dictionary
        blnSubtractNoiseFloor = thisWorkbook.Worksheets("Settings").Range("B25").Value
        thisWorkbook.Worksheets("Variables (do not edit)").Range("E7").Value = thisWorkbook.Worksheets("Settings").Range("B18").Value
       
        Dim vChannelMapper As Variant
       
        Dim dBlocks As Dictionary
        Set dBlocks = New Dictionary
        Dim i As Integer
        i = 2
        
        While thisWorkbook.Worksheets("Variables (do not edit)").Range("N" & i).Value <> ""
            If Not dBlocks.Exists(thisWorkbook.Worksheets("Variables (do not edit)").Range("N" & i).Value) Then
                Call dBlocks.Add(thisWorkbook.Worksheets("Variables (do not edit)").Range("N" & i).Value, 0)
            End If
            i = i + 1
        Wend
    
        Dim theBlocks As Variant
        theBlocks = dBlocks.Keys
        
        Application.DisplayAlerts = False
        
'        Dim outputWorkbook As Workbook
        
        Dim vArrExcl As Variant
        Dim strErr As String
        
        For i = LBound(theBlocks) To UBound(theBlocks)
            'Call Worksheets("Totals").UsedRange.ClearContents
            'Call Worksheets("StdDev").UsedRange.ClearContents
            'Call Worksheets("Means").UsedRange.ClearContents
            'Call Worksheets("N").UsedRange.ClearContents
            theTank = Left(theBlocks(i), InStr(theBlocks(i), ":") - 1)
            theBlock = Right(theBlocks(i), Len(theBlocks(i)) - Len(theTank) - 1)
            theTank = bulkImportRootDir & "\" & theTank
            
            vArrExcl = checkForMapExclusion(objFS.GetFolder(theTank & "\" & theBlock))
            If vArrExcl(0) = "generate" Then
                thisWorkbook.Worksheets("Settings").Cells(successfullyProcessedOffset, 1).Value = theBlocks(i) & " excluded"
                thisWorkbook.Worksheets("Settings").Cells(successfullyProcessedOffset, 2).Value = theBlocks(i) & " Excluded: " & vArrExcl(1)
            Else
                thisWorkbook.Worksheets("Settings").Cells(successfullyProcessedOffset, 1).Value = theBlocks(i) & " processing"
            
                If i = 0 Then
                    templatePath = getFilename(templatePath, theTank)
                End If
                Set outputWorkbook = Workbooks.Open(templatePath)
                
                Set rngProcTimings = outputWorkbook.Worksheets("Settings").Range("I1")
                rngProcTimingOffset = 1
                
                If specifiedOutputDir = "" Then
                    outputDir = getDirName("", theTank)
                    outputFilename = outputDir & "\" & outputFilePrefix & theBlock
                Else
                    outputDir = getDirName(specifiedOutputDir, theTank)
                    If outputDir = "" Then
                        MsgBox ("Output directory " & outputDir & " could not be found." & vbCrLf & "Please update the path and try again")
                        Exit Sub
                    End If
                    outputFilename = outputDir & "\" & Right(Replace(theTank, "\", "."), Len(theTank) - Len(bulkImportRootDir) - 1) & "_" & outputFilePrefix & theBlock
                End If
                
                outputWorkbook.Worksheets("Variables (do not edit)").Range("B2").Value = theTank 'update the block on the worksheet
                outputWorkbook.Worksheets("Variables (do not edit)").Range("B3").Value = theBlock 'update the block on the worksheet
                'outputWorkbook.Worksheets("Settings").Range("B18").Value = thisWorkbook.Worksheets("Settings").Range("B18").Value
                outputWorkbook.Worksheets("Variables (do not edit)").Range("E7").Value = thisWorkbook.Worksheets("Variables (do not edit)").Range("E7").Value
                
                outputWorkbook.Worksheets("Settings").Range("B6").Value = blnPlotOnlyCandidates
                outputWorkbook.Worksheets("Settings").Range("B24").Value = blnPlotOnlyDriven
                outputWorkbook.Worksheets("Settings").Range("B25").Value = blnSubtractNoiseFloor
                
                If writeTuningsToFile And vArrExcl(0) <> "cf" Then 'check if the CFs should be written to text files
                    Dim strTxtFileName As String
                    strTxtFileName = objFS.GetFolder(theTank).ParentFolder.ParentFolder.Path & "\Map-" & Right(Replace(theTank, "\", "."), Len(theTank) - Len(bulkImportRootDir) - 1) & "_" & outputFilePrefix & theBlock & ".txt"
                    Set CFTextStream = objFS.CreateTextFile(strTxtFileName, True, False)
                    CFTextStream.WriteLine ("Generated: " & Chr(9) & Now())
                    CFTextStream.WriteLine ("Tank: " & Chr(9) & theTank)
                    CFTextStream.WriteLine ("Block: " & Chr(9) & theBlock)
                Else
                    Set CFTextStream = Nothing
                End If
                
                strErr = ""
                strErr = processImport(False, blnPlotOnlyDriven, dDrivenChans, blnSubtractNoiseFloor, dNoiseFloorList, vChannelMapper)
                
                If strErr = "" Then
                    'Call detectTunedSegments
                    Application.Calculation = xlCalculationAutomatic
                    Application.Calculation = xlCalculationManual
                    If blnAutosave Then
                        Call outputWorkbook.SaveAs(outputFilename, 52)
                        If blnAutoPlot Then
                            If Not useSendKeys Then
                                If IsNull(SigmaPlotHandle) Or SigmaPlotHandle = 0 Then
                                    Call findSigmplotWindow
                                End If
                            End If
                            Set plotWorkbook = outputWorkbook
                            If blnPlotOnlyCandidates Then
                                Call transferCandidatesToSigmaplot(dDrivenChans, outputFilename & ".JNB", vChannelMapper)
                            Else
                                Call transferAllToSigmaplot(outputFilename & ".JNB")
                            End If
                        End If
                        If blnAutoclose Then
                            Call outputWorkbook.Close
                        End If
                    End If
                    thisWorkbook.Worksheets("Settings").Cells(successfullyProcessedOffset, 2).Value = "Processed"
                Else
                    thisWorkbook.Worksheets("Settings").Cells(successfullyProcessedOffset, 2).Value = "Problem with import: " & strErr
                End If
            End If
            If Not IsEmpty(CFTextStream) Then 'close the CF listing file
                If Not CFTextStream Is Nothing Then
                    Call CFTextStream.Close
                    Set CFTextStream = Nothing
                End If
            End If
            successfullyProcessedOffset = successfullyProcessedOffset + 1
            thisWorkbook.Worksheets("Variables (do not edit)").Range("D10").Value = successfullyProcessedOffset
        Next
        Set objFS = Nothing
        Application.DisplayAlerts = True
    End If
End Sub

Sub buildTuningCurves()
    MsgBox "Currently no workee"
    Exit Sub

    isFirstChart = True
'        Dim thisWorkbook As Workbook
    Set thisWorkbook = Application.ActiveWorkbook

    If IsEmpty(theTank) Then
        theServer = thisWorkbook.Worksheets("Variables (do not edit)").Range("B1").Value
        theTank = thisWorkbook.Worksheets("Variables (do not edit)").Range("B2").Value
        theBlock = thisWorkbook.Worksheets("Variables (do not edit)").Range("B3").Value
    End If

    ImportFrom.Show
    
    If doImport Then
        Dim outputDir As String
        outputDir = thisWorkbook.Worksheets("Settings").Range("B12").Value
        outputDir = getDirName(outputDir, theTank)
               
        If outputDir = "" Then
            MsgBox ("Output directory " & outputDir & " could not be found." & vbCrLf & "Please update the path and try again")
            Exit Sub
        End If
                
        Dim templatePath As String
        templatePath = thisWorkbook.Worksheets("Settings").Range("B16").Value
        templatePath = getFilename(templatePath, theTank)
        
        Dim outputFilePrefix As String
        outputFilePrefix = thisWorkbook.Worksheets("Settings").Range("B11").Value
        
        Dim blnAutoclose As Boolean
        blnAutoclose = thisWorkbook.Worksheets("Settings").Range("B10").Value
        
        Dim blnAutosave As Boolean
        If blnAutoclose Then
            blnAutosave = True
        Else
            blnAutosave = thisWorkbook.Worksheets("Settings").Range("B9").Value
        End If
        
        Dim blnAutoPlot As Boolean
        blnAutoPlot = thisWorkbook.Worksheets("Settings").Range("B5").Value

        
        Dim dBlocks As Dictionary
        Set dBlocks = New Dictionary
        Dim i As Integer
        i = 2
        
        While thisWorkbook.Worksheets("Variables (do not edit)").Range("N" & i).Value <> ""
            If Not dBlocks.Exists(thisWorkbook.Worksheets("Variables (do not edit)").Range("N" & i).Value) Then
                Call dBlocks.Add(thisWorkbook.Worksheets("Variables (do not edit)").Range("N" & i).Value, 0)
            End If
            i = i + 1
        Wend
    
        Dim theBlocks As Variant
        theBlocks = dBlocks.Keys
        
        Application.DisplayAlerts = False
        
'        Dim outputWorkbook As Workbook
        
        For i = LBound(theBlocks) To UBound(theBlocks)
            Set outputWorkbook = Workbooks.Open(templatePath)
            'Call Worksheets("Totals").UsedRange.ClearContents
            'Call Worksheets("StdDev").UsedRange.ClearContents
            'Call Worksheets("Means").UsedRange.ClearContents
            'Call Worksheets("N").UsedRange.ClearContents
            theBlock = theBlocks(i)
            outputWorkbook.Worksheets("Variables (do not edit)").Range("B3").Value = theBlock 'update the block on the worksheet
            outputWorkbook.Worksheets("Settings").Range("B18").Value = thisWorkbook.Worksheets("Settings").Range("B18").Value
            Call processImport(False)
            If blnAutosave Then
                Call outputWorkbook.SaveAs(outputDir & "\" & outputFilePrefix & theBlock, 52)
                If blnAutoPlot Then
                    Set plotWorkbook = outputWorkbook
                    If blnPlotOnlyCandidates Then
                        Call transferCandidatesToSigmaplot(vDrivenChannels, outputFilename & ".JNB")
                    Else
                        Call transferAllToSigmaplot(outputFilename & ".JNB")
                    End If
                End If

                If blnAutoclose Then
                    Call outputWorkbook.Close
                End If
            End If
        Next
        Application.DisplayAlerts = True
    End If
End Sub


Function processImport(importIntoSigmaplot As Boolean, Optional vDetectDriven As Variant, Optional vDrivenChans As Variant, Optional vSubtractNoiseFloor As Variant, Optional vNoiseFloorList As Variant, Optional vChannelMapper As Variant) As String
'    processImport = True
    Dim strTotalErrString As String
    strTotalErrString = ""
    Dim lNumOfChans As Long
    Dim lFirstChan As Long
    Dim oDriveDetectionParams As DriveDetection

    If loadConfigParams(outputWorkbook, thisWorkbook, stimStartEpoc, dblBinWidth, dblIgnoreFirstMsec, lNumOfChans, lFirstChan, iRowOffset, iColOffset, arrOtherEp, xAxisEp, yAxisEp, bReverseX, bReverseY, oDriveDetectionParams, vChannelMapper) Then
        
        'used to store the maximum histogram peak for normalisation
        Dim lMaxHistHeight As Double
        lMaxHistHeight = 0
        Dim lMaxHistMeanHeight As Double
        lMaxHistMeanHeight = 0
        
        Dim arrHistTmp() As Long 'used to store the histogram data for each channel as it is generated
        ReDim arrHistTmp(lNumOfChans - 1)
        
        Dim yCount As Long 'number of items on y axis per block
        Dim xCount As Long 'number of items on x axis per block
        Dim zOffsetSize As Long 'the total length that needs to be offset per set of grouping parameters
    
    '    theWorksheets = buildWorksheetArray() 'build the worksheets for writing data
        
        'connect to the tank
        Dim objTTX As TTankX
        Set objTTX = New TTankX
        Dim strErr As String
        strErr = connectToTDTReportError(connectToTDT(objTTX, False, theServer, theTank, theBlock))
        If strErr = "" Then 'if error occurred, then not blank
            'index epochs - required to use filters
            Call objTTX.CreateEpocIndexing
            
            Dim dblStartTime As Double
            Dim dblEndTime As Double
            
            Dim varReturn As Variant
            
        '    Dim vXAxisKeys As Variant
        '    Dim vYAxisKeys As Variant
            
            rngProcStartTime = Now()
            rngProcTimings.Offset(rngProcTimingOffset, 0) = "Build Epoc List"
            vXAxisKeys = BuildEpocList(objTTX, xAxisEp, bReverseX, lNumOfChans, lFirstChan)
            vYAxisKeys = BuildEpocList(objTTX, yAxisEp, bReverseY, lNumOfChans, lFirstChan)
            rngProcTimings.Offset(rngProcTimingOffset, 1) = (Now() - rngProcStartTime) * 3600 * 24
            rngProcTimingOffset = rngProcTimingOffset + 1
            
            
            If Not IsMissing(vDetectDriven) Or IsMissing(vDrivenChans) Then
                If VarType(vDetectDriven) = vbBoolean Then
                    If vDetectDriven = True Then
                        rngProcStartTime = Now()
                        rngProcTimings.Offset(rngProcTimingOffset, 0) = "Check channels for drive"
                        If Not checkChannelsForDrive(objTTX, xAxisEp, vXAxisKeys, yAxisEp, vYAxisKeys, stimStartEpoc, oDriveDetectionParams, lNumOfChans, lFirstChan, vDrivenChans, outputWorkbook.Worksheets("Drive detection output"), vChannelMapper) Then 'WARNING! this doesn't correctly support channel mappings etc yet!!
                            'processImport = False
                            strTotalErrString = strTotalErrString & "Check channels for drive failed, "
                        End If
                        rngProcTimings.Offset(rngProcTimingOffset, 1) = (Now() - rngProcStartTime) * 3600 * 24
                        rngProcTimingOffset = rngProcTimingOffset + 1
                    End If
                End If
            End If
            
            If Not IsMissing(vSubtractNoiseFloor) Or IsMissing(vNoiseFloorList) Then
                If VarType(vSubtractNoiseFloor) = vbBoolean Then
                    If vSubtractNoiseFloor = True Then
                        rngProcStartTime = Now()
                        rngProcTimings.Offset(rngProcTimingOffset, 0) = "Detect Noise Floor"
                        
                        Dim lArrIndex As Long
                        Dim lMinAmp As Long
                        For lArrIndex = 0 To UBound(vYAxisKeys)
                            If lArrIndex = 0 Then
                                lMinAmp = vYAxisKeys(lArrIndex)
                            ElseIf vYAxisKeys(lArrIndex) < lMinAmp Then
                                lMinAmp = vYAxisKeys(lArrIndex)
                            End If
                        Next
                                                
                        Call objTTX.ResetFilters
                        Call objTTX.SetFilterWithDescEx(yAxisEp & " = " & lMinAmp)
            
                        If Not detectNoiseFloor(objTTX, stimStartEpoc, oDriveDetectionParams, lNumOfChans, lFirstChan, vNoiseFloorList, outputWorkbook.Worksheets("Noise Floor"), vChannelMapper) Then  'WARNING! this doesn't correctly support channel mappings etc yet!!
                            strTotalErrString = strTotalErrString & "Detect noise floor failed, "
                            'processImport = False
                        End If
                        Call objTTX.ResetFilters
                        rngProcTimings.Offset(rngProcTimingOffset, 1) = (Now() - rngProcStartTime) * 3600 * 24
                        rngProcTimingOffset = rngProcTimingOffset + 1
                    End If
                End If
            End If
            
            Dim i As Long
            Dim j As Long
            Dim k As Long
            Dim l As Long
            
            Dim arrOtherEpocKeys() As Variant
            If UBound(arrOtherEp) <> -1 Then
                ReDim arrOtherEpocKeys(UBound(arrOtherEp))
                
                For i = 0 To UBound(arrOtherEp)
                    arrOtherEpocKeys(i) = BuildEpocList(objTTX, arrOtherEp(i), False, lNumOfChans, lFirstChan)
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
        
            rngProcStartTime = Now()
            rngProcTimings.Offset(rngProcTimingOffset, 0) = "Process search (the big job)"
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
                        Call processSearch(objTTX, arrOtherEp, arrOtherEpocKeys, 0, xAxisSearchString & yAxisSearchString, i + 1, j + 1, UBound(vYAxisKeys) + 3, iChanNum, "", xCount, yCount, zOffsetSize, lMaxHistHeight, lMaxHistMeanHeight, vNoiseFloorList, vDrivenChans, vChannelMapper)
                    Next
                Next
            End If
            rngProcTimings.Offset(rngProcTimingOffset, 1) = (Now() - rngProcStartTime) * 3600 * 24
            rngProcTimingOffset = rngProcTimingOffset + 1

        
        '    Call writeAxes(theWorksheets, vXAxisKeys, vYAxisKeys, iColOffset, iRowOffset)
        
            outputWorkbook.Worksheets("Variables (do not edit)").Range("H1").Value = xCount
            outputWorkbook.Worksheets("Variables (do not edit)").Range("H2").Value = yCount
            outputWorkbook.Worksheets("Variables (do not edit)").Range("H3").Value = zOffsetSize
            outputWorkbook.Worksheets("Variables (do not edit)").Range("H4").Value = lMaxHistHeight
            outputWorkbook.Worksheets("Variables (do not edit)").Range("H5").Value = iColOffset
            outputWorkbook.Worksheets("Variables (do not edit)").Range("H6").Value = iRowOffset
            outputWorkbook.Worksheets("Variables (do not edit)").Range("H7").Value = lMaxHistMeanHeight
            
            rngProcStartTime = Now()
            rngProcTimings.Offset(rngProcTimingOffset, 0) = "Find CF"
            'Call findCF(objTTX, lNumOfChans, vDrivenChans, outputWorkbook.Worksheets("Means"), outputWorkbook.Worksheets("Variables (do not edit)"), outputWorkbook.Worksheets("Channel Tuning"), vChannelMapper)
            rngProcTimings.Offset(rngProcTimingOffset, 1) = (Now() - rngProcStartTime) * 3600 * 24
            rngProcTimingOffset = rngProcTimingOffset + 1
            
            rngProcStartTime = Now()
            rngProcTimings.Offset(rngProcTimingOffset, 0) = "Generate histograms"
            'Call generateChanHistograms(objTTX, outputWorkbook.Worksheets("Channel Tuning"), outputWorkbook.Worksheets("Histograms"), xAxisEp, yAxisEp, vXAxisKeys, vYAxisKeys, lNumOfChans, stimStartEpoc, vChannelMapper)
            rngProcTimings.Offset(rngProcTimingOffset, 1) = (Now() - rngProcStartTime) * 3600 * 24
            rngProcTimingOffset = rngProcTimingOffset + 1
        
            Call objTTX.CloseTank
            Call objTTX.ReleaseServer
            
            'If importIntoSigmaplot Then
            '    Call TransferToSigmaplot(xCount, yCount, zOffsetSize, iColOffset, iRowOffset, lMaxHistHeight)
            'End If
        Else
            strTotalErrString = strTotalErrString & "TDT connection error: " & strErr & ", "
            'processImport = False
        End If
    Else
        'processImport = False
        strTotalErrString = strTotalErrString & "Loading settings failed, "
    End If
    processImport = strTotalErrString
End Function

Sub writeAxes(colLabels As Variant, rowLabels As Variant, iColOffset, iRowOffset, zOffset)
    Dim j As Long
        
    For j = 0 To UBound(rowLabels)
        outputWorkbook.Worksheets("Totals").Cells(iRowOffset + j + 2 + zOffset, iColOffset + 1).Value = rowLabels(j)
        outputWorkbook.Worksheets("Totals Noise Floor").Cells(iRowOffset + j + 2 + zOffset, iColOffset + 1).Value = rowLabels(j)
        outputWorkbook.Worksheets("StdDev").Cells(iRowOffset + j + 2 + zOffset, iColOffset + 1).Value = rowLabels(j)
        outputWorkbook.Worksheets("Means").Cells(iRowOffset + j + 2 + zOffset, iColOffset + 1).Value = rowLabels(j)
        outputWorkbook.Worksheets("Means Noise Floor").Cells(iRowOffset + j + 2 + zOffset, iColOffset + 1).Value = rowLabels(j)
        outputWorkbook.Worksheets("N").Cells(iRowOffset + j + 2 + zOffset, iColOffset + 1).Value = rowLabels(j)
        outputWorkbook.Worksheets("Noise-adjusted Totals").Cells(iRowOffset + j + 2 + zOffset, iColOffset + 1).Value = rowLabels(j)
        outputWorkbook.Worksheets("Noise-adjusted Means").Cells(iRowOffset + j + 2 + zOffset, iColOffset + 1).Value = rowLabels(j)
    Next
    For j = 0 To UBound(colLabels)
        outputWorkbook.Worksheets("Totals").Cells(iRowOffset + zOffset + 1, j + 2).Value = colLabels(j)
        outputWorkbook.Worksheets("Totals Noise Floor").Cells(iRowOffset + zOffset + 1, j + 2).Value = colLabels(j)
        outputWorkbook.Worksheets("StdDev").Cells(iRowOffset + zOffset + 1, j + 2).Value = colLabels(j)
        outputWorkbook.Worksheets("Means").Cells(iRowOffset + zOffset + 1, j + 2).Value = colLabels(j)
        outputWorkbook.Worksheets("Means Noise Floor").Cells(iRowOffset + zOffset + 1, j + 2).Value = colLabels(j)
        outputWorkbook.Worksheets("N").Cells(iRowOffset + zOffset + 1, j + 2).Value = colLabels(j)
        outputWorkbook.Worksheets("Noise-adjusted Totals").Cells(iRowOffset + zOffset + 1, j + 2).Value = colLabels(j)
        outputWorkbook.Worksheets("Noise-adjusted Means").Cells(iRowOffset + zOffset + 1, j + 2).Value = colLabels(j)
    Next

End Sub


Function BuildEpocList(objTTX, AxisEp, bReverseOrder, lNumOfChan, lFirstChan)
    'build list of epocs for the given axis epoc name
    
    Dim AxisList As Dictionary
    Set AxisList = New Dictionary
    
    Dim dblStartTime As Double
    Dim varReturn As Variant
    
    Dim i As Integer
    Dim j As Integer
    
    If AxisEp = "Channel" Then
        For i = 1 To lNumOfChan
            Call AxisList.Add(lFirstChan + i - 1, 0)
        Next
    Else
        Do
            i = objTTX.ReadEventsV(10000, AxisEp, 0, 0, dblStartTime, 0#, "ALL")
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
        BuildEpocList = returnArr
    Else
        BuildEpocList = AxisList.Keys
    End If

End Function


Function processSearch(ByRef objTTX, ByRef arrOtherEp, ByRef arrOtherEpocKeys, iOtherEpocNum, strSearchString As String, xOffset, yOffset, zOffset, iChanNum, strTitle, ByRef xCount, ByRef yCount, ByRef zOffsetSize, ByRef lMaxHistHeight, ByRef lMaxHistMeanHeight, Optional ByRef vNoiseFloorList As Variant, Optional ByRef vDrivenChans As Variant, Optional ByRef vChannelMapper As Variant)
    Dim i As Integer
    Dim j As Integer
    Dim strAddedSearchString As String
    Dim strFilter As String
    Dim strAddedTitle As String
    Dim gotSomeResults As Boolean
    
    Dim iRowsWithData As Long
    iRowsWithData = 0
    
    'Stop
    
    Dim searchStringList As Dictionary
    Set searchStringList = New Dictionary
    Call buildSearchStringList(arrOtherEp, arrOtherEpocKeys, searchStringList, 0, strSearchString)
    
    'Stop
    
    Call testSearchStringList(objTTX, searchStringList)
    
    Dim searchStrings As Variant
    searchStrings = searchStringList.Keys
    
    For i = 0 To UBound(searchStrings)
        Call objTTX.SetFilterWithDescEx(searchStrings(i))
        
        If Not IsMissing(vNoiseFloorList) Then
            If Not IsMissing(vDrivenChans) Then
                Call writeResults(objTTX, xOffset, yOffset, i * zOffset, iChanNum, lMaxHistHeight, lMaxHistMeanHeight, vNoiseFloorList, vDrivenChans, vChannelMapper, gotSomeResults)
            Else
                Call writeResults(objTTX, xOffset, yOffset, i * zOffset, iChanNum, lMaxHistHeight, lMaxHistMeanHeight, vNoiseFloorList, , vChannelMapper, gotSomeResults)
            End If
        Else
            If Not IsMissing(vDrivenChans) Then
                Call writeResults(objTTX, xOffset, yOffset, i * zOffset, iChanNum, lMaxHistHeight, lMaxHistMeanHeight, , vDrivenChans, vChannelMapper, gotSomeResults)
            Else
                Call writeResults(objTTX, xOffset, yOffset, i * zOffset, iChanNum, lMaxHistHeight, lMaxHistMeanHeight, , , vChannelMapper, gotSomeResults)
            End If
        End If
        
        If gotSomeResults Then
            If xOffset = 1 And yOffset = 1 Then
                outputWorkbook.Worksheets("Totals").Cells(iRowOffset + (i * zOffset), iColOffset + 1).Value = searchStrings(i)
                outputWorkbook.Worksheets("Totals Noise Floor").Cells(iRowOffset + (i * zOffset), iColOffset + 1).Value = searchStrings(i)
                outputWorkbook.Worksheets("N").Cells(iRowOffset + (i * zOffset), iColOffset + 1).Value = searchStrings(i)
                outputWorkbook.Worksheets("Means").Cells(iRowOffset + (i * zOffset), iColOffset + 1).Value = searchStrings(i)
                outputWorkbook.Worksheets("Means Noise Floor").Cells(iRowOffset + (i * zOffset), iColOffset + 1).Value = searchStrings(i)
                outputWorkbook.Worksheets("StdDev").Cells(iRowOffset + (i * zOffset), iColOffset + 1).Value = searchStrings(i)
                outputWorkbook.Worksheets("Noise-adjusted Totals").Cells(iRowOffset + (i * zOffset), iColOffset + 1).Value = searchStrings(i)
                outputWorkbook.Worksheets("Noise-adjusted Means").Cells(iRowOffset + (i * zOffset), iColOffset + 1).Value = searchStrings(i)
                Call writeAxes(vXAxisKeys, vYAxisKeys, iColOffset, iRowOffset, (i * zOffset))
            End If
        End If
        
        If xOffset > xCount Then
            xCount = xOffset
        End If
        If yOffset > yCount Then
            yCount = yOffset
        End If
    Next

End Function

Sub writeResults(ByRef objTTX, xOffset, yOffset, zOffset, iChanNum, ByRef lMaxHistHeight, ByRef lMaxHistMeanHeight, Optional ByRef vNoiseFloorList As Variant, Optional ByRef vDrivenChans As Variant, Optional ByRef vChannelMapper As Variant, Optional ByRef gotSomeResults As Boolean)
    Dim strTmpAddr As String
    Dim strTmpFormula As String
    Dim varReturn As Variant
    Dim varChanData As Variant
    
    Dim lMappedChan As Long
    If Not IsMissing(vChannelMapper) Then
        lMappedChan = vChannelMapper.fwdLookup(CLng(iChanNum))
    Else
        lMappedChan = CLng(iChanNum)
    End If
    
    Dim dblStartTime As Double
    Dim dblEndTime As Double
    Dim dblSwepStartTime As Double
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Dim histTmp As Long
    Dim histVariance As Double
    Dim histStddev As Double
    Dim histMean As Double
    Dim nSweps As Long
    nSweps = 0

    Dim swepVals()

    varReturn = objTTX.GetEpocsExV(stimStartEpoc, 0)
    'Stop
    If IsArray(varReturn) Then
        gotSomeResults = True
        ReDim swepVals(UBound(varReturn, 2))
        nSweps = UBound(varReturn, 2) + 1
        For i = 0 To UBound(varReturn, 2)
            dblStartTime = varReturn(2, i) + dblIgnoreFirstMsec
            dblEndTime = dblStartTime + dblBinWidth + dblIgnoreFirstMsec
            dblSwepStartTime = dblStartTime
            Do
                k = objTTX.ReadEventsV(10000, "CSPK", iChanNum, 0, dblStartTime, dblEndTime, "JUSTTIMES")
                If k = 0 Then
                    Exit Do
                End If
    
                histTmp = CLng(histTmp) + CLng(k)
                swepVals(i) = CLng(swepVals(i)) + CLng(k)
                If k < 10000 Then
                    Exit Do
                Else
                    varChanData = objTTX.ParseEvInfoV(k - 1, 1, 6)
                    dblStartTime = varChanData(0) + (1 / 100000)
                End If
                
            Loop
            dblStartTime = dblSwepStartTime
        Next
        
        histMean = CDbl(histTmp) / CDbl((UBound(swepVals) + 1))
        histVariance = 0#
        
        For i = 0 To UBound(swepVals)
            histVariance = histVariance + (histMean - CDbl(swepVals(i))) ^ 2
        Next
        histStddev = (histVariance / UBound(swepVals)) ^ 0.5
                
        If xAxisEp = "Channel" Then
            outputWorkbook.Worksheets("Totals").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Value = histTmp
            outputWorkbook.Worksheets("Means").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Value = histMean
            outputWorkbook.Worksheets("StdDev").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Value = histStddev
            outputWorkbook.Worksheets("N").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Value = nSweps
        ElseIf yAxisEp = "Channel" Then
            outputWorkbook.Worksheets("Totals").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Value = histTmp
            outputWorkbook.Worksheets("Means").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Value = histMean
            outputWorkbook.Worksheets("StdDev").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Value = histStddev
            outputWorkbook.Worksheets("N").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Value = nSweps
        Else
            outputWorkbook.Worksheets("Totals").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Value = histTmp
            outputWorkbook.Worksheets("Means").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Value = histMean
            outputWorkbook.Worksheets("StdDev").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Value = histStddev
            outputWorkbook.Worksheets("N").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Value = nSweps
            If iChanNum <> 0 Then
                If Not IsMissing(vNoiseFloorList) Then
                    If Not vNoiseFloorList Is Nothing Then
                        outputWorkbook.Worksheets("Totals Noise Floor").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Value = vNoiseFloorList(lMappedChan)(0) * dblBinWidth * nSweps
                        outputWorkbook.Worksheets("Means Noise Floor").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Value = vNoiseFloorList(lMappedChan)(0) * dblBinWidth
                        
                        'outputWorkbook.Worksheets("Totals Noise Floor").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Value = (vNoiseFloorList(iChanNum)(1) / vNoiseFloorList(iChanNum)(3)) * dblBinWidth * nSweps * 2
                        'outputWorkbook.Worksheets("Means Noise Floor").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Value = (vNoiseFloorList(iChanNum)(1) / vNoiseFloorList(iChanNum)(3)) * dblBinWidth
                    
                        strTmpAddr = outputWorkbook.Worksheets("Totals").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Address
                        strTmpFormula = "=IF('Totals'!" & strTmpAddr & "-(('Noise Floor'!F" & (lMappedChan + 1) & " * 'Settings'!B1)*('N'!" & strTmpAddr & ")) < 0,0,'Totals'!" & strTmpAddr & "-(('Noise Floor'!F" & (lMappedChan + 1) & " * 'Settings'!B1)*('N'!" & strTmpAddr & ")))"
                        outputWorkbook.Worksheets("Noise-adjusted Totals").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Formula = strTmpFormula
                        strTmpFormula = "=IF('Means'!" & strTmpAddr & "-('Noise Floor'!F" & (lMappedChan + 1) & " * 'Settings'!B1) < 0,0,'Means'!" & strTmpAddr & "-('Noise Floor'!F" & (lMappedChan + 1) & " * 'Settings'!B1))"
                        outputWorkbook.Worksheets("Noise-adjusted Means").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Formula = strTmpFormula
                    End If
                End If
            End If
        End If
        If Not IsMissing(vDrivenChans) Then
            If Not vDrivenChans Is Nothing Then
                If vDrivenChans.Exists(lMappedChan) Then
                    If histMean > lMaxHistMeanHeight Then
                        lMaxHistMeanHeight = histMean
                    End If
                    If histTmp > lMaxHistHeight Then
                        lMaxHistHeight = histTmp
                    End If
                End If
            Else
                If histMean > lMaxHistMeanHeight Then
                    lMaxHistMeanHeight = histMean
                End If
                If histTmp > lMaxHistHeight Then
                    lMaxHistHeight = histTmp
                End If
            End If
        Else
            If histMean > lMaxHistMeanHeight Then
                lMaxHistMeanHeight = histMean
            End If
            If histTmp > lMaxHistHeight Then
                lMaxHistHeight = histTmp
            End If
        End If
    Else
        gotSomeResults = False
    End If
    
End Sub
Sub buildTuningCurvesIntoSigmaplot()
'    ImportFrom.Show
    
'    If doImport Then
'        Call processImport(True)
'    End If
    Call TransferToSigmaplot
End Sub

Sub detectTunedSegments()
    If outputWorkbook Is Nothing Then
        Set outputWorkbook = Application.ActiveWorkbook
    End If
    Dim iOutputOffset As Integer
    iOutputOffset = 2
    
    outputWorkbook.Worksheets("Likely Tuned Channels").Range("A2:D200").Clear

    Dim zOffsetSize As Long
    Dim iColOffset As Integer
    Dim iRowOffset As Integer

    Dim xCount As Integer
    Dim yCount As Integer

    zOffsetSize = outputWorkbook.Worksheets("Variables (do not edit)").Range("H3").Value
    iColOffset = outputWorkbook.Worksheets("Variables (do not edit)").Range("H5").Value
    iRowOffset = outputWorkbook.Worksheets("Variables (do not edit)").Range("H6").Value

    xCount = outputWorkbook.Worksheets("Variables (do not edit)").Range("H1").Value
    yCount = outputWorkbook.Worksheets("Variables (do not edit)").Range("H2").Value

    Dim xPos As Long
    Dim yPos As Long
   
    xPos = iColOffset + 1
    yPos = iRowOffset

    Dim dRowTotal As Double
    Dim dFirstRowTotal As Double
    
    Dim iRow As Integer
    Dim iCol As Integer
    Dim blnLooksGood As Boolean
    
    Dim dblMean() As Double
    ReDim dblMean(yCount - 1)
    Dim dblVar() As Double
    ReDim dblVar(yCount - 1)
    Dim iOffset As Integer

    Do
        dRowTotal = 0#
        dFirstRowTotal = 0#
        For iOffset = 0 To (yCount - 1)
            dblVar(iOffset) = 0#
            dblMean(iOffset) = 0#
        Next
        iOffset = 0
        
        If outputWorkbook.Worksheets("Means").Cells(yPos, xPos).Value <> "" Then
            blnLooksGood = True
            For iRow = (yPos + 2) To (yPos + yCount + 1) 'only want to look at the first 2 rows - after than there is no real guarantees
                For iCol = (xPos + 1) To (xPos + xCount)
                    dRowTotal = dRowTotal + outputWorkbook.Worksheets("Means").Cells(iRow, iCol).Value
                    dblVar(iOffset) = dblVar(iOffset) + ((outputWorkbook.Worksheets("Means").Cells(iRow, iCol).Value) ^ 2)
                Next
                
                dblVar(iOffset) = ((dblVar(iOffset) - ((dRowTotal ^ 2) / xCount)) / xCount)
                dblMean(iOffset) = dRowTotal / xCount
                iOffset = iOffset + 1
                If iRow > (yPos + 2) Then 'can only compare to previous row if not first row
                    If (dRowTotal * marginForGoodTuning) > dFirstRowTotal Then
                        blnLooksGood = False
                        Exit For
                    End If
                Else
                    If dblVar(0) < 0.05 Then 'if the variance is less than 0.1 it is probably a dead or noise channel - insufficient variability for even a moderate tuning curve?
                        blnLooksGood = False
                        Exit For
                    End If
                    dFirstRowTotal = dRowTotal
                End If
                dRowTotal = 0
            Next
            If blnLooksGood Then
                outputWorkbook.Worksheets("Likely Tuned Channels").Cells(iOutputOffset, 1).Value = outputWorkbook.Worksheets("Means").Cells(yPos, xPos).Value
                outputWorkbook.Worksheets("Likely Tuned Channels").Cells(iOutputOffset, 2).Value = yPos
                For iOffset = 0 To (yCount - 1)
                    outputWorkbook.Worksheets("Likely Tuned Channels").Cells(iOutputOffset, 4 + (3 * iOffset)).Value = dblVar(iOffset)
                    outputWorkbook.Worksheets("Likely Tuned Channels").Cells(iOutputOffset, 5 + (3 * iOffset)).Value = dblMean(iOffset)
                Next
                iOutputOffset = iOutputOffset + 1
            End If
            
            yPos = yPos + zOffsetSize
        Else
            Exit Do
        End If
    Loop
    
End Sub

Sub Broadcast_It()
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
        'lWindHandle = oDynWrap.FindWindowA("Afx:00400000:8:00010017:00000000:0002066D", vbNullString) 'find the SigmaPlot window
        lWindHandle = oDynWrap.FindWindowA(vbNullString, "SigmaPlot") 'find the SigmaPlot window
        iRet = oDynWrap.PostMessageA(lWindHandle, WM_COMMAND, MAKELPARAM(57604, 1), 0&) 'send the 'save as' command
        lDialogHandle = oDynWrap.FindWindowA("#32770", "Save As") 'get the dialog box
        lButtonHandle = oDynWrap.FindWindowEx(lDialogHandle, 0&, vbNullString, "&Save") 'get the save button
        iRet = oDynWrap.SendMessageA(lButtonHandle, BM_CLICK, 0&, 0&)
        iRet = oDynWrap.SendMessageA(lWindHandle, WM_COMMAND, MAKELPARAM(780, 0), 0&) 'send the 'close all notebooks' command
    Set oDynWrap = Nothing
End Sub

Function checkForMapExclusion(objFolder As Folder) As Variant
    Dim arrExclArr(1) As Variant
    
    arrExclArr(0) = ""
    Dim Files As Files
    Dim objFile As File

    Set Files = objFolder.Files

    For Each objFile In Files
        If LCase(objFile.Name) = "exclude from cf saving.txt" Then
            arrExclArr(0) = "cf"
            arrExclArr(1) = readCommentFromFile(objFile)
            If arrExclArr(1) = "" Then
                arrExclArr(1) = "No message"
            End If
            Exit For
        End If

        If LCase(objFile.Name) = "exclude from map generation.txt" Then
            arrExclArr(0) = "generate"
            arrExclArr(1) = readCommentFromFile(objFile)
            If arrExclArr(1) = "" Then
                arrExclArr(1) = "No message"
            End If
            Exit For
        End If
    Next
    
    checkForMapExclusion = arrExclArr
End Function

Function readCommentFromFile(objFile As File) As String
    Dim ts As TextStream
    Set ts = objFile.OpenAsTextStream
    readCommentFromFile = ts.ReadLine
    ts.Close
End Function

Function buildSearchStringList(ByRef arrOtherEp As Variant, ByRef arrOtherEpocKeys As Variant, ByRef searchStringList As Variant, iOtherEpocNum As Integer, strSearchString As String)
    Dim i As Integer
    
    Dim strAddedSearchString As String
    
    For i = 0 To UBound(arrOtherEpocKeys(iOtherEpocNum))
        If arrOtherEp(iOtherEpocNum) <> "Channel" Then
            'add to search string
            strAddedSearchString = strSearchString & arrOtherEp(iOtherEpocNum) & " = " & CStr(arrOtherEpocKeys(iOtherEpocNum)(i)) & " and "
        End If
        If iOtherEpocNum < UBound(arrOtherEp) Then
            'there are still more epocs to add to the search
            Call buildSearchStringList(arrOtherEp, arrOtherEpocKeys, searchStringList, iOtherEpocNum + 1, strAddedSearchString)
        Else
            If Right(strAddedSearchString, 5) = " and " Then 'this should always be the case - should be a trailing 'and' to remove
                strAddedSearchString = Left(strAddedSearchString, Len(strAddedSearchString) - 5)
            Else
                strAddedSearchString = strAddedSearchString
            End If
            Call searchStringList.Add(strAddedSearchString, True)
        End If
    Next
        
End Function


Function testSearchStringList(ByRef objTTX As Variant, ByRef searchStringList As Variant)

    Dim arrKeys As Variant
    arrKeys = searchStringList.Keys
    Dim strFilter As String

    Dim i As Integer
    Dim varReturn As Variant

    For i = 0 To UBound(arrKeys)
        'we have reached the end of the list of epocs - can actually do a search now
        If Right(arrKeys(i), 5) = " and " Then 'this should always be the case - should be a trailing 'and' to remove
            strFilter = Left(arrKeys(i), Len(arrKeys(i)) - 5)
        Else
            strFilter = arrKeys(i)
        End If
        
        Call objTTX.SetFilterWithDescEx(strFilter)
        varReturn = objTTX.GetEpocsExV(stimStartEpoc, 0)
        
        If Not IsArray(varReturn) Then
            searchStringList.Remove (arrKeys(i))
        End If
    Next

End Function









