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

Global Const marginForGoodTuning = 1#

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
        ByRef iRowOffset As Integer, _
        ByRef iColOffset As Integer, _
        ByRef arrOtherEp As Variant, _
        ByRef xAxisEp As String, _
        ByRef yAxisEp As String, _
        ByRef bReverseX As Boolean, _
        ByRef bReverseY As Boolean)
        
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

    
End Function

Sub bulkBuildTuningCurves()
    isFirstChart = True
'        Dim thisWorkbook As Workbook
    Set thisWorkbook = Application.ActiveWorkbook

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
'        Stop
        successfullyProcessedOffset = 25
        
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
            'Call Worksheets("Totals").UsedRange.ClearContents
            'Call Worksheets("StdDev").UsedRange.ClearContents
            'Call Worksheets("Means").UsedRange.ClearContents
            'Call Worksheets("N").UsedRange.ClearContents
            theTank = Left(theBlocks(i), InStr(theBlocks(i), ":") - 1)
            theBlock = Right(theBlocks(i), Len(theBlocks(i)) - Len(theTank) - 1)
            theTank = bulkImportRootDir & "\" & theTank
            
            thisWorkbook.Worksheets("Settings").Cells(successfullyProcessedOffset, 1).Value = theBlocks(i) & " processing"
            
            If i = 0 Then
                templatePath = getFilename(templatePath, theTank)
            End If
            Set outputWorkbook = Workbooks.Open(templatePath)
            
            If specifiedOutputDir = "" Then
                outputDir = getDirName("", theTank)
                outputFilename = outputDir & "\" & outputFilePrefix & theBlock
            Else
                outputDir = getDirName(specifiedOutputDir, theTank)
                If outputDir = "" Then
                    MsgBox ("Output directory " & outputDir & " could not be found." & vbCrLf & "Please update the path and try again")
                    Exit Sub
                End If
                outputFilename = outputDir & "\" & Replace(theTank, "\", ".") & "_" & outputFilePrefix & theBlock
            End If
            
            outputWorkbook.Worksheets("Variables (do not edit)").Range("B2").Value = theTank 'update the block on the worksheet
            outputWorkbook.Worksheets("Variables (do not edit)").Range("B3").Value = theBlock 'update the block on the worksheet
            outputWorkbook.Worksheets("Settings").Range("B18").Value = thisWorkbook.Worksheets("Settings").Range("B18").Value
            Call processImport(False)
            Call detectTunedSegments
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
                        Call transferCandidatesToSigmaplot(outputFilename & ".JNB")
                    Else
                        Call transferAllToSigmaplot(outputFilename & ".JNB")
                    End If
                End If
                If blnAutoclose Then
                    Call outputWorkbook.Close
                End If
            End If
            thisWorkbook.Worksheets("Settings").Cells(successfullyProcessedOffset, 2).Value = "Processed"
            successfullyProcessedOffset = successfullyProcessedOffset + 1
        Next
        Application.DisplayAlerts = True
    End If
End Sub

Sub buildTuningCurves()
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
                        Call transferCandidatesToSigmaplot(outputFilename & ".JNB")
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


Sub processImport(importIntoSigmaplot As Boolean)
    Dim lNumOfChans As Long

    If loadConfigParams(outputWorkbook, thisWorkbook, stimStartEpoc, dblBinWidth, dblIgnoreFirstMsec, lNumOfChans, iRowOffset, iColOffset, arrOtherEp, xAxisEp, yAxisEp, bReverseX, bReverseY) Then
        
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
        
        If Not connectToTDTReportError(connectToTDT(objTTX, False, theServer, theTank, theBlock)) Then 'returns false if an error occurred
            'index epochs - required to use filters
            Call objTTX.CreateEpocIndexing
            
            Dim dblStartTime As Double
            Dim dblEndTime As Double
            
            Dim varReturn As Variant
            
        '    Dim vXAxisKeys As Variant
        '    Dim vYAxisKeys As Variant
            
            vXAxisKeys = BuildEpocList(objTTX, xAxisEp, bReverseX)
            vYAxisKeys = BuildEpocList(objTTX, yAxisEp, bReverseY)
                
            Dim i As Long
            Dim j As Long
            Dim k As Long
            Dim l As Long
            
            Dim arrOtherEpocKeys() As Variant
            If UBound(arrOtherEp) <> -1 Then
                ReDim arrOtherEpocKeys(UBound(arrOtherEp))
                
                For i = 0 To UBound(arrOtherEp)
                    arrOtherEpocKeys(i) = BuildEpocList(objTTX, arrOtherEp(i), False)
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
                        Call processSearch(objTTX, arrOtherEp, arrOtherEpocKeys, 0, xAxisSearchString & yAxisSearchString, i + 1, j + 1, UBound(vYAxisKeys) + 3, iChanNum, "", xCount, yCount, zOffsetSize, lMaxHistHeight, lMaxHistMeanHeight)
                    Next
                Next
            End If
        
        '    Call writeAxes(theWorksheets, vXAxisKeys, vYAxisKeys, iColOffset, iRowOffset)
        
            Call objTTX.CloseTank
            Call objTTX.ReleaseServer
            
            outputWorkbook.Worksheets("Variables (do not edit)").Range("H1").Value = xCount
            outputWorkbook.Worksheets("Variables (do not edit)").Range("H2").Value = yCount
            outputWorkbook.Worksheets("Variables (do not edit)").Range("H3").Value = zOffsetSize
            outputWorkbook.Worksheets("Variables (do not edit)").Range("H4").Value = lMaxHistHeight
            outputWorkbook.Worksheets("Variables (do not edit)").Range("H5").Value = iColOffset
            outputWorkbook.Worksheets("Variables (do not edit)").Range("H6").Value = iRowOffset
            outputWorkbook.Worksheets("Variables (do not edit)").Range("H7").Value = lMaxHistMeanHeight
            
            'If importIntoSigmaplot Then
                'Call transferToSigmaplot(xCount, yCount, zOffsetSize, iColOffset, iRowOffset, lMaxHistHeight)
            'End If
        End If
    End If
End Sub

Sub writeAxes(colLabels As Variant, rowLabels As Variant, iColOffset, iRowOffset, zOffset)
    Dim j As Long
        
    For j = 0 To UBound(rowLabels)
        outputWorkbook.Worksheets("Totals").Cells(iRowOffset + j + 2 + zOffset, iColOffset + 1).Value = rowLabels(j)
        outputWorkbook.Worksheets("StdDev").Cells(iRowOffset + j + 2 + zOffset, iColOffset + 1).Value = rowLabels(j)
        outputWorkbook.Worksheets("Means").Cells(iRowOffset + j + 2 + zOffset, iColOffset + 1).Value = rowLabels(j)
        outputWorkbook.Worksheets("N").Cells(iRowOffset + j + 2 + zOffset, iColOffset + 1).Value = rowLabels(j)
    Next
    For j = 0 To UBound(colLabels)
        outputWorkbook.Worksheets("Totals").Cells(iRowOffset + zOffset + 1, j + 2).Value = colLabels(j)
        outputWorkbook.Worksheets("StdDev").Cells(iRowOffset + zOffset + 1, j + 2).Value = colLabels(j)
        outputWorkbook.Worksheets("Means").Cells(iRowOffset + zOffset + 1, j + 2).Value = colLabels(j)
        outputWorkbook.Worksheets("N").Cells(iRowOffset + zOffset + 1, j + 2).Value = colLabels(j)
    Next

End Sub





Function BuildEpocList(objTTX, AxisEp, bReverseOrder)
    'build list of epocs for the given axis epoc name
    
    Dim AxisList As Dictionary
    Set AxisList = New Dictionary
    
    Dim dblStartTime As Double
    Dim varReturn As Variant
    
    Dim i As Integer
    Dim j As Integer
    
    If AxisEp = "Channel" Then
        For i = 1 To thisWorkbook.Worksheets("Settings").Range("B3").Value
            Call AxisList.Add(i, 0)
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


Function processSearch(ByRef objTTX, ByRef arrOtherEp, ByRef arrOtherEpocKeys, iOtherEpocNum, strSearchString As String, xOffset, yOffset, zOffset, iChanNum, strTitle, ByRef xCount, ByRef yCount, ByRef zOffsetSize, ByRef lMaxHistHeight, ByRef lMaxHistMeanHeight)
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
            Call processSearch(objTTX, arrOtherEp, arrOtherEpocKeys, iOtherEpocNum + 1, strAddedSearchString, xOffset, yOffset, (zOffset * UBound(arrOtherEpocKeys(iOtherEpocNum))) + i, iChanNum, strAddedTitle, xCount, yCount, zOffsetSize, lMaxHistHeight, lMaxHistMeanHeight)
        Else
            'we have reached the end of the list of epocs - can actually do a search now
            If Right(strAddedSearchString, 5) = " and " Then 'this should always be the case - should be a trailing 'and' to remove
                strFilter = Left(strAddedSearchString, Len(strAddedSearchString) - 5)
            Else
                strFilter = strAddedSearchString
            End If
            Call objTTX.SetFilterWithDescEx(strFilter)
            
            If xOffset = 1 And yOffset = 1 Then
                outputWorkbook.Worksheets("Totals").Cells(iRowOffset + (i * zOffset), iColOffset + 1).Value = Left(strAddedTitle, Len(strAddedTitle) - 2)
                outputWorkbook.Worksheets("N").Cells(iRowOffset + (i * zOffset), iColOffset + 1).Value = Left(strAddedTitle, Len(strAddedTitle) - 2)
                outputWorkbook.Worksheets("Means").Cells(iRowOffset + (i * zOffset), iColOffset + 1).Value = Left(strAddedTitle, Len(strAddedTitle) - 2)
                outputWorkbook.Worksheets("StdDev").Cells(iRowOffset + (i * zOffset), iColOffset + 1).Value = Left(strAddedTitle, Len(strAddedTitle) - 2)
                Call writeAxes(vXAxisKeys, vYAxisKeys, iColOffset, iRowOffset, (i * zOffset))
            End If

            Call writeResults(objTTX, xOffset, yOffset, i * zOffset, iChanNum, lMaxHistHeight, lMaxHistMeanHeight)
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

Sub writeResults(ByRef objTTX, xOffset, yOffset, zOffset, iChanNum, ByRef lMaxHistHeight, ByRef lMaxHistMeanHeight)
    Dim varReturn As Variant
    Dim varChanData As Variant
    
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
    If IsArray(varReturn) Then
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
        End If
        If histMean > lMaxHistMeanHeight Then
            lMaxHistMeanHeight = histMean
        End If
        If histTmp > lMaxHistHeight Then
            lMaxHistHeight = histTmp
        End If
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


