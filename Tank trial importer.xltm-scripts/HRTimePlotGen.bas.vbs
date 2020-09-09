Attribute VB_Name = "HRTimePlotGen"
Option Explicit

Global Const REFPOINT_LASTTRIALEND = "REFPOINT_LASTTRIALEND"
Global Const REFPOINT_THISTRIALSTART = "REFPOINT_THISTRIALSTART"
Global Const REFPOINT_THISTRIALEND = "REFPOINT_THISTRIALEND"

Const alignToZeroPoint = True
Const avgWithXEitherSide = 2
Const maxAllowableInstantChangeProp = 0.2
Global Const FILTER_EXCESS_VARIABILITY = False


Sub generateHrAtTimePoints(lTrialStartSampleOffset As Long, lRealTrialStartSampleOffset As Long, lTrialEndSampleOffset As Long, sRefPoint As String, sourceWorksheet As Worksheet, outputWSName As String, sProcessType As String)
    Application.Calculation = xlCalculationManual
    Dim iTrialNum As Integer
    Dim lStartSample As Long
    Dim lCurrSample As Long
'    Dim arrCells() As Double
    
'    ReDim arrCells(2 * avgWidthXEitherSide + 1)
    
    Dim lRefPointValue As Long
    Dim lTrialStartSample As Long
    Dim lRealTrialStartSample As Long
    Dim lTrialEndSample As Long
    
    Dim outputWS As Worksheet
    If Not WorksheetExists(outputWSName) Then
        'Call Worksheets("HRLine").Copy(Worksheets("HRLine"))
        Call Worksheets("HRLine").Copy(, Worksheets("HRLine"))
        Set outputWS = Worksheets("HRLine (2)")
        outputWS.Name = outputWSName
    Else
        Set outputWS = Worksheets(outputWSName)
    End If
    Call outputWS.Select
    
    Dim dStartingHR As Double
    'Dim dZeroPointHR As Double
    Dim dLastVal As Double
    Dim dCurrVal As Double
    Dim dCurrValSum As Double
    Dim sNextVal As String
    
    Dim iIter As Integer
    Dim iCtr As Integer
            
    
    Dim l100msCounter As Long
    
    Dim lInColNum As Long
    Dim lOutColNum As Long
    iTrialNum = 1
    
    outputWS.UsedRange.Clear
    
    outputWS.Cells(1, 1) = "Trial"
    'For l100msCounter = 0 To ((lTrialEndSampleOffset - lTrialStartSampleOffset / 200))
    For l100msCounter = 0 To ((lTrialEndSampleOffset - lTrialStartSampleOffset) / 200)
        outputWS.Cells(1, 2 + l100msCounter) = (l100msCounter - 80) * 100
    Next
    
    Do
        l100msCounter = 0
        'dZeroPointHR = 0
        dStartingHR = 0
        lOutColNum = 1
        lInColNum = 4
        Dim bTooMuchVariability As Boolean
        Dim sExcessVariabilityLocations As String
        Dim bAlreadyProcessedThisSample As Boolean
        'sourceWorksheet = Worksheets("-8.5-8.5s HRs")
        If sourceWorksheet.Cells(1 + ((iTrialNum - 1) * 2), 3) <> "" Then
            bTooMuchVariability = False
            sExcessVariabilityLocations = ""
            
            Select Case sRefPoint
                Case REFPOINT_LASTTRIALEND:
                    lRefPointValue = Worksheets("Trial points from LabChart").Cells(iTrialNum + 1, 2)
                Case REFPOINT_THISTRIALSTART:
                    lRefPointValue = Worksheets("Trial points from LabChart").Cells(iTrialNum + 1, 3)
                Case REFPOINT_THISTRIALEND:
                    lRefPointValue = Worksheets("Trial points from LabChart").Cells(iTrialNum + 1, 4)
            End Select
            
            
            lTrialStartSample = lRefPointValue + lTrialStartSampleOffset
            lRealTrialStartSample = lRefPointValue + lRealTrialStartSampleOffset
            lTrialEndSample = lRefPointValue + lTrialEndSampleOffset
            
            lStartSample = sourceWorksheet.Cells(1 + ((iTrialNum - 1) * 2), 4)
            outputWS.Cells(iTrialNum + 1, 1).Value = "Trial " & iTrialNum
            l100msCounter = lStartSample
            
            If alignToZeroPoint Then
                While sourceWorksheet.Cells(2 + ((iTrialNum - 1) * 2), lInColNum).Value <> "" And dStartingHR = 0
                    lCurrSample = sourceWorksheet.Cells(1 + ((iTrialNum - 1) * 2), lInColNum).Value
                    While l100msCounter < lCurrSample And dStartingHR = 0
                        l100msCounter = l100msCounter + 200
                        If l100msCounter >= lRealTrialStartSample Then
                            dCurrVal = sourceWorksheet.Cells(2 + ((iTrialNum - 1) * 2), lInColNum - 1).Value
                            dCurrValSum = dCurrVal
                            iCtr = 1
                            For iIter = 1 To avgWithXEitherSide
                                If Not lInColNum - 1 - iIter < 4 Then
                                    sNextVal = sourceWorksheet.Cells(2 + ((iTrialNum - 1) * 2), lInColNum - 1 + iIter).Value
                                    If Not sNextVal = "" Then
                                        iCtr = iCtr + 2
                                        dCurrValSum = dCurrValSum + sourceWorksheet.Cells(2 + ((iTrialNum - 1) * 2), lInColNum - 1 - iIter).Value
                                        dCurrValSum = dCurrValSum + CDbl(sNextVal)
                                    End If
                                End If
                            Next
                            dCurrValSum = dCurrValSum / iCtr
                            dStartingHR = dCurrValSum
                            outputWS.Cells(iTrialNum + 1, 190).Value = dCurrValSum
                        End If
                    Wend
                    lInColNum = lInColNum + 1
                Wend
            
                lInColNum = 4
                l100msCounter = lStartSample
            Else
                While sourceWorksheet.Cells(2 + ((iTrialNum - 1) * 2), lInColNum).Value <> "" And dStartingHR = 0
                    lCurrSample = sourceWorksheet.Cells(1 + ((iTrialNum - 1) * 2), lInColNum).Value
                    While l100msCounter < lCurrSample And dStartingHR = 0
                        l100msCounter = l100msCounter + 200
                        If l100msCounter >= lTrialStartSample Then
                            dStartingHR = sourceWorksheet.Cells(2 + ((iTrialNum - 1) * 2), lInColNum - 1).Value
                        End If
                    Wend
                    lInColNum = lInColNum + 1
                Wend
            
                lInColNum = 4
                l100msCounter = lStartSample
'                dStartingHR = sourceWorksheet.Cells(2 + ((iTrialNum - 1) * 2), lInColNum).Value
'                outputWS.Cells(iTrialNum + 1, lOutColNum + 1).Value = 1#
            End If
            
            dLastVal = 0
            While sourceWorksheet.Cells(2 + ((iTrialNum - 1) * 2), lInColNum).Value <> ""
                lCurrSample = sourceWorksheet.Cells(1 + ((iTrialNum - 1) * 2), lInColNum).Value
                While l100msCounter < lCurrSample
                    l100msCounter = l100msCounter + 200
                    If l100msCounter > lTrialStartSample And l100msCounter <= (lTrialEndSample + 200) Then
                        lOutColNum = lOutColNum + 1
                        dCurrVal = sourceWorksheet.Cells(2 + ((iTrialNum - 1) * 2), lInColNum - 1).Value
                        dCurrValSum = dCurrVal
                        iCtr = 1
                        For iIter = 1 To avgWithXEitherSide
                            If Not lInColNum - 1 - iIter < 4 Then
                                sNextVal = sourceWorksheet.Cells(2 + ((iTrialNum - 1) * 2), lInColNum - 1 + iIter).Value
                                If Not sNextVal = "" Then
                                    iCtr = iCtr + 2
                                    dCurrValSum = dCurrValSum + sourceWorksheet.Cells(2 + ((iTrialNum - 1) * 2), lInColNum - 1 - iIter).Value
                                    dCurrValSum = dCurrValSum + CDbl(sNextVal)
                                End If
                            End If
                        Next
                        dCurrValSum = dCurrValSum / iCtr
                        If (dLastVal <> 0) And Abs(dCurrVal - dLastVal) > (dLastVal * maxAllowableInstantChangeProp) Then
                            If Not bAlreadyProcessedThisSample Then
                                sExcessVariabilityLocations = sExcessVariabilityLocations & sourceWorksheet.Cells(2 + ((iTrialNum - 1) * 2), lInColNum).Address & " "
                                bAlreadyProcessedThisSample = True
                            End If
                            If sProcessType = PROCESSTYPE_NORMALISED Then
                                outputWS.Cells(iTrialNum + 1, lOutColNum).Value = (dCurrValSum / dStartingHR)
                            ElseIf sProcessType = PROCESSTYPE_DELTA Then
                                'outputWS.Cells(iTrialNum + 1, lOutColNum).Value = "x" & (dCurrValSum)
                                outputWS.Cells(iTrialNum + 1, lOutColNum).Value = (dCurrValSum - dStartingHR)
                            ElseIf sProcessType = PROCESSTYPE_RAW Then
                                outputWS.Cells(iTrialNum + 1, lOutColNum).Value = dCurrValSum
                            End If
                            bTooMuchVariability = True
                        Else
                            dLastVal = dCurrVal
                            If sProcessType = PROCESSTYPE_NORMALISED Then
                                outputWS.Cells(iTrialNum + 1, lOutColNum).Value = (dCurrValSum / dStartingHR)
                            ElseIf sProcessType = PROCESSTYPE_DELTA Then
                                'outputWS.Cells(iTrialNum + 1, lOutColNum).Value = dCurrValSum
                                outputWS.Cells(iTrialNum + 1, lOutColNum).Value = (dCurrValSum - dStartingHR)
                            ElseIf sProcessType = PROCESSTYPE_RAW Then
                                outputWS.Cells(iTrialNum + 1, lOutColNum).Value = dCurrValSum
                            End If
                        End If
                    End If
                Wend
                bAlreadyProcessedThisSample = False
                lInColNum = lInColNum + 1
            Wend
            If bTooMuchVariability Then
                If FILTER_EXCESS_VARIABILITY Then
                    outputWS.Range(outputWS.Cells(iTrialNum + 1, 1), outputWS.Cells(iTrialNum + 1, lOutColNum)).Clear
                End If
                outputWS.Cells(iTrialNum + 1, lOutColNum + 1).Value = "x"
                sourceWorksheet.Cells(1 + ((iTrialNum - 1) * 2), 1).Value = "x"
                sourceWorksheet.Cells(1 + ((iTrialNum - 1) * 2), 2).Value = sExcessVariabilityLocations
            Else
                outputWS.Cells(iTrialNum + 1, lOutColNum + 1).Value = ""
                sourceWorksheet.Cells(1 + ((iTrialNum - 1) * 2), 1).Value = ""
                sourceWorksheet.Cells(1 + ((iTrialNum - 1) * 2), 2).Value = ""
            End If
        End If
        iTrialNum = iTrialNum + 1
        If iTrialNum > 50 Then
            Exit Do
        End If
    Loop
    
    outputWS.Cells(iTrialNum + 1, 1).Value = "Mean"
    outputWS.Cells(iTrialNum + 2, 1).Value = "StdDev"
    For lOutColNum = 2 To (2 + ((lTrialEndSampleOffset - lTrialStartSampleOffset) / 200))
        outputWS.Cells(iTrialNum + 1, lOutColNum) = "=AVERAGE(" & outputWS.Cells(2, lOutColNum).Address & ":" & outputWS.Cells(iTrialNum, lOutColNum).Address & ")"
        outputWS.Cells(iTrialNum + 2, lOutColNum) = "=CONFIDENCE(0.05,STDEV(" & outputWS.Cells(2, lOutColNum).Address & ":" & outputWS.Cells(iTrialNum, lOutColNum).Address & "),COUNT(" & outputWS.Cells(2, lOutColNum).Address & ":" & outputWS.Cells(iTrialNum, lOutColNum).Address & "))"
    Next
        
    
End Sub







