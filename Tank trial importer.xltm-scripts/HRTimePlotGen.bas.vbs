Attribute VB_Name = "HRTimePlotGen"
Option Explicit

Const alignToZeroPoint = True
Const avgWithXEitherSide = 2
Const maxAllowableInstantChangeProp = 0.2


Sub generateHrAtTimePoints()
    Application.Calculation = xlCalculationManual
    Dim iTrialNum As Integer
    Dim lStartSample As Long
    Dim lCurrSample As Long
'    Dim arrCells() As Double
    
'    ReDim arrCells(2 * avgWidthXEitherSide + 1)
    
    Dim lTrialStartSample As Long
    Dim lRealTrialStartSample As Long
    Dim lTrialEndSample As Long
    
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
    
    Worksheets("HRLine").UsedRange.Clear
    
    Worksheets("HRLine").Cells(1, 1) = "Trial"
    For l100msCounter = 0 To 160
        Worksheets("HRLine").Cells(1, 2 + l100msCounter) = (l100msCounter - 80) * 100
    Next
        
    
    Do
        l100msCounter = 0
        'dZeroPointHR = 0
        dStartingHR = 0
        lOutColNum = 1
        lInColNum = 4
        If Worksheets("-4.5-9.5s HRs").Cells(1 + ((iTrialNum - 1) * 2), 3) <> "" Then
            lTrialStartSample = Worksheets("Trial points from LabChart").Cells(iTrialNum + 1, 3) - 16000
            lRealTrialStartSample = Worksheets("Trial points from LabChart").Cells(iTrialNum + 1, 3)
            lTrialEndSample = Worksheets("Trial points from LabChart").Cells(iTrialNum + 1, 3) + 16000
            
            lStartSample = Worksheets("-4.5-9.5s HRs").Cells(1 + ((iTrialNum - 1) * 2), 4)
            Worksheets("HRLine").Cells(iTrialNum + 1, 1).Value = "Trial " & iTrialNum
            l100msCounter = lStartSample
            
            If alignToZeroPoint Then
                While Worksheets("-4.5-9.5s HRs").Cells(2 + ((iTrialNum - 1) * 2), lInColNum).Value <> "" And dStartingHR = 0
                    lCurrSample = Worksheets("-4.5-9.5s HRs").Cells(1 + ((iTrialNum - 1) * 2), lInColNum).Value
                    While l100msCounter < lCurrSample And dStartingHR = 0
                        l100msCounter = l100msCounter + 200
                        If l100msCounter >= lRealTrialStartSample Then
                            dCurrVal = Worksheets("-4.5-9.5s HRs").Cells(2 + ((iTrialNum - 1) * 2), lInColNum - 1).Value
                            dCurrValSum = dCurrVal
                            iCtr = 1
                            For iIter = 1 To avgWithXEitherSide
                                If Not lInColNum - 1 - iIter < 4 Then
                                    sNextVal = Worksheets("-4.5-9.5s HRs").Cells(2 + ((iTrialNum - 1) * 2), lInColNum - 1 + iIter).Value
                                    If Not sNextVal = "" Then
                                        iCtr = iCtr + 2
                                        dCurrValSum = dCurrValSum + Worksheets("-4.5-9.5s HRs").Cells(2 + ((iTrialNum - 1) * 2), lInColNum - 1 - iIter).Value
                                        dCurrValSum = dCurrValSum + CDbl(sNextVal)
                                    End If
                                End If
                            Next
                            dCurrValSum = dCurrValSum / iCtr
                            dStartingHR = dCurrValSum
                            Worksheets("HRLine").Cells(iTrialNum + 1, 190).Value = dCurrValSum
                        End If
                    Wend
                    lInColNum = lInColNum + 1
                Wend
            
                lInColNum = 4
                l100msCounter = lStartSample
            Else
                While Worksheets("-4.5-9.5s HRs").Cells(2 + ((iTrialNum - 1) * 2), lInColNum).Value <> "" And dStartingHR = 0
                    lCurrSample = Worksheets("-4.5-9.5s HRs").Cells(1 + ((iTrialNum - 1) * 2), lInColNum).Value
                    While l100msCounter < lCurrSample And dStartingHR = 0
                        l100msCounter = l100msCounter + 200
                        If l100msCounter >= lTrialStartSample Then
                            dStartingHR = Worksheets("-4.5-9.5s HRs").Cells(2 + ((iTrialNum - 1) * 2), lInColNum - 1).Value
                        End If
                    Wend
                    lInColNum = lInColNum + 1
                Wend
            
                lInColNum = 4
                l100msCounter = lStartSample
'                dStartingHR = Worksheets("-4.5-9.5s HRs").Cells(2 + ((iTrialNum - 1) * 2), lInColNum).Value
'                Worksheets("HRLine").Cells(iTrialNum + 1, lOutColNum + 1).Value = 1#
            End If
            
            dLastVal = 0
            While Worksheets("-4.5-9.5s HRs").Cells(2 + ((iTrialNum - 1) * 2), lInColNum).Value <> ""
                lCurrSample = Worksheets("-4.5-9.5s HRs").Cells(1 + ((iTrialNum - 1) * 2), lInColNum).Value
                While l100msCounter < lCurrSample
                    l100msCounter = l100msCounter + 200
                    If l100msCounter > lTrialStartSample And l100msCounter <= (lTrialEndSample + 200) Then
                        lOutColNum = lOutColNum + 1
                        dCurrVal = Worksheets("-4.5-9.5s HRs").Cells(2 + ((iTrialNum - 1) * 2), lInColNum - 1).Value
                        dCurrValSum = dCurrVal
                        iCtr = 1
                        For iIter = 1 To avgWithXEitherSide
                            If Not lInColNum - 1 - iIter < 4 Then
                                sNextVal = Worksheets("-4.5-9.5s HRs").Cells(2 + ((iTrialNum - 1) * 2), lInColNum - 1 + iIter).Value
                                If Not sNextVal = "" Then
                                    iCtr = iCtr + 2
                                    dCurrValSum = dCurrValSum + Worksheets("-4.5-9.5s HRs").Cells(2 + ((iTrialNum - 1) * 2), lInColNum - 1 - iIter).Value
                                    dCurrValSum = dCurrValSum + CDbl(sNextVal)
                                End If
                            End If
                        Next
                        dCurrValSum = dCurrValSum / iCtr
                        If (dLastVal <> 0) And Abs(dCurrVal - dLastVal) > (dLastVal * maxAllowableInstantChangeProp) Then
                                Worksheets("HRLine").Cells(iTrialNum + 1, lOutColNum).Value = "x" & (dCurrValSum / dStartingHR)
                        Else
                            dLastVal = dCurrVal
                            Worksheets("HRLine").Cells(iTrialNum + 1, lOutColNum).Value = (dCurrValSum / dStartingHR)
                        End If
                    End If
                Wend
                lInColNum = lInColNum + 1
            Wend
        End If
        iTrialNum = iTrialNum + 1
        If iTrialNum > 50 Then
            Exit Do
        End If
    Loop
    
End Sub






