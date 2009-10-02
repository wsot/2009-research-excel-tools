Attribute VB_Name = "HRTimePlotGen"
Option Explicit
Sub generateHrAtTimePoints()
    Dim iTrialNum As Integer
    Dim lStartSample As Long
    Dim lCurrSample As Long
    
    Dim lTrialStartSample As Long
    Dim lTrialEndSample As Long
    
    Dim dStartingHR As Double
    
    Dim l100msCounter As Long
    
    Dim lInColNum As Long
    Dim lOutColNum As Long
    iTrialNum = 1
    Do
        l100msCounter = 0
        lOutColNum = 1
        lInColNum = 4
        If Worksheets("-4.5-9.5s HRs").Cells(1 + ((iTrialNum - 1) * 2), 3) <> "" Then
            lTrialStartSample = Worksheets("Trial points from LabChart").Cells(iTrialNum + 1, 3) - 8000
            lTrialEndSample = Worksheets("Trial points from LabChart").Cells(iTrialNum + 1, 3) + 18000
            
            lStartSample = Worksheets("-4.5-9.5s HRs").Cells(1 + ((iTrialNum - 1) * 2), 4)
            dStartingHR = Worksheets("-4.5-9.5s HRs").Cells(2 + ((iTrialNum - 1) * 2), lInColNum).Value
            Worksheets("HRLine").Cells(iTrialNum + 1, 1).Value = "Trial " & iTrialNum
            Worksheets("HRLine").Cells(iTrialNum + 1, lOutColNum + 1).Value = 1#
            l100msCounter = lStartSample
            
            While Worksheets("-4.5-9.5s HRs").Cells(2 + ((iTrialNum - 1) * 2), lInColNum).Value <> ""
                lCurrSample = Worksheets("-4.5-9.5s HRs").Cells(1 + ((iTrialNum - 1) * 2), lInColNum).Value
                While l100msCounter < lCurrSample
                    l100msCounter = l100msCounter + 200
                    If l100msCounter >= lTrialStartSample And l100msCounter <= lTrialEndSample Then
                        lOutColNum = lOutColNum + 1
                        Worksheets("HRLine").Cells(iTrialNum + 1, lOutColNum + 1).Value = (Worksheets("-4.5-9.5s HRs").Cells(2 + ((iTrialNum - 1) * 2), lInColNum - 1).Value / dStartingHR)
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

