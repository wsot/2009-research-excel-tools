Attribute VB_Name = "DetectDeadZone"
Option Explicit
Const samplesForDeadzone = 4000


Sub buildDeadzoneLists()
    Dim beatLists As Worksheet
    Dim deadZoneLists As Worksheet
    Dim iTrialNum As Integer
    Dim iCol As Long
    Dim iDeadzoneNum As Long
    
    Dim interBeatDiff As Long
    
    Set beatLists = Worksheets("Beat points from LabChart")
    Set deadZoneLists = Worksheets("Deadzones")
    iTrialNum = 1
    iCol = 1
    iDeadzoneNum = 0
 
    While beatLists.Cells(iTrialNum, 1).Value <> ""
        deadZoneLists.Cells(1, ((iTrialNum - 1) * 4) + 1).Value = "Trial " & iTrialNum
        deadZoneLists.Cells(2, ((iTrialNum - 1) * 4) + 1).Value = "LC Sample"
        deadZoneLists.Cells(2, ((iTrialNum - 1) * 4) + 2).Value = "LC Time"
        
        While beatLists.Cells(iTrialNum, iCol).Value <> ""
            If iCol > 1 Then
                interBeatDiff = (beatLists.Cells(iTrialNum, iCol).Value - beatLists.Cells(iTrialNum, iCol - 1).Value)
                If interBeatDiff >= samplesForDeadzone Then
                    iDeadzoneNum = iDeadzoneNum + 1
                    deadZoneLists.Cells(iDeadzoneNum + 2, ((iTrialNum - 1) * 4) + 1).Value = beatLists.Cells(iTrialNum, iCol - 1).Value
                    deadZoneLists.Cells(iDeadzoneNum + 2, ((iTrialNum - 1) * 4) + 2).Value = "'" & calculateLCTime(beatLists.Cells(iTrialNum, iCol - 1).Value)
                End If
            End If
            iCol = iCol + 1
        Wend
        iCol = 1
        iTrialNum = iTrialNum + 1
    Wend
    
End Sub

Function calculateLCTime(lSampleNum) As String

    Dim iHrs As Integer
    Dim iMins As Integer
    Dim iSecs As Integer
    Dim iMSecs As Integer
    
    iHrs = Int(lSampleNum / 2000 / 60 / 60)
    iMins = (Int(lSampleNum / 2000 / 60) Mod 60)
    iSecs = (Int(lSampleNum / 2000) Mod 60)
    iMSecs = Int(lSampleNum / 1000) Mod 1000

    calculateLCTime = Right("00" & iHrs, 2) & ":" & Right("00" & iMins, 2) & ":" & Right("00" & iSecs, 2) & "." & Right("0000" & iMSecs, 4)

End Function

