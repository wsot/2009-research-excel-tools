Attribute VB_Name = "HRDetection"
Option Explicit
Const minAcceptableHR = 180
Const maxAcceptableHR = 650
Const maxInterBeatOverrun = 1.6
Const maxInterBeatUnderrun = 0.6

Sub processHeartRate()
  
    Dim detectedHR As Double
    Dim overlyCloseBeats As Integer
    Dim interpolations As Integer
    Dim longestInterpolation As Long
    Dim interpolationDuration As Long
    Dim interpolatedBeatsMax As Double
    Dim interpolatedBeats As Double
    
    Dim beatWorksheet As Worksheet
    Set beatWorksheet = Worksheets("Beat points from LabChart")
    
    Dim commentWorksheet As Worksheet
    Set commentWorksheet = Worksheets("Trial points from LabChart")
    
    Dim interpWS As Worksheet
    Set interpWS = Worksheets("Interpolations")
          
    Dim checkForDropouts As Boolean
    checkForDropouts = False

    Dim iTrialNum As Integer
    Dim blnNeedToSave As Boolean

    Dim lLastOffset As Long

    Dim lPretrialSampStart As Long
    Dim lTrialSampStart As Long
    Dim lTrialSampEnd As Long

    Dim cumulativeInterpolations As Long

    iTrialNum = 1

    Do
        cumulativeInterpolations = 0
    
        If commentWorksheet.Cells(iTrialNum + 1, 1) = "" Then 'go until there are no more trial numbers
            Exit Do
        End If
        
        interpWS.Cells(1, ((iTrialNum - 1) * 5) + 1).Value = "Trial " & iTrialNum
        interpWS.Cells(2, ((iTrialNum - 1) * 5) + 1).Value = "For range"
        interpWS.Cells(2, ((iTrialNum - 1) * 5) + 1).Value = "LC Sample"
        interpWS.Cells(2, ((iTrialNum - 1) * 5) + 2).Value = "LC Time"
        
        lPretrialSampStart = commentWorksheet.Cells(iTrialNum + 1, 2)
        lTrialSampStart = commentWorksheet.Cells(iTrialNum + 1, 3)
        lTrialSampEnd = commentWorksheet.Cells(iTrialNum + 1, 4)
        
        Call detectHROnSelection(lPretrialSampStart + 20000, lPretrialSampStart + 60000, detectedHR, overlyCloseBeats, interpolations, longestInterpolation, interpolationDuration, interpolatedBeatsMax, interpolatedBeats, iTrialNum, "10-30s", 0)

        cumulativeInterpolations = interpolations

        Worksheets("Settings").Range("O" & (iTrialNum + 1)).Value = detectedHR
        Worksheets("HR detection").Range("A" & (iTrialNum + 2)).Value = detectedHR
        Worksheets("HR detection").Range("B" & (iTrialNum + 2)).Value = interpolations
        Worksheets("HR detection").Range("C" & (iTrialNum + 2)).Value = interpolatedBeats
        Worksheets("HR detection").Range("D" & (iTrialNum + 2)).Value = longestInterpolation
        Worksheets("HR detection").Range("E" & (iTrialNum + 2)).Value = interpolatedBeatsMax

        Call detectHROnSelection(lTrialSampStart - 8000, lTrialSampStart, detectedHR, overlyCloseBeats, interpolations, longestInterpolation, interpolationDuration, interpolatedBeatsMax, interpolatedBeats, iTrialNum, "-4-0s", cumulativeInterpolations)
        
        cumulativeInterpolations = cumulativeInterpolations + interpolations
        
        Worksheets("Settings").Range("P" & (iTrialNum + 1)).Value = detectedHR
        Worksheets("HR detection").Range("G" & (iTrialNum + 2)).Value = detectedHR
        Worksheets("HR detection").Range("H" & (iTrialNum + 2)).Value = interpolations
        Worksheets("HR detection").Range("I" & (iTrialNum + 2)).Value = interpolatedBeats
        Worksheets("HR detection").Range("J" & (iTrialNum + 2)).Value = longestInterpolation
        Worksheets("HR detection").Range("K" & (iTrialNum + 2)).Value = interpolatedBeatsMax
        
        Call detectHROnSelection(lTrialSampStart + 10000, lTrialSampStart + 18000, detectedHR, overlyCloseBeats, interpolations, longestInterpolation, interpolationDuration, interpolatedBeatsMax, interpolatedBeats, iTrialNum, "-5-9s", cumulativeInterpolations)
        
        Worksheets("Settings").Range("Q" & (iTrialNum + 1)).Value = detectedHR
        Worksheets("HR detection").Range("M" & (iTrialNum + 2)).Value = detectedHR
        Worksheets("HR detection").Range("N" & (iTrialNum + 2)).Value = interpolations
        Worksheets("HR detection").Range("O" & (iTrialNum + 2)).Value = interpolatedBeats
        Worksheets("HR detection").Range("P" & (iTrialNum + 2)).Value = longestInterpolation
        Worksheets("HR detection").Range("Q" & (iTrialNum + 2)).Value = interpolatedBeatsMax

        iTrialNum = iTrialNum + 1
    Loop
    
End Sub


Sub detectHROnSelection(lStartPoint As Long, lEndPoint As Long, ByRef detectedHR, ByRef overlyCloseBeats, ByRef interpolations, ByRef longestInterpolation, ByRef interpolationDuration, ByRef interpolatedBeatsMax, ByRef interpolatedBeats, iTrialNum As Integer, strRangeTitle As String, iInterpOffset As Long)

    detectedHR = 0
    overlyCloseBeats = 0
    interpolations = 0
    interpolatedBeats = 0
    longestInterpolation = 0
    interpolatedBeatsMax = 0
    
    Dim returnFailed As Boolean
    
    Dim strInterpolationAddr As String
    
    Dim beatWorksheet As Worksheet
    Set beatWorksheet = Worksheets("Beat points from LabChart")
    
    Dim interpWS As Worksheet
    Set interpWS = Worksheets("Interpolations")
      
    interpWS.Cells(1, ((iTrialNum - 1) * 5) + 1).Value = "Trial " & iTrialNum
    interpWS.Cells(2, ((iTrialNum - 1) * 5) + 1).Value = "For range"
    interpWS.Cells(2, ((iTrialNum - 1) * 5) + 1).Value = "LC Sample"
    interpWS.Cells(2, ((iTrialNum - 1) * 5) + 2).Value = "LC Time"
    
    Dim beatCount As Double
    beatCount = 0#
    Dim beatDuration As Long
    Dim currentBeatOffset As Long 'offset (in columns) from first beat
    Dim currentBeatSamp As Long 'current beat position in time (in samples)
    Dim prevAcceptedBeatSamp As Long 'previous beat position in time (in samples)

    Dim thisInterpolationBeats As Double
    Dim thisInterpolationDuration As Long
    Dim lPostBeatDuration As Long

    Dim lStartColNum As Long
    lStartColNum = getPrecedingBeatOffset(lStartPoint, iTrialNum) 'get offset for region start
    
    Const minInterBeatIntervalMsec = 80
    
    Dim minInterBeatIntervalSamples As Long
    minInterBeatIntervalSamples = minInterBeatIntervalMsec * 2

    Dim meanHeartRate As Double

    beatDuration = findStartBeatDuration(lStartColNum, iTrialNum)

    'if we couldn't detect a starting beat duration, we can't detect a drop-out; return HR of -1 (can't detect)
    If beatDuration = -1 Then
        detectedHR = -1
        Exit Sub
    End If
    
    currentBeatOffset = 1

    prevAcceptedBeatSamp = beatWorksheet.Cells(iTrialNum, lStartColNum).Value 'set the point of the first accepted beat to the starting beat
    currentBeatSamp = beatWorksheet.Cells(iTrialNum, lStartColNum + currentBeatOffset).Value
    
    Do
        If (currentBeatSamp - prevAcceptedBeatSamp) > (maxInterBeatOverrun * beatDuration) Then
            thisInterpolationDuration = (currentBeatSamp - prevAcceptedBeatSamp)
            strInterpolationAddr = beatWorksheet.Cells(iTrialNum, lStartColNum + currentBeatOffset).Address()
            interpolations = interpolations + 1
            
            interpWS.Cells(interpolations + iInterpOffset + 2, ((iTrialNum - 1) * 5) + 1).Value = strRangeTitle
            interpWS.Cells(interpolations + iInterpOffset + 2, ((iTrialNum - 1) * 5) + 2).Value = currentBeatSamp
            interpWS.Cells(interpolations + iInterpOffset + 2, ((iTrialNum - 1) * 5) + 3).Value = "'" & calculateLCTime(currentBeatSamp)
            
            'Inter-beat variation is more than what is allowable, so probably missed beats - calculate beat duration after gap for interpolation
            lPostBeatDuration = (beatWorksheet.Cells(iTrialNum, lStartColNum + currentBeatOffset + 1).Value - currentBeatSamp)
            If lPostBeatDuration > ((maxInterBeatOverrun + (maxInterBeatOverrun * 0.1)) * beatDuration) Then   'check if the next beat might also have missed
                'next beat also looks like a miss; check the following beat
                lPostBeatDuration = (beatWorksheet.Cells(iTrialNum, lStartColNum + currentBeatOffset + 2).Value - beatWorksheet.Cells(iTrialNum, lStartColNum + currentBeatOffset + 1))
                If lPostBeatDuration > ((maxInterBeatOverrun + (maxInterBeatOverrun * 0.2)) * beatDuration) Then 'check if the next beat might also have missed. Give a bit more leeway on how much the duration can have changed, as it is more temporally distant
                    'beat after is also a miss. Give up the ghost.
                    returnFailed = True
                Else
                    thisInterpolationBeats = thisInterpolationDuration / ((beatDuration + lPostBeatDuration) / 2) 'calculate the number of beats to interpolate;
                End If
            Else
                thisInterpolationBeats = thisInterpolationDuration / ((beatDuration + lPostBeatDuration) / 2) 'calculate the number of beats to interpolate;
            End If
            
            beatCount = beatCount + thisInterpolationBeats
            
            'update cumulative information
            If thisInterpolationDuration > longestInterpolation Then
                longestInterpolation = thisInterpolationDuration
            End If
            If thisInterpolationBeats > interpolatedBeatsMax Then
                interpolatedBeatsMax = thisInterpolationBeats
            End If
            interpolationDuration = interpolationDuration + thisInterpolationDuration
            interpolatedBeats = interpolatedBeats + thisInterpolationBeats
        Else
            beatCount = beatCount + 1#
            beatDuration = ((currentBeatSamp - prevAcceptedBeatSamp) + beatDuration) / 2
        End If
        
        prevAcceptedBeatSamp = currentBeatSamp
        currentBeatOffset = currentBeatOffset + 1
        currentBeatSamp = beatWorksheet.Cells(iTrialNum, lStartColNum + currentBeatOffset).Value
        If currentBeatSamp > lEndPoint Then 'check if we've overrun our endpoint
            Exit Do
        End If
    Loop
    
    If returnFailed Then
        detectedHR = -1
    Else
        detectedHR = beatCount / ((((prevAcceptedBeatSamp - beatWorksheet.Cells(iTrialNum, lStartColNum).Value) / 2000) / 60))
    End If

End Sub


Function findStartBeatDuration(lStartColNum As Long, iTrialNum As Integer)

    Dim beatDuration As Long
    'Dim strStartLoc As String
    'strStartLoc = Worksheets("Beat points from LabChart").Cells(iTrialNum, lStartColNum).Address()

    Dim lastFourBeats(3) As Long
    Dim HR(2) As Double
    
    lastFourBeats(0) = Worksheets("Beat points from LabChart").Cells(iTrialNum, lStartColNum).Value
    lastFourBeats(1) = Worksheets("Beat points from LabChart").Cells(iTrialNum, lStartColNum - 1).Value
    lastFourBeats(2) = Worksheets("Beat points from LabChart").Cells(iTrialNum, lStartColNum - 2).Value
    lastFourBeats(3) = Worksheets("Beat points from LabChart").Cells(iTrialNum, lStartColNum - 3).Value
    
    HR(0) = 1 / ((((lastFourBeats(0) - lastFourBeats(1)) / 2000) / 60))
    HR(1) = 1 / ((((lastFourBeats(1) - lastFourBeats(2)) / 2000) / 60))
    HR(2) = 1 / ((((lastFourBeats(2) - lastFourBeats(3)) / 2000) / 60))

    'check variation is within acceptable bounds, otherwise probably missed beat
    If (HR(1) / HR(0) > maxInterBeatOverrun) Or (HR(2) / HR(0) > maxInterBeatOverrun) Or (HR(1) / HR(0) < maxInterBeatUnderrun) Or (HR(2) / HR(0) < maxInterBeatUnderrun) Or (HR(0) > maxAcceptableHR) Or (HR(0) < minAcceptableHR) Or (HR(1) > maxAcceptableHR) Or (HR(1) < minAcceptableHR) Then
        If (HR(2) / HR(1) > maxInterBeatOverrun) Or (HR(2) / HR(1) > maxInterBeatUnderrun) Or (HR(1) > maxAcceptableHR) Or (HR(1) < minAcceptableHR) Or (HR(2) > maxAcceptableHR) Or (HR(2) < minAcceptableHR) Then
            beatDuration = -1
        Else
            beatDuration = (lastFourBeats(1) - lastFourBeats(2)) 'beat duration in samples
        End If
    Else
            beatDuration = (lastFourBeats(0) - lastFourBeats(1)) 'beat duration in samples
    End If

    findStartBeatDuration = beatDuration

End Function


Function getPrecedingBeatOffset(lSampNum As Long, iTrialNum As Integer)
    Dim ws As Worksheet
    Set ws = Worksheets("Beat points from LabChart")
    Dim lOffset As Long
    
    lOffset = 1
    
    While (ws.Cells(iTrialNum, lOffset).Value < lSampNum) And (ws.Cells(iTrialNum, lOffset).Value <> "") 'check we haven't passed our desired location
        lOffset = lOffset + ((lSampNum - ws.Cells(iTrialNum, lOffset).Value) / 600) + 10 'estimate how many columns along we want to go - work out the number of samples difference and divide by 600 samples (1 beat at 200bpm) to estimate the offset
    Wend
    While (ws.Cells(iTrialNum, lOffset).Value > lSampNum) Or (ws.Cells(iTrialNum, lOffset).Value = "") 'move backwards until we are at the beat before the designated start point
        lOffset = lOffset - 1
    Wend
    
    getPrecedingBeatOffset = lOffset
End Function




