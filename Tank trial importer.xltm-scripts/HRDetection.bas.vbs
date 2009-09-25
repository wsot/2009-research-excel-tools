Attribute VB_Name = "HRDetection"
Option Explicit
Global minAcceptableHR As Integer
Global maxAcceptableHR As Integer
Global maxInterBeatOverrun As Double
Global maxInterBeatUnderrun As Double

Global maxPercOfBeatsInt As Double
Global maxSingleIntSamples As Double
Global maxSingleIntBeats As Double

Global hrProbHighlightCell As Range
Global hrNoteHighlightCell As Range
Global hrClearHighlightCell As Range

'Const minAcceptableHR = 180
'Const maxAcceptableHR = 650
'Const maxInterBeatOverrun = 1.6
'Const maxInterBeatUnderrun = 0.6

Sub processHeartRate()
  
    Dim maxAllowVariation As Double
    
    maxAllowVariation = Worksheets("Settings").Cells(5, 2).Value
    
    minAcceptableHR = Worksheets("Settings").Cells(2, 2).Value
    maxAcceptableHR = Worksheets("Settings").Cells(3, 2).Value
    maxInterBeatOverrun = 1 + maxAllowVariation
    maxInterBeatUnderrun = 1 - maxAllowVariation
  
    maxPercOfBeatsInt = Worksheets("Settings").Cells(9, 2).Value
    maxSingleIntSamples = Worksheets("Settings").Cells(10, 2).Value
    maxSingleIntBeats = Worksheets("Settings").Cells(11, 2).Value
    
    Set hrProbHighlightCell = Worksheets("Settings").Cells(13, 2)
    Set hrNoteHighlightCell = Worksheets("Settings").Cells(14, 2)
    Set hrClearHighlightCell = Worksheets("Settings").Cells(15, 2)
  
    Dim detectedHR As Double
    Dim overlyCloseBeats As Integer
    Dim abberantBeats As Integer
    Dim interpolations As Integer
    Dim longestInterpolation As Long
    Dim shortestInterpolation As Long
    Dim interpolationDuration As Long
    Dim interpolatedBeatsMax As Double
    Dim interpolatedBeatsMin As Double
    Dim interpolatedBeats As Double
    Dim beatCount As Double
    Dim proportionInterpolated As Double
    
    Dim beatWorksheet As Worksheet
    Set beatWorksheet = Worksheets("Beat points from LabChart")
    
    Dim commentWorksheet As Worksheet
    Set commentWorksheet = Worksheets("Trial points from LabChart")
    
    Dim interpWS As Worksheet
    Set interpWS = Worksheets("Interpolations")
    
    Dim overbWS As Worksheet
    Set overbWS = Worksheets("Overbeats")

    Dim abberWS As Worksheet
    Set abberWS = Worksheets("Abberant beats")

    Dim checkForDropouts As Boolean
    checkForDropouts = False

    Dim iTrialNum As Integer
    Dim iOutputNum As Integer
    Dim iColsPerOutput As Integer
    Dim blnNeedToSave As Boolean

    Dim lLastOffset As Long

    Dim lPretrialSampStart As Long
    Dim lTrialSampStart As Long
    Dim lTrialSampEnd As Long

    Dim thisStartPoint As Long
    Dim thisEndPoint As Long

    Dim cumulativeInterpolations As Long
    Dim iOverlyCloseBeatsOffset As Long
    Dim iAbberOffset As Long

    iTrialNum = 1
    iColsPerOutput = 16
        
    Do
        cumulativeInterpolations = 0
        iAbberOffset = 0
        iOverlyCloseBeatsOffset = 0
        iOutputNum = 0
    
        If commentWorksheet.Cells(iTrialNum + 1, 1) = "" Then 'go until there are no more trial numbers
            Exit Do
        End If
        
        interpWS.Cells(1, ((iTrialNum - 1) * 5) + 1).Value = "Trial " & iTrialNum
        interpWS.Cells(2, ((iTrialNum - 1) * 5) + 1).Value = "For range"
        interpWS.Cells(2, ((iTrialNum - 1) * 5) + 2).Value = "LC Sample"
        interpWS.Cells(2, ((iTrialNum - 1) * 5) + 3).Value = "LC Time"
        interpWS.Cells(2, ((iTrialNum - 1) * 5) + 4).Value = "Beats"
    
        overbWS.Cells(1, ((iTrialNum - 1) * 5) + 1).Value = "Trial " & iTrialNum
        overbWS.Cells(2, ((iTrialNum - 1) * 5) + 1).Value = "For range"
        overbWS.Cells(2, ((iTrialNum - 1) * 5) + 2).Value = "LC Sample"
        overbWS.Cells(2, ((iTrialNum - 1) * 5) + 3).Value = "LC Time"
        overbWS.Cells(2, ((iTrialNum - 1) * 5) + 4).Value = "Beats"
        
        abberWS.Cells(1, ((iTrialNum - 1) * 5) + 1).Value = "Trial " & iTrialNum
        abberWS.Cells(2, ((iTrialNum - 1) * 5) + 1).Value = "For range"
        abberWS.Cells(2, ((iTrialNum - 1) * 5) + 2).Value = "LC Sample"
        abberWS.Cells(2, ((iTrialNum - 1) * 5) + 3).Value = "LC Time"
        abberWS.Cells(2, ((iTrialNum - 1) * 5) + 4).Value = "Beats"
        
        lPretrialSampStart = commentWorksheet.Cells(iTrialNum + 1, 2)
        lTrialSampStart = commentWorksheet.Cells(iTrialNum + 1, 3)
        lTrialSampEnd = commentWorksheet.Cells(iTrialNum + 1, 4)
        
        thisStartPoint = lTrialSampStart - 168000
        thisEndPoint = lTrialSampStart - 8000
        Call detectHROnSelection(thisStartPoint, thisEndPoint, proportionInterpolated, detectedHR, overlyCloseBeats, interpolations, abberantBeats, longestInterpolation, shortestInterpolation, interpolationDuration, interpolatedBeatsMax, interpolatedBeatsMin, interpolatedBeats, beatCount, iTrialNum, "-84-4s", cumulativeInterpolations, iOverlyCloseBeatsOffset, iAbberOffset)

        cumulativeInterpolations = cumulativeInterpolations + interpolations
        iAbberOffset = iAbberOffset + abberantBeats
        iOverlyCloseBeatsOffset = iOverlyCloseBeatsOffset + overlyCloseBeats

        iOutputNum = iOutputNum + 1

        Worksheets("Output").Range("O" & (iTrialNum + 1)).Value = detectedHR
        If detectedHR = -1 Or proportionInterpolated >= maxPercOfBeatsInt Or longestInterpolation >= maxSingleIntSamples Or interpolatedBeatsMax >= maxSingleIntBeats Then
            Call highlightCell(Worksheets("Output").Range("O" & (iTrialNum + 1)), "Problem")
        Else
            Call highlightCell(Worksheets("Output").Range("O" & (iTrialNum + 1)), "Clear")
        End If
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 1)).Value = "Trial " & iTrialNum
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 2)).Value = thisStartPoint
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 3)).Value = "'" & calculateLCTime(thisStartPoint)
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 4)).Value = thisEndPoint
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 5)).Value = "'" & calculateLCTime(thisEndPoint)
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 6)).Value = detectedHR
        If detectedHR = -1 Or proportionInterpolated >= maxPercOfBeatsInt Or longestInterpolation >= maxSingleIntSamples Or interpolatedBeatsMax >= maxSingleIntBeats Then
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 6)), "Problem")
        Else
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 6)), "Clear")
        End If
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 7)).Value = interpolations
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 8)).Value = interpolatedBeats
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 9)).Value = proportionInterpolated
        If proportionInterpolated >= maxPercOfBeatsInt Then
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 9)), "Problem")
        Else
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 9)), "Clear")
        End If
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 10)).Value = shortestInterpolation
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 11)).Value = longestInterpolation
        If longestInterpolation >= maxSingleIntSamples Then
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 11)), "Problem")
        Else
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 11)), "Clear")
        End If
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 12)).Value = interpolatedBeatsMin
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 13)).Value = interpolatedBeatsMax
        If interpolatedBeatsMax >= maxSingleIntBeats Then
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 13)), "Problem")
        Else
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 13)), "Clear")
        End If
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 14)).Value = overlyCloseBeats
        If overlyCloseBeats > 0 Then
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 14)), "Note")
        Else
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 14)), "Clear")
        End If
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 15)).Value = abberantBeats
        If abberantBeats > 0 Then
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 15)), "Note")
        Else
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 15)), "Clear")
        End If
        
        thisStartPoint = lTrialSampStart - 8000
        thisEndPoint = lTrialSampStart

        Call detectHROnSelection(thisStartPoint, thisEndPoint, proportionInterpolated, detectedHR, overlyCloseBeats, interpolations, abberantBeats, longestInterpolation, shortestInterpolation, interpolationDuration, interpolatedBeatsMax, interpolatedBeatsMin, interpolatedBeats, beatCount, iTrialNum, "-4-0s", cumulativeInterpolations, iOverlyCloseBeatsOffset, iAbberOffset)
        
        cumulativeInterpolations = cumulativeInterpolations + interpolations
        iAbberOffset = iAbberOffset + abberantBeats
        iOverlyCloseBeatsOffset = iOverlyCloseBeatsOffset + overlyCloseBeats
        
        iOutputNum = iOutputNum + 1

        Worksheets("Output").Range("P" & (iTrialNum + 1)).Value = detectedHR
        If detectedHR = -1 Or proportionInterpolated >= maxPercOfBeatsInt Or longestInterpolation >= maxSingleIntSamples Or interpolatedBeatsMax >= maxSingleIntBeats Then
            Call highlightCell(Worksheets("Output").Range("P" & (iTrialNum + 1)), "Problem")
        Else
            Call highlightCell(Worksheets("Output").Range("P" & (iTrialNum + 1)), "Clear")
        End If
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 1)).Value = "Trial " & iTrialNum
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 2)).Value = thisStartPoint
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 3)).Value = "'" & calculateLCTime(thisStartPoint)
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 4)).Value = thisEndPoint
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 5)).Value = "'" & calculateLCTime(thisEndPoint)
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 6)).Value = detectedHR
        If detectedHR = -1 Or proportionInterpolated >= maxPercOfBeatsInt Or longestInterpolation >= maxSingleIntSamples Or interpolatedBeatsMax >= maxSingleIntBeats Then
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 6)), "Problem")
        Else
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 6)), "Clear")
        End If
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 7)).Value = interpolations
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 8)).Value = interpolatedBeats
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 9)).Value = proportionInterpolated
        If proportionInterpolated >= maxPercOfBeatsInt Then
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 9)), "Problem")
        Else
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 9)), "Clear")
        End If
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 10)).Value = shortestInterpolation
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 11)).Value = longestInterpolation
        If longestInterpolation >= maxSingleIntSamples Then
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 11)), "Problem")
        Else
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 11)), "Clear")
        End If
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 12)).Value = interpolatedBeatsMin
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 13)).Value = interpolatedBeatsMax
        If interpolatedBeatsMax >= maxSingleIntBeats Then
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 13)), "Problem")
        Else
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 13)), "Clear")
        End If
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 14)).Value = overlyCloseBeats
        If overlyCloseBeats > 0 Then
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 14)), "Note")
        Else
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 14)), "Clear")
        End If
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 15)).Value = abberantBeats
        If abberantBeats > 0 Then
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 15)), "Note")
        Else
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 15)), "Clear")
        End If


        thisStartPoint = lTrialSampStart + 10000
        thisEndPoint = lTrialSampStart + 18000
        
        Call detectHROnSelection(thisStartPoint, thisEndPoint, proportionInterpolated, detectedHR, overlyCloseBeats, interpolations, abberantBeats, longestInterpolation, shortestInterpolation, interpolationDuration, interpolatedBeatsMax, interpolatedBeatsMin, interpolatedBeats, beatCount, iTrialNum, "-5-9s", cumulativeInterpolations, iOverlyCloseBeatsOffset, iAbberOffset)
        
        iOutputNum = iOutputNum + 1

        Worksheets("Output").Range("Q" & (iTrialNum + 1)).Value = detectedHR
        If detectedHR = -1 Or proportionInterpolated >= maxPercOfBeatsInt Or longestInterpolation >= maxSingleIntSamples Or interpolatedBeatsMax >= maxSingleIntBeats Then
            Call highlightCell(Worksheets("Output").Range("Q" & (iTrialNum + 1)), "Problem")
        Else
            Call highlightCell(Worksheets("Output").Range("Q" & (iTrialNum + 1)), "Clear")
        End If
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 1)).Value = "Trial " & iTrialNum
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 2)).Value = thisStartPoint
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 3)).Value = "'" & calculateLCTime(thisStartPoint)
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 4)).Value = thisEndPoint
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 5)).Value = "'" & calculateLCTime(thisEndPoint)
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 6)).Value = detectedHR
        If detectedHR = -1 Or proportionInterpolated >= maxPercOfBeatsInt Or longestInterpolation >= maxSingleIntSamples Or interpolatedBeatsMax >= maxSingleIntBeats Then
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 6)), "Problem")
        Else
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 6)), "Clear")
        End If
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 7)).Value = interpolations
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 8)).Value = interpolatedBeats
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 9)).Value = proportionInterpolated
        If proportionInterpolated >= maxPercOfBeatsInt Then
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 9)), "Problem")
        Else
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 9)), "Clear")
        End If
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 10)).Value = shortestInterpolation
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 11)).Value = longestInterpolation
        If longestInterpolation >= maxSingleIntSamples Then
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 11)), "Problem")
        Else
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 11)), "Clear")
        End If
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 12)).Value = interpolatedBeatsMin
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 13)).Value = interpolatedBeatsMax
        If interpolatedBeatsMax >= maxSingleIntBeats Then
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 13)), "Problem")
        Else
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 13)), "Clear")
        End If
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 14)).Value = overlyCloseBeats
        If overlyCloseBeats > 0 Then
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 14)), "Note")
        Else
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 14)), "Clear")
        End If
        Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 15)).Value = abberantBeats
        If abberantBeats > 0 Then
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 15)), "Note")
        Else
            Call highlightCell(Worksheets("HR detection").Cells((iTrialNum + 2), (((iOutputNum - 1) * iColsPerOutput) + 15)), "Clear")
        End If
        
        iTrialNum = iTrialNum + 1
    Loop
    
End Sub


Sub detectHROnSelection(lStartPoint As Long, lEndPoint As Long, ByRef proportionInterpolated, ByRef detectedHR, ByRef overlyCloseBeats, ByRef interpolations, ByRef abberantBeats, ByRef longestInterpolation, ByRef shortestInterpolation, ByRef interpolationDuration, ByRef interpolatedBeatsMax, ByRef interpolatedBeatsMin, ByRef interpolatedBeats, ByRef beatCount, iTrialNum As Integer, strRangeTitle As String, iInterpOffset As Long, iOverlyCloseBeatsOffset As Long, iAbberOffset As Long)

    detectedHR = 0
    overlyCloseBeats = 0
    interpolations = 0
    interpolatedBeats = 0
    longestInterpolation = 0
    shortestInterpolation = 0
    interpolatedBeatsMax = 0
    interpolatedBeatsMin = 0
    abberantBeats = 0
    beatCount = 0#
    proportionInterpolated = 0#
    
    Dim returnFailed As Boolean
    Dim isFirstBeat As Boolean
    Dim isLastBeat As Boolean
    
    Dim strInterpolationAddr As String
    
    Dim beatWorksheet As Worksheet
    Set beatWorksheet = Worksheets("Beat points from LabChart")
    
    Dim interpWS As Worksheet
    Set interpWS = Worksheets("Interpolations")
      
    Dim overbWS As Worksheet
    Set overbWS = Worksheets("Overbeats")
    
    Dim abberWS As Worksheet
    Set abberWS = Worksheets("Abberant beats")
    
    Dim beatDuration As Long
    Dim currentBeatOffset As Long 'offset (in columns) from first beat
    Dim currentBeatSamp As Long 'current beat position in time (in samples)
    Dim prevAcceptedBeatSamp As Long 'previous beat position in time (in samples)
    Dim nextBeatSamp As Long 'the beat position of the next beat in samples - only used in interpolation/overly close beat processing

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
    isFirstBeat = True

    prevAcceptedBeatSamp = beatWorksheet.Cells(iTrialNum, lStartColNum).Value 'set the point of the first accepted beat to the starting beat
    currentBeatSamp = beatWorksheet.Cells(iTrialNum, lStartColNum + currentBeatOffset).Value
    
    Do
        If (currentBeatSamp - prevAcceptedBeatSamp) > (maxInterBeatOverrun * beatDuration) Then
            thisInterpolationDuration = (currentBeatSamp - prevAcceptedBeatSamp)
            strInterpolationAddr = beatWorksheet.Cells(iTrialNum, lStartColNum + currentBeatOffset).Address()
            
            'Inter-beat variation is more than what is allowable, so probably missed beats - calculate beat duration after gap for interpolation
            lPostBeatDuration = (beatWorksheet.Cells(iTrialNum, lStartColNum + currentBeatOffset + 1).Value - currentBeatSamp)
            If lPostBeatDuration > ((maxInterBeatOverrun + (maxInterBeatOverrun * 0.1)) * beatDuration) Then 'check if the next beat might also have missed
                'next beat also looks like a miss; check the following beat
                lPostBeatDuration = (beatWorksheet.Cells(iTrialNum, lStartColNum + currentBeatOffset + 2).Value - beatWorksheet.Cells(iTrialNum, lStartColNum + currentBeatOffset + 1))
                If lPostBeatDuration > ((maxInterBeatOverrun + (maxInterBeatOverrun * 0.2)) * beatDuration) Then 'check if the next beat might also have missed. Give a bit more leeway on how much the duration can have changed, as it is more temporally distant
                    'beat after is also a miss. Give up the ghost.
                    returnFailed = True
                End If
            End If
            thisInterpolationBeats = thisInterpolationDuration / ((beatDuration + lPostBeatDuration) / 2) 'calculate the number of beats to interpolate;
            If isFirstBeat Then
                If Round(thisInterpolationBeats) = 1 Then
                    abberantBeats = abberantBeats + 1
                    beatCount = beatCount + 1# * ((currentBeatSamp - lEndPoint) / (currentBeatSamp - prevAcceptedBeatSamp)) 'if first beat, only include a potion of a beat matching the proportion within the set boundaries
                    abberWS.Cells(abberantBeats + iAbberOffset + 2, ((iTrialNum - 1) * 5) + 1).Value = strRangeTitle
                    abberWS.Cells(abberantBeats + iAbberOffset + 2, ((iTrialNum - 1) * 5) + 2).Value = currentBeatSamp
                    abberWS.Cells(abberantBeats + iAbberOffset + 2, ((iTrialNum - 1) * 5) + 3).Value = "'" & calculateLCTime(currentBeatSamp)
                    abberWS.Cells(abberantBeats + iAbberOffset + 2, ((iTrialNum - 1) * 5) + 4).Value = thisInterpolationBeats - 1
                Else
                    interpolations = interpolations + 1
                    If (thisInterpolationBeats - 1) > interpolatedBeatsMax Or interpolations = 1 Then
                        interpolatedBeatsMax = (thisInterpolationBeats - 1)
                    End If
                    If (thisInterpolationBeats - 1) < interpolatedBeatsMin Or interpolations = 1 Then
                        interpolatedBeatsMin = (thisInterpolationBeats - 1)
                    End If
                    
                    interpWS.Cells(interpolations + iInterpOffset + 2, ((iTrialNum - 1) * 5) + 1).Value = strRangeTitle
                    interpWS.Cells(interpolations + iInterpOffset + 2, ((iTrialNum - 1) * 5) + 2).Value = currentBeatSamp
                    interpWS.Cells(interpolations + iInterpOffset + 2, ((iTrialNum - 1) * 5) + 3).Value = "'" & calculateLCTime(currentBeatSamp)
                    interpWS.Cells(interpolations + iInterpOffset + 2, ((iTrialNum - 1) * 5) + 4).Value = thisInterpolationBeats - 1
                    
                    beatCount = beatCount + (CDbl(Round(thisInterpolationBeats)) * ((currentBeatSamp - lStartPoint) / (currentBeatSamp - prevAcceptedBeatSamp))) 'if first beat, only include a potion of a beat matching the proportion within the set boundaries
                    thisInterpolationBeats = CDbl(Round((thisInterpolationBeats - 1))) * ((currentBeatSamp - lStartPoint) / (currentBeatSamp - prevAcceptedBeatSamp)) 'if first beat, only include a potion of a beat matching the proportion within the set boundaries
                End If
                isFirstBeat = False
            ElseIf isLastBeat Then
                If Round(thisInterpolationBeats) = 1 Then
                    abberantBeats = abberantBeats + 1
                    beatCount = beatCount + 1# * (1 - ((currentBeatSamp - lEndPoint) / (currentBeatSamp - prevAcceptedBeatSamp)))  'if first beat, only include a potion of a beat matching the proportion within the set boundaries
                    abberWS.Cells(abberantBeats + iAbberOffset + 2, ((iTrialNum - 1) * 5) + 1).Value = strRangeTitle
                    abberWS.Cells(abberantBeats + iAbberOffset + 2, ((iTrialNum - 1) * 5) + 2).Value = currentBeatSamp
                    abberWS.Cells(abberantBeats + iAbberOffset + 2, ((iTrialNum - 1) * 5) + 3).Value = "'" & calculateLCTime(currentBeatSamp)
                    abberWS.Cells(abberantBeats + iAbberOffset + 2, ((iTrialNum - 1) * 5) + 4).Value = thisInterpolationBeats - 1
                Else
                    interpolations = interpolations + 1
                    If (thisInterpolationBeats - 1) > interpolatedBeatsMax Or interpolations = 1 Then
                        interpolatedBeatsMax = (thisInterpolationBeats - 1)
                    End If
                    If (thisInterpolationBeats - 1) < interpolatedBeatsMin Or interpolations = 1 Then
                        interpolatedBeatsMin = (thisInterpolationBeats - 1)
                    End If
                    
                    interpWS.Cells(interpolations + iInterpOffset + 2, ((iTrialNum - 1) * 5) + 1).Value = strRangeTitle
                    interpWS.Cells(interpolations + iInterpOffset + 2, ((iTrialNum - 1) * 5) + 2).Value = currentBeatSamp
                    interpWS.Cells(interpolations + iInterpOffset + 2, ((iTrialNum - 1) * 5) + 3).Value = "'" & calculateLCTime(currentBeatSamp)
                    interpWS.Cells(interpolations + iInterpOffset + 2, ((iTrialNum - 1) * 5) + 4).Value = thisInterpolationBeats - 1
                
                    beatCount = beatCount + (CDbl(Round(thisInterpolationBeats)) * (1 - ((currentBeatSamp - lEndPoint) / (currentBeatSamp - prevAcceptedBeatSamp)))) 'if first beat, only include a potion of a beat matching the proportion within the set boundaries
                    thisInterpolationBeats = CDbl(Round((thisInterpolationBeats - 1))) * (1 - ((currentBeatSamp - lEndPoint) / (currentBeatSamp - prevAcceptedBeatSamp))) 'if first beat, only include a potion of a beat matching the proportion within the set boundaries
                End If
            Else
                If Round(thisInterpolationBeats) = 1 Then
                    abberantBeats = abberantBeats + 1
                    beatCount = beatCount + 1#
                    abberWS.Cells(abberantBeats + iAbberOffset + 2, ((iTrialNum - 1) * 5) + 1).Value = strRangeTitle
                    abberWS.Cells(abberantBeats + iAbberOffset + 2, ((iTrialNum - 1) * 5) + 2).Value = currentBeatSamp
                    abberWS.Cells(abberantBeats + iAbberOffset + 2, ((iTrialNum - 1) * 5) + 3).Value = "'" & calculateLCTime(currentBeatSamp)
                    abberWS.Cells(abberantBeats + iAbberOffset + 2, ((iTrialNum - 1) * 5) + 4).Value = thisInterpolationBeats - 1
                Else
                    interpolations = interpolations + 1
                    If (thisInterpolationBeats - 1) > interpolatedBeatsMax Or interpolations = 1 Then
                        interpolatedBeatsMax = (thisInterpolationBeats - 1)
                    End If
                    If (thisInterpolationBeats - 1) < interpolatedBeatsMin Or interpolations = 1 Then
                        interpolatedBeatsMin = (thisInterpolationBeats - 1)
                    End If
                    
                    interpWS.Cells(interpolations + iInterpOffset + 2, ((iTrialNum - 1) * 5) + 1).Value = strRangeTitle
                    interpWS.Cells(interpolations + iInterpOffset + 2, ((iTrialNum - 1) * 5) + 2).Value = currentBeatSamp
                    interpWS.Cells(interpolations + iInterpOffset + 2, ((iTrialNum - 1) * 5) + 3).Value = "'" & calculateLCTime(currentBeatSamp)
                    interpWS.Cells(interpolations + iInterpOffset + 2, ((iTrialNum - 1) * 5) + 4).Value = thisInterpolationBeats - 1
                
                    beatCount = beatCount + CDbl(Round((thisInterpolationDuration / ((beatDuration + lPostBeatDuration) / 2)))) 'calculate the number of beats to interpolate;
                    thisInterpolationBeats = CDbl(Round(thisInterpolationBeats - 1)) 'correct the number of actual interpolated beats
                End If
            End If
            
            'beatCount = beatCount + thisInterpolationBeats
            
            'update cumulative information
            If thisInterpolationDuration > longestInterpolation Or interpolations = 1 Then
                longestInterpolation = thisInterpolationDuration
            End If
            If thisInterpolationDuration < shortestInterpolation Or interpolations = 1 Then
                shortestInterpolation = thisInterpolationDuration
            End If
            
            interpolationDuration = interpolationDuration + thisInterpolationDuration
            interpolatedBeats = interpolatedBeats + thisInterpolationBeats
            
            prevAcceptedBeatSamp = currentBeatSamp
            
        ElseIf (currentBeatSamp - prevAcceptedBeatSamp) < (maxInterBeatUnderrun * beatDuration) Then 'is the beat smaller than the minimum variation allowed
            If (nextBeatSamp - prevAcceptedBeatSamp) < (maxInterBeatOverrun * beatDuration) And (nextBeatSamp - prevAcceptedBeatSamp) > (maxInterBeatUnderrun * beatDuration) Then
                'looks like this is an abberant beat that shouldn't be counted
                overlyCloseBeats = overlyCloseBeats + 1
                overbWS.Cells(overlyCloseBeats + iOverlyCloseBeatsOffset + 2, ((iTrialNum - 1) * 5) + 1).Value = strRangeTitle
                overbWS.Cells(overlyCloseBeats + iOverlyCloseBeatsOffset + 2, ((iTrialNum - 1) * 5) + 2).Value = currentBeatSamp
                overbWS.Cells(overlyCloseBeats + iOverlyCloseBeatsOffset + 2, ((iTrialNum - 1) * 5) + 3).Value = "'" & calculateLCTime(currentBeatSamp)
                overbWS.Cells(overlyCloseBeats + iOverlyCloseBeatsOffset + 2, ((iTrialNum - 1) * 5) + 4).Value = ((currentBeatSamp - prevAcceptedBeatSamp) / beatDuration)
            Else
                'maybe a normal beat - count it as normal
                If isFirstBeat Then
                    beatCount = beatCount + ((currentBeatSamp - lStartPoint) / (currentBeatSamp - prevAcceptedBeatSamp))  'if first beat, only include a potion of a beat matching the proportion within the set boundaries
                    isFirstBeat = False
                ElseIf isLastBeat Then
                    beatCount = beatCount + (1 - ((currentBeatSamp - lEndPoint) / (currentBeatSamp - prevAcceptedBeatSamp))) 'if first beat, only include a potion of a beat matching the proportion within the set boundaries
                Else
                    beatCount = beatCount + 1#
                End If

                beatDuration = ((currentBeatSamp - prevAcceptedBeatSamp) + beatDuration) / 2
                prevAcceptedBeatSamp = currentBeatSamp
            End If
        Else
            If isFirstBeat Then
                beatCount = beatCount + ((currentBeatSamp - lStartPoint) / (currentBeatSamp - prevAcceptedBeatSamp))  'if first beat, only include a potion of a beat matching the proportion within the set boundaries
                isFirstBeat = False
            ElseIf isLastBeat Then
                beatCount = beatCount + (1 - ((currentBeatSamp - lEndPoint) / (currentBeatSamp - prevAcceptedBeatSamp)))  'if first beat, only include a potion of a beat matching the proportion within the set boundaries
            Else
                beatCount = beatCount + 1#
            End If
            beatDuration = ((currentBeatSamp - prevAcceptedBeatSamp) + beatDuration) / 2
            prevAcceptedBeatSamp = currentBeatSamp
        End If
        
        'prevAcceptedBeatSamp = currentBeatSamp
        currentBeatOffset = currentBeatOffset + 1
        currentBeatSamp = beatWorksheet.Cells(iTrialNum, lStartColNum + currentBeatOffset).Value
        nextBeatSamp = beatWorksheet.Cells(iTrialNum, lStartColNum + currentBeatOffset + 1).Value
        If currentBeatSamp > lEndPoint Then 'check if we've overrun our endpoint
            If Not isLastBeat Then
                isLastBeat = True
            Else
                Exit Do
            End If
        End If
    Loop
    
    If returnFailed Then
        detectedHR = -1
    Else
        If interpolatedBeats > 0 Then
            proportionInterpolated = (interpolatedBeats / beatCount)
        End If
        detectedHR = beatCount / ((((lEndPoint - lStartPoint) / 2000) / 60))
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
        If (HR(2) / HR(1) > maxInterBeatOverrun) Or (HR(2) / HR(1) < maxInterBeatUnderrun) Or (HR(1) > maxAcceptableHR) Or (HR(1) < minAcceptableHR) Or (HR(2) > maxAcceptableHR) Or (HR(2) < minAcceptableHR) Then
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


Sub highlightCell(theCell As Range, strStyle As String)
    Select Case strStyle
        Case "Problem":
            theCell.Interior.Color = hrProbHighlightCell.Interior.Color
            'theCell.Interior.ColorIndex = hrProbHighlightCell.Interior.ColorIndex
            theCell.Font.Color = hrProbHighlightCell.Font.Color
            'theCell.Font.ColorIndex = hrProbHighlightCell.Font.ColorIndex
        Case "Note":
            theCell.Interior.Color = hrNoteHighlightCell.Interior.Color
            'theCell.Interior.ColorIndex = hrNoteHighlightCell.Interior.ColorIndex
            theCell.Font.Color = hrNoteHighlightCell.Font.Color
            'theCell.Font.ColorIndex = hrNoteHighlightCell.Font.ColorIndex
        Case "Clear":
'            theCell.Interior.Color = hrClearHighlightCell.Interior.Color
            'theCell.Interior.ColorIndex = hrClearHighlightCell.Interior.ColorIndex
'            theCell.Font.Color = hrClearHighlightCell.Font.Color
            'theCell.Font.ColorIndex = hrClearHighlightCell.Font.ColorIndex
            theCell.ClearFormats
            
    End Select
End Sub


