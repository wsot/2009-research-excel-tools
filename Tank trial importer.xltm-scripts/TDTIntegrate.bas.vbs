Attribute VB_Name = "TDTIntegrate"
Option Explicit
Global doImport
Global theServer, theTank, theBlock

Sub importTrials()
    ImportFrom.Show
    
    If doImport Then
        Call processImport
    End If
End Sub


Sub importTrialsFromLabchart()
    theServer = "Local"
    theTank = Worksheets("Variables (do not edit)").Range("B2").Value
    theBlock = Worksheets("Variables (do not edit)").Range("B3").Value
    
    Call processImport
End Sub

Sub processImport()
    Application.Calculation = xlCalculationManual
    Dim dChannelMappings As Dictionary
    Set dChannelMappings = New Dictionary

    Dim dAtten As Dictionary
    Set dAtten = New Dictionary
    
    Dim dOldAtten As Dictionary
    Set dOldAtten = New Dictionary
    
    Call loadAttenList(dAtten, "Attenuations")
    Call loadAttenList(dOldAtten, "Attenuations (incorrect)")

    Dim strChan As String
    Dim strRef As String
    
    Dim strSChnName As String
    
    Dim i As Long

    Call getMappings(dChannelMappings)
    
'    i = 0
'    Do
'        If Worksheets("Site mappings").Range("A" & (i + 2)).Value <> "" Then
'            If Not dChannelMappings.Exists(Worksheets("Site mappings").Range("A" & (i + 2)).Value) Then
'                Call dChannelMappings.Add(CStr(Worksheets("Site mappings").Range("A" & (i + 2)).Value), Worksheets("Site mappings").Range("B" & (i + 2)).Value)
'            End If
'            i = i + 1
'        Else
'            Exit Do
'        End If
'    Loop

    Dim arrBlocks()
    Dim arrTrials()

    Dim objTTX
    Set objTTX = CreateObject("TTank.X")
    
    If objTTX.ConnectServer(theServer, "Me") <> CLng(1) Then
        MsgBox ("Connecting to server " & theServer & " failed.")
        Exit Sub
    End If
    
    If objTTX.OpenTank(theTank, "R") <> CLng(1) Then
        MsgBox ("Connecting to tank " & theTank & " on server " & theServer & " failed .")
        Call objTTX.ReleaseServer
        Exit Sub
    End If
    
    If objTTX.SelectBlock(theBlock) <> CLng(1) Then
        MsgBox ("Connecting to block " & theBlock & " in tank " & theTank & " on server " & theServer & " failed.")
        Call objTTX.CloseTank
        Call objTTX.ReleaseServer
        Exit Sub
    End If
    
    Dim j As Long
    Dim returnVal As Variant
    Dim dblStartTime As Double
    dblStartTime = 0#
    
    Dim iArrLen As Long
    Dim iArrOffset
    iArrOffset = 0
    
    Do
        'locate the start of blocks
        i = objTTX.ReadEventsV(10000, "BloS", 0, 0, dblStartTime, 0#, "ALL")
        If i = 0 Then
            Exit Do
        End If
        
        returnVal = objTTX.ParseEvInfoV(0, i, 0)
        If (iArrOffset = 0) Then
            ReDim Preserve arrBlocks(i - 1)
        Else
            ReDim Preserve arrBlocks(UBound(arrBlocks) + i)
        End If
        
        'build an array of the block numbers and start times for each block
        'returnVal(6,j) contains the value of the epoch (the block number); returnVal(5,j) contains the offset time in msec
        'arrBlocks contains array: 0=block number, 1=start time of block, 2=start time of next block
        For j = 0 To (i - 1)
            arrBlocks(iArrOffset) = Array(returnVal(6, j), returnVal(5, j), 0#)
            If iArrOffset > 0 Then
                arrBlocks(iArrOffset - 1)(2) = returnVal(5, j) - (1 / 100000)
            End If
            dblStartTime = returnVal(5, j) + (1 / 100000)
            iArrOffset = iArrOffset + 1
        Next
        
        'check if this retrieved all the blocks - if <10000 (the maximum number requested) then all blocks have been retrieved
        If i < 10000 Then
            Exit Do
        End If
    Loop

    iArrOffset = 0
    
    'for each block, locate the start of each trial
    Dim iBlockOffset As Long
    For iBlockOffset = 0 To UBound(arrBlocks)
        dblStartTime = arrBlocks(iBlockOffset)(1)
        Do
            'search for trials between the start of the block, and the start of the next block
            i = objTTX.ReadEventsV(10000, "TriS", 0, 0, dblStartTime, arrBlocks(iBlockOffset)(2), "ALL")
            If i = 0 Then
                Exit Do
            End If
            
            returnVal = objTTX.ParseEvInfoV(0, i, 0)
            
            If iArrOffset = 0 Then
                ReDim Preserve arrTrials(i - 1)
            Else
                ReDim Preserve arrTrials(UBound(arrTrials) + i)
            End If
            
            'build an array of all trials
            'arrTrials contains array with:
            '   0=block number
            '   1=trial number
            '   2=trial start time
            '   3=trial start time+2
            '   4=Trial type (acoustic or electrical) completed further along
            '   5=Stim 1 properties
            '   6=Stim 2 properties
            '   7=Max stim amp/atten stim 1
            '   8=Min stim amp/atten stim 1
            '   9=Mean stim amp/atten stim 1
            '   10=Max stim amp/atten stim 2
            '   11=Min stim amp/atten stim 2
            '   12=Mean stim amp/atten stim 2
            For j = 0 To (i - 1)
                arrTrials(iArrOffset) = Array(arrBlocks(iBlockOffset)(0), returnVal(6, j), returnVal(5, j), returnVal(5, j) + 2, "", "", "", "", 0#, 0#, 0#, 0#, 0#, 0#)
                dblStartTime = returnVal(5, j) + (1 / 100000)
                iArrOffset = iArrOffset + 1
            Next
            
            If i < 10000 Then
                Exit Do
            End If
        Loop
    Next
    
    Dim strStim1Filter As String
    Dim strStim2Filter As String
    
    Call objTTX.CreateEpocIndexing
    
    Dim iTrialOffset As Long
    For iTrialOffset = 0 To UBound(arrTrials)
        dblStartTime = arrTrials(iTrialOffset)(2)

        'find the first 3 acoustic sweeps of the trial (SweS). If >2 sweeps, then currently doing acoustic presentations
        i = objTTX.ReadEventsV(3, "SweS", 0, 0, dblStartTime, arrTrials(iTrialOffset)(3), "ALL")
        If i > 2 Then
            'acoustic presentations - get the first three instances of AFrq, including the one before first alternating. This will be the two different frequencies, and tell us the 'repeated' frequency
            i = objTTX.ReadEventsV(3, "AFrq", 0, 0, dblStartTime - 0.5, arrTrials(iTrialOffset)(3), "ALL")
            If i = 0 Then 'catch and exit on errors - e.g. if manually terminated trials
                Exit For
            End If
            arrTrials(iTrialOffset)(4) = "Acoustic"
            returnVal = objTTX.ParseEvInfoV(0, i, 0)
            
            If returnVal(6, 0) = returnVal(6, 1) Then 'check if first stim is same as leadup (which it will be in more recent versions of the approach)
                'first stim of 'trial' is same as repeated; use the first and second stim of trial as the alternators
                arrTrials(iTrialOffset)(5) = CStr(returnVal(6, 1))
                arrTrials(iTrialOffset)(6) = CStr(returnVal(6, 2))
            Else
                'first stim of 'trial' is NOT same as repeated; use the stim preceding the trial as 'first' and first stim of trial as the one alternated with
                arrTrials(iTrialOffset)(5) = CStr(returnVal(6, 0))
                arrTrials(iTrialOffset)(6) = CStr(returnVal(6, 1))
            End If
            
            strStim1Filter = "TriS = " & arrTrials(iTrialOffset)(1) & " AND AFrq = " & arrTrials(iTrialOffset)(5)
            strStim2Filter = "TriS = " & arrTrials(iTrialOffset)(1) & " AND AFrq = " & arrTrials(iTrialOffset)(6)
            'Obtain the attenuations for each of the 20 presentations

            Call readAcousticAttens(objTTX, arrTrials, iTrialOffset, 1, strStim1Filter, dAtten, dOldAtten)
            Call readAcousticAttens(objTTX, arrTrials, iTrialOffset, 2, strStim2Filter, dAtten, dOldAtten)

            arrTrials(iTrialOffset)(5) = arrTrials(iTrialOffset)(5) & "Hz"
            arrTrials(iTrialOffset)(6) = arrTrials(iTrialOffset)(6) & "Hz"
            arrTrials(iTrialOffset)(7) = CStr(Round(arrTrials(iTrialOffset)(7), 2)) & "dB"
            arrTrials(iTrialOffset)(8) = CStr(Round(arrTrials(iTrialOffset)(8), 2)) & "dB"
            arrTrials(iTrialOffset)(9) = CStr(Round(arrTrials(iTrialOffset)(9), 2)) & "dB"
            arrTrials(iTrialOffset)(10) = CStr(Round(arrTrials(iTrialOffset)(10), 2)) & "dB"
            arrTrials(iTrialOffset)(11) = CStr(Round(arrTrials(iTrialOffset)(11), 2)) & "dB"
            arrTrials(iTrialOffset)(12) = CStr(Round(arrTrials(iTrialOffset)(12), 2)) & "dB"

        Else
            'electrical trial - identify stimulation parameters
            'first two Chan epochs will contain the channels stimulated
            i = objTTX.ReadEventsV(2, "SChn", 0, 0, dblStartTime, arrTrials(iTrialOffset)(3), "ALL")
            If i = 0 Then
                i = objTTX.ReadEventsV(2, "Chan", 0, 0, dblStartTime, arrTrials(iTrialOffset)(3), "ALL")
                strSChnName = "Chan"
                If i = 0 Then 'catch and exit on errors - e.g. if manually terminated trials
                    Exit For
                End If
            Else
                strSChnName = "SChn"
            End If
            arrTrials(iTrialOffset)(4) = "Electrical"
            returnVal = objTTX.ParseEvInfoV(0, i, 0)
            arrTrials(iTrialOffset)(5) = CStr(returnVal(6, 0))
            arrTrials(iTrialOffset)(6) = CStr(returnVal(6, 1))
            
            strStim1Filter = "TriS = " & arrTrials(iTrialOffset)(1) & " AND " & strSChnName & " = " & returnVal(6, 0)
            strStim2Filter = "TriS = " & arrTrials(iTrialOffset)(1) & " AND " & strSChnName & " = " & returnVal(6, 1)
            
            If dChannelMappings.Exists(arrTrials(iTrialOffset)(5)) Then 'process channel mapping
                arrTrials(iTrialOffset)(5) = dChannelMappings(arrTrials(iTrialOffset)(5))
            Else
                arrTrials(iTrialOffset)(5) = CStr(arrTrials(iTrialOffset)(5)) & "*"
            End If
            If dChannelMappings.Exists(arrTrials(iTrialOffset)(6)) Then 'process channel mapping
                arrTrials(iTrialOffset)(6) = dChannelMappings(arrTrials(iTrialOffset)(6))
            Else
                arrTrials(iTrialOffset)(6) = CStr(arrTrials(iTrialOffset)(6)) & "*"
            End If
            
            'first two RefC epochs will contain the reference channels
            i = objTTX.ReadEventsV(2, "RefC", 0, 0, dblStartTime, arrTrials(iTrialOffset)(3), "ALL")
            returnVal = objTTX.ParseEvInfoV(0, i, 0)
            strStim1Filter = strStim1Filter & " AND RefC = " & returnVal(6, 0)
            strStim2Filter = strStim2Filter & " AND RefC = " & returnVal(6, 1)
            'arrTrials(iTrialOffset)(5) = arrTrials(iTrialOffset)(5) & " ref " & CStr(returnVal(6, 0))
            If dChannelMappings.Exists(CStr(returnVal(6, 0))) Then 'process channel mapping
                arrTrials(iTrialOffset)(5) = arrTrials(iTrialOffset)(5) & " ref " & dChannelMappings(CStr(returnVal(6, 0)))
            Else
                arrTrials(iTrialOffset)(5) = arrTrials(iTrialOffset)(5) & " ref " & CStr(returnVal(6, 0)) & "*"
            End If

            'arrTrials(iTrialOffset)(6) = arrTrials(iTrialOffset)(6) & " ref " & CStr(returnVal(6, 1))
            If dChannelMappings.Exists(CStr(returnVal(6, 1))) Then  'process channel mapping
                arrTrials(iTrialOffset)(6) = arrTrials(iTrialOffset)(6) & " ref " & dChannelMappings(CStr(returnVal(6, 1)))
            Else
                arrTrials(iTrialOffset)(6) = arrTrials(iTrialOffset)(6) & " ref " & CStr(returnVal(6, 1)) & "*"
            End If
            
            'first two Freq epochs will contain the stimulation frequency
            i = objTTX.ReadEventsV(2, "Freq", 0, 0, dblStartTime, arrTrials(iTrialOffset)(3), "ALL")
            returnVal = objTTX.ParseEvInfoV(0, i, 0)
            arrTrials(iTrialOffset)(5) = arrTrials(iTrialOffset)(5) & " @ " & CStr(returnVal(6, 0)) & "Hz"
            arrTrials(iTrialOffset)(6) = arrTrials(iTrialOffset)(6) & " @ " & CStr(returnVal(6, 1)) & "Hz"
            strStim1Filter = strStim1Filter & " AND Freq = " & returnVal(6, 0)
            strStim2Filter = strStim2Filter & " AND Freq = " & returnVal(6, 1)
            'Obtain the stimulation current for each of the 20 presentations
'            i = objTTX.ReadEventsV(20, "Curr", 0, 0, dblStartTime, dblStartTime + 10, "ALL")
'            returnVal = objTTX.ParseEvInfoV(0, i, 0)

            Call objTTX.SetFilterWithDescEx(strStim1Filter)
            returnVal = objTTX.GetEpocsExV("Curr", 0)
            For j = 0 To 9 'Could go to UBound(returnVal, 2) but this may inclue a trailing tone at the end of the trial
                'if this is the first presentation, set it as max, min, and avg values
                If j = 0 Then
                    arrTrials(iTrialOffset)(7) = returnVal(0, j) 'set min
                    arrTrials(iTrialOffset)(8) = returnVal(0, j) 'set max
                    arrTrials(iTrialOffset)(9) = returnVal(0, j) 'set mean
                Else
                    If Not j > UBound(returnVal, 2) Then
                        'check if less than current min atten; if so, update value
                        If returnVal(0, j) < arrTrials(iTrialOffset)(7) Then
                            arrTrials(iTrialOffset)(7) = returnVal(0, j)
                        End If
                        'check if more than current max atten; if so, update value
                        If returnVal(0, j) > arrTrials(iTrialOffset)(8) Then
                            arrTrials(iTrialOffset)(8) = returnVal(0, j)
                        End If
                        'calculate mean atten
                        arrTrials(iTrialOffset)(9) = arrTrials(iTrialOffset)(9) + ((returnVal(0, j) - arrTrials(iTrialOffset)(9)) / (j + 1))
                    End If
                End If
            Next
            Call objTTX.ResetFilters
            Call objTTX.SetFilterWithDescEx(strStim2Filter)
            returnVal = objTTX.GetEpocsExV("Curr", 2)
            For j = 0 To 9 'Could go to UBound(returnVal, 2) but this may inclue a trailing tone at the end of the trial
                'if this is the first presentation, set it as max, min, and avg values
                If j = 0 Then
                    arrTrials(iTrialOffset)(10) = returnVal(0, j) 'set min
                    arrTrials(iTrialOffset)(11) = returnVal(0, j) 'set max
                    arrTrials(iTrialOffset)(12) = returnVal(0, j) 'set mean
                Else
                    If Not j > UBound(returnVal, 2) Then
                        'check if less than current min atten; if so, update value
                        If returnVal(0, j) < arrTrials(iTrialOffset)(10) Then
                            arrTrials(iTrialOffset)(10) = returnVal(0, j)
                        End If
                        'check if more than current max atten; if so, update value
                        If returnVal(0, j) > arrTrials(iTrialOffset)(11) Then
                            arrTrials(iTrialOffset)(11) = returnVal(0, j)
                        End If
                        'calculate mean atten
                        arrTrials(iTrialOffset)(12) = arrTrials(iTrialOffset)(12) + ((returnVal(0, j) - arrTrials(iTrialOffset)(12)) / (j + 1))
                    End If
                End If
            Next
            Call objTTX.ResetFilters
            arrTrials(iTrialOffset)(7) = CStr(arrTrials(iTrialOffset)(7)) & "uA"
            arrTrials(iTrialOffset)(8) = CStr(arrTrials(iTrialOffset)(8)) & "uA"
            arrTrials(iTrialOffset)(9) = CStr(Round(arrTrials(iTrialOffset)(9), 2)) & "uA"
            arrTrials(iTrialOffset)(10) = CStr(arrTrials(iTrialOffset)(10)) & "uA"
            arrTrials(iTrialOffset)(11) = CStr(arrTrials(iTrialOffset)(11)) & "uA"
            arrTrials(iTrialOffset)(12) = CStr(Round(arrTrials(iTrialOffset)(12), 2)) & "uA"
        End If
    Next
    
    
    For i = 0 To UBound(arrTrials)
        Worksheets("Output").Range("A" & (i + 2)).Value = arrTrials(i)(0)
        Worksheets("Output").Range("B" & (i + 2)).Value = arrTrials(i)(1)
        Worksheets("Output").Range("C" & (i + 2)).Value = arrTrials(i)(2)
        Worksheets("Output").Range("D" & (i + 2)).Value = arrTrials(i)(2) - arrBlocks(0)(1)
        Worksheets("Output").Range("E" & (i + 2)).Value = arrTrials(i)(4)
        Worksheets("Output").Range("F" & (i + 2)).Value = arrTrials(i)(5)
        Worksheets("Output").Range("G" & (i + 2)).Value = arrTrials(i)(7)
        Worksheets("Output").Range("H" & (i + 2)).Value = arrTrials(i)(8)
        Worksheets("Output").Range("I" & (i + 2)).Value = arrTrials(i)(9)
        Worksheets("Output").Range("J" & (i + 2)).Value = arrTrials(i)(6)
        Worksheets("Output").Range("K" & (i + 2)).Value = arrTrials(i)(10)
        Worksheets("Output").Range("L" & (i + 2)).Value = arrTrials(i)(11)
        Worksheets("Output").Range("M" & (i + 2)).Value = arrTrials(i)(12)
        Worksheets("Output").Range("N" & (i + 2)).Value = "=INT(D" & (i + 2) & "/3600) & "":""&MOD(INT(D" & (i + 2) & "/60),60) &"":""&ROUND(MOD(D" & (i + 2) & ",60),2)"
    Next

    Worksheets("Variables (do not edit)").Range("B4").Value = arrBlocks(0)(1)

    Call objTTX.CloseTank
    Call objTTX.ReleaseServer
End Sub

Sub loadAttenList(dAtten As Dictionary, whichWorksheet As String)
    Dim i As Long
    'Dim lFrq As Long
    'Dim dblAtten As Double
    
    i = 0
    
    While Worksheets(whichWorksheet).Range("A" & (i + 2)).Value <> ""
        Call dAtten.Add(CLng(Worksheets(whichWorksheet).Range("A" & (i + 2)).Value), CDbl(Worksheets(whichWorksheet).Range("B" & (i + 2)).Value))
        i = i + 1
    Wend
End Sub


Sub readAcousticAttens(objTTX, arrTrials, iTrialOffset, iWhichTone As Integer, strStimFilter As String, dAtten, dOldAtten)
            Dim isAtten As Boolean
            Dim returnVal As Variant
            Dim j As Long
            Dim dblAmpl As Double
            
            Dim iFreqOffset As Integer
            Dim iAmpOffsets As Integer
            
            Select Case iWhichTone
                Case 1
                    iFreqOffset = 5
                    iAmpOffsets = 7
                Case 2
                    iFreqOffset = 6
                    iAmpOffsets = 10
            End Select
            
            
            Call objTTX.ResetFilters
            Call objTTX.SetFilterWithDescEx(strStimFilter)

            returnVal = objTTX.GetEpocsExV("Attn", 0)
            If Not IsArray(returnVal) Then
                returnVal = objTTX.GetEpocsExV("Ampl", 0)
                isAtten = False
            Else
                isAtten = True
            End If
            
            For j = 0 To 9 'Could go to UBound(returnVal, 2) but this may inclue a trailing tone at the end of the trial
                If isAtten Then
                    dblAmpl = dAtten(CLng(arrTrials(iTrialOffset)(iFreqOffset))) - returnVal(0, j)
                Else
                    dblAmpl = dAtten(CLng(arrTrials(iTrialOffset)(iFreqOffset))) - (dOldAtten(CLng(arrTrials(iTrialOffset)(iFreqOffset))) - returnVal(0, j))
                End If
                'if this is the first presentation, set it as max, min, and avg values
                If j = 0 Then
                    arrTrials(iTrialOffset)(iAmpOffsets) = dblAmpl 'set min
                    arrTrials(iTrialOffset)(iAmpOffsets + 1) = dblAmpl 'set max
                    arrTrials(iTrialOffset)(iAmpOffsets + 2) = dblAmpl 'set mean
                Else
                    'check if less than current min atten; if so, update value
                    If dblAmpl < arrTrials(iTrialOffset)(iAmpOffsets) Then
                        arrTrials(iTrialOffset)(iAmpOffsets) = dblAmpl
                    End If
                    'check if more than current max atten; if so, update value
                    If dblAmpl > arrTrials(iTrialOffset)(iAmpOffsets + 1) Then
                        arrTrials(iTrialOffset)(iAmpOffsets + 1) = dblAmpl
                    End If
                    'calculate mean atten
                    arrTrials(iTrialOffset)(iAmpOffsets + 2) = arrTrials(iTrialOffset)(iAmpOffsets + 2) + ((dblAmpl - arrTrials(iTrialOffset)(iAmpOffsets + 2)) / (j + 1))
                End If
            Next
            
            Call objTTX.ResetFilters
End Sub

Sub getMappings(ByRef dChannelMappings As Dictionary)

    Dim objFS As FileSystemObject
    Set objFS = New FileSystemObject
    Dim i As Integer
    
    Dim strFilename As String
    
    i = 0
    If objFS.FolderExists(theTank) Then
        strFilename = objFS.GetParentFolderName(theTank) & "\Stim mappings.txt"
        If objFS.FileExists(objFS.GetParentFolderName(theTank) & "\Stim mappings.txt") Then
            Call Worksheets("Site mappings").UsedRange.Clear
            
            Dim objTxt As TextStream
            Set objTxt = objFS.OpenTextFile(strFilename, ForReading, False)
            Dim strLine As String
            Dim vSplitLine As Variant
            
            
            Do
                If objTxt.AtEndOfStream Then
                    Exit Do
                End If
                strLine = objTxt.ReadLine
                vSplitLine = Split(strLine, Chr(9), , vbTextCompare)
                If Not UBound(vSplitLine) = 1 Then
                    Exit Do
                End If
                Worksheets("Site mappings").Range("A" & (i + 1)).Value = vSplitLine(0)
                Worksheets("Site mappings").Range("B" & (i + 1)).Value = vSplitLine(1)
                i = i + 1
            Loop
        End If
    End If
    
    i = 0
    Do
        If Worksheets("Site mappings").Range("A" & (i + 2)).Value <> "" Then
            If Not dChannelMappings.Exists(Worksheets("Site mappings").Range("A" & (i + 2)).Value) Then
                Call dChannelMappings.Add(CStr(Worksheets("Site mappings").Range("A" & (i + 2)).Value), Worksheets("Site mappings").Range("B" & (i + 2)).Value)
            End If
            i = i + 1
        Else
            Exit Do
        End If
    Loop
End Sub



