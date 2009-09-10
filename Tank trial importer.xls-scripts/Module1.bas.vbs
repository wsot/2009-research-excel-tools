Attribute VB_Name = "Module1"
Global doImport
Global theServer, theTank, theBlock

Sub importTrials()
    ImportFrom.Show
    
    If doImport Then
        Call processImport
    End If
End Sub

Sub processImport()
    Dim dChannelMappings As Dictionary
    Set dChannelMappings = New Dictionary

    Dim strChan As String
    Dim strRef As String
    
    Dim strSChnName As String

    Dim i As Long
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
    iArrOffset = 0
    
    Do
        'locate the start of blocks
        i = objTTX.ReadEventsV(500, "BloS", 0, 0, dblStartTime, 0#, "ALL")
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
        
        'check if this retrieved all the blocks - if <500 (the maximum number requested) then all blocks have been retrieved
        If i < 500 Then
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
            i = objTTX.ReadEventsV(500, "TriS", 0, 0, dblStartTime, arrBlocks(iBlockOffset)(2), "ALL")
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
            
            If i < 500 Then
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
            'acoustic presentations - get the first two instances of AFrq, which will be the two different frequencies
            i = objTTX.ReadEventsV(2, "AFrq", 0, 0, dblStartTime, arrTrials(iTrialOffset)(3), "ALL")
            If i = 0 Then 'catch and exit on errors - e.g. if manually terminated trials
                Exit For
            End If
            arrTrials(iTrialOffset)(4) = "Acoustic"
            returnVal = objTTX.ParseEvInfoV(0, i, 0)
            arrTrials(iTrialOffset)(5) = CStr(returnVal(6, 0)) & "Hz"
            arrTrials(iTrialOffset)(6) = CStr(returnVal(6, 1)) & "Hz"
            strStim1Filter = "TriS = " & arrTrials(iTrialOffset)(1) & " AND AFrq = " & returnVal(6, 0)
            strStim2Filter = "TriS = " & arrTrials(iTrialOffset)(1) & " AND AFrq = " & returnVal(6, 1)
            'Obtain the attenuations for each of the 20 presentations
            Call objTTX.ResetFilters
            Call objTTX.SetFilterWithDescEx(strStim1Filter)
            returnVal = objTTX.GetEpocsExV("Attn", 0)
            For j = 0 To 9 'Could go to UBound(returnVal, 2) but this may inclue a trailing tone at the end of the trial
                'if this is the first presentation, set it as max, min, and avg values
                If j = 0 Then
                    arrTrials(iTrialOffset)(7) = returnVal(0, j) 'set min
                    arrTrials(iTrialOffset)(8) = returnVal(0, j) 'set max
                    arrTrials(iTrialOffset)(9) = returnVal(0, j) 'set mean
                Else
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
            Next
            Call objTTX.ResetFilters
            Call objTTX.SetFilterWithDescEx(strStim2Filter)
            returnVal = objTTX.GetEpocsExV("Attn", 0)
            For j = 0 To 9 'Could go to UBound(returnVal, 2) but this may inclue a trailing tone at the end of the trial
                'if this is the first presentation, set it as max, min, and avg values
                If j = 0 Then
                    arrTrials(iTrialOffset)(10) = returnVal(0, j) 'set min
                    arrTrials(iTrialOffset)(11) = returnVal(0, j) 'set max
                    arrTrials(iTrialOffset)(12) = returnVal(0, j) 'set mean
                Else
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
            Next
            Call objTTX.ResetFilters
            arrTrials(iTrialOffset)(7) = CStr(arrTrials(iTrialOffset)(7)) & "dB"
            arrTrials(iTrialOffset)(8) = CStr(arrTrials(iTrialOffset)(8)) & "dB"
            arrTrials(iTrialOffset)(9) = CStr(Round(arrTrials(iTrialOffset)(9), 2)) & "dB"
            arrTrials(iTrialOffset)(10) = CStr(arrTrials(iTrialOffset)(10)) & "dB"
            arrTrials(iTrialOffset)(11) = CStr(arrTrials(iTrialOffset)(11)) & "dB"
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
            
            strStim1Filter = "TriS = " & arrTrials(iTrialOffset)(1) & " AND " & strSChnName & " = " & returnVal(6, 0)
            strStim2Filter = "TriS = " & arrTrials(iTrialOffset)(1) & " AND " & strSChnName & " = " & returnVal(6, 1)
            
            If dChannelMappings.Exists(arrTrials(iTrialOffset)(5)) Then 'process channel mapping
                arrTrials(iTrialOffset)(5) = dChannelMappings(arrTrials(iTrialOffset)(5))
            Else
                arrTrials(iTrialOffset)(5) = arrTrials(iTrialOffset)(5) & "*"
            End If
            'arrTrials(iTrialOffset)(6) = CStr(returnVal(6, 1))
            If dChannelMappings.Exists(CStr(returnVal(6, 1))) Then 'process channel mapping
                arrTrials(iTrialOffset)(6) = dChannelMappings(CStr(returnVal(6, 1)))
            Else
                arrTrials(iTrialOffset)(6) = CStr(CStr(returnVal(6, 1))) & "*"
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
        Worksheets("Settings").Range("A" & (i + 2)).Value = arrTrials(i)(0)
        Worksheets("Settings").Range("B" & (i + 2)).Value = arrTrials(i)(1)
        Worksheets("Settings").Range("C" & (i + 2)).Value = arrTrials(i)(2)
        Worksheets("Settings").Range("D" & (i + 2)).Value = arrTrials(i)(2) - arrBlocks(0)(1)
        Worksheets("Settings").Range("E" & (i + 2)).Value = arrTrials(i)(4)
        Worksheets("Settings").Range("F" & (i + 2)).Value = arrTrials(i)(5)
        Worksheets("Settings").Range("G" & (i + 2)).Value = arrTrials(i)(7)
        Worksheets("Settings").Range("H" & (i + 2)).Value = arrTrials(i)(8)
        Worksheets("Settings").Range("I" & (i + 2)).Value = arrTrials(i)(9)
        Worksheets("Settings").Range("J" & (i + 2)).Value = arrTrials(i)(6)
        Worksheets("Settings").Range("K" & (i + 2)).Value = arrTrials(i)(10)
        Worksheets("Settings").Range("L" & (i + 2)).Value = arrTrials(i)(11)
        Worksheets("Settings").Range("M" & (i + 2)).Value = arrTrials(i)(12)
        Worksheets("Settings").Range("N" & (i + 2)).Value = "=INT(D" & (i + 2) & "/3600) & "":""&MOD(INT(D" & (i + 2) & "/60),60) &"":""&ROUND(MOD(D" & (i + 2) & ",60),2)"
    Next

    Worksheets("Variables (do not edit)").Range("B4").Value = arrBlocks(0)(1)

    Call objTTX.CloseTank
    Call objTTX.ReleaseServer
End Sub
