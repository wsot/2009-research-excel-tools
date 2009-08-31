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
    
    Dim i As Long
    Dim j As Long
    Dim returnVal As Variant
    Dim dblStartTime As Double
    dblStartTime = 0#
    
    Dim iArrLen As Long
    iArrOffset = 0
    
    Do
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
        
        For j = 0 To (i - 1)
            arrBlocks(iArrOffset) = Array(returnVal(6, j), returnVal(5, j), 0#)
            If iArrOffset > 0 Then
                arrBlocks(iArrOffset - 1)(2) = returnVal(5, j) - (1 / 100000)
            End If
            dblStartTime = returnVal(5, j) + (1 / 100000)
            iArrOffset = iArrOffset + 1
        Next
        
        If i < 500 Then
            Exit Do
        End If
    Loop

    iArrOffset = 0
    
    Dim iBlockOffset As Long
    For iBlockOffset = 0 To UBound(arrBlocks)
        dblStartTime = arrBlocks(iBlockOffset)(1)
        Do
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
            
            For j = 0 To (i - 1)
                arrTrials(iArrOffset) = Array(arrBlocks(iBlockOffset)(0), returnVal(6, j), returnVal(5, j), returnVal(5, j) + 2, "", "", "", "")
                dblStartTime = returnVal(5, j) + (1 / 100000)
                iArrOffset = iArrOffset + 1
            Next
            
            If i < 500 Then
                Exit Do
            End If
        Loop
    Next
    
    
    Dim iTrialOffset As Long
    For iTrialOffset = 0 To UBound(arrTrials)
        dblStartTime = arrTrials(iTrialOffset)(2)

        i = objTTX.ReadEventsV(3, "SweS", 0, 0, dblStartTime, arrTrials(iTrialOffset)(3), "ALL")
        If i > 2 Then
            i = objTTX.ReadEventsV(2, "AFrq", 0, 0, dblStartTime, arrTrials(iTrialOffset)(3), "ALL")
            arrTrials(iTrialOffset)(4) = "Acoustic"
            returnVal = objTTX.ParseEvInfoV(0, i, 0)
            arrTrials(iTrialOffset)(5) = CStr(returnVal(6, 0)) & "Hz"
            arrTrials(iTrialOffset)(6) = CStr(returnVal(6, 1)) & "Hz"
        Else
            i = objTTX.ReadEventsV(2, "Chan", 0, 0, dblStartTime, arrTrials(iTrialOffset)(3), "ALL")
            arrTrials(iTrialOffset)(4) = "Electrical"
            returnVal = objTTX.ParseEvInfoV(0, i, 0)
            arrTrials(iTrialOffset)(5) = CStr(returnVal(6, 0))
            arrTrials(iTrialOffset)(6) = CStr(returnVal(6, 1))
            i = objTTX.ReadEventsV(2, "RefC", 0, 0, dblStartTime, arrTrials(iTrialOffset)(3), "ALL")
            returnVal = objTTX.ParseEvInfoV(0, i, 0)
            arrTrials(iTrialOffset)(5) = arrTrials(iTrialOffset)(5) & " ref " & CStr(returnVal(6, 0))
            arrTrials(iTrialOffset)(6) = arrTrials(iTrialOffset)(6) & " ref " & CStr(returnVal(6, 1))
            i = objTTX.ReadEventsV(2, "Freq", 0, 0, dblStartTime, arrTrials(iTrialOffset)(3), "ALL")
            returnVal = objTTX.ParseEvInfoV(0, i, 0)
            arrTrials(iTrialOffset)(5) = arrTrials(iTrialOffset)(5) & " @ " & CStr(returnVal(6, 0)) & "Hz"
            arrTrials(iTrialOffset)(6) = arrTrials(iTrialOffset)(6) & " @ " & CStr(returnVal(6, 1)) & "Hz"
        End If
    Next
    
    
    For i = 0 To UBound(arrTrials)
        Worksheets("Settings").Range("A" & (i + 2)).Value = arrTrials(i)(0)
        Worksheets("Settings").Range("B" & (i + 2)).Value = arrTrials(i)(1)
        Worksheets("Settings").Range("C" & (i + 2)).Value = arrTrials(i)(2)
        Worksheets("Settings").Range("D" & (i + 2)).Value = arrTrials(i)(3)
        Worksheets("Settings").Range("E" & (i + 2)).Value = arrTrials(i)(4)
        Worksheets("Settings").Range("F" & (i + 2)).Value = arrTrials(i)(5)
        Worksheets("Settings").Range("G" & (i + 2)).Value = arrTrials(i)(6)
    Next

    Call objTTX.CloseTank
    Call objTTX.ReleaseServer
End Sub
