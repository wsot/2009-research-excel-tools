Attribute VB_Name = "Module1"
Option Explicit
Global doImport
Global theServer, theTank, theBlock

Sub buildTuningCurves()
    ImportFrom.Show
    
    If doImport Then
        Call processImport
    End If
End Sub

Sub processImport()
    Dim lBinWidth As Double
    lBinWidth = Worksheets("Settings").Range("B1").Value
    Dim theWorksheets As Variant
    'Dim chanHistTmp(32) As Long 'used as a temporary store to build a histogram across multiple 'Swep's before outputting the data
    Dim arrHistTmp(31) As Long

    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long

    theWorksheets = buildWorksheetArray()
    
    Dim objttx
    Set objttx = CreateObject("TTank.X")
    
    If objttx.ConnectServer(theServer, "Me") <> CLng(1) Then
        MsgBox ("Connecting to server " & theServer & " failed.")
        Exit Sub
    End If
    
    If objttx.OpenTank(theTank, "R") <> CLng(1) Then
        MsgBox ("Connecting to tank " & theTank & " on server " & theServer & " failed .")
        Call objttx.ReleaseServer
        Exit Sub
    End If
    
    If objttx.SelectBlock(theBlock) <> CLng(1) Then
        MsgBox ("Connecting to block " & theBlock & " in tank " & theTank & " on server " & theServer & " failed.")
        Call objttx.CloseTank
        Call objttx.ReleaseServer
        Exit Sub
    End If
    
    Call objttx.CreateEpocIndexing
    
    Dim freqList As Dictionary
    Dim ampList As Dictionary
    
    Set freqList = New Dictionary
    Set ampList = New Dictionary

    Dim dblStartTime As Double
    Dim dblEndTime As Double
    Dim varReturn As Variant
    Dim varAmp As Variant
    
    
    Do
        i = objttx.ReadEventsV(500, "Frq1", 0, 0, dblStartTime, 0#, "ALL")
        If i = 0 Then
            Exit Do
        End If
        
        varReturn = objttx.ParseEvInfoV(0, i, 0)
        For j = 0 To (i - 1)
            If Not freqList.Exists(varReturn(6, j)) Then
                Call freqList.Add(varReturn(6, j), "")
            End If
            dblStartTime = varReturn(5, j) + (1 / 100000)
            varAmp = objttx.QryEpocAtV("Lev1", varReturn(5, j), 0)
            If Not ampList.Exists(varAmp) Then
                Call ampList.Add(varAmp, "")
            End If
            
        Next
        
        If i < 500 Then
            Exit Do
        End If
    Loop
    
    i = 0
    j = 0
    
'    Dim freqAmpArray()
'    Dim freqAmpArray(freqList.Count - 1, ampList.Count - 1, 32)
    Dim iFreqIndex As Integer
    Dim iAmpIndex As Integer

    Dim vFreqKeys As Variant
    Dim vAmpKeys As Variant
                
        
    vFreqKeys = freqList.Keys
    vAmpKeys = ampList.Keys
        
    Dim varChanData As Variant
    Dim dblSwepStartTime As Double

    For iFreqIndex = 0 To UBound(vFreqKeys)
        For iAmpIndex = 0 To UBound(vAmpKeys)
            Call objttx.SetFilterWithDescEx("Frq1 = " & CStr(vFreqKeys(iFreqIndex)) & " and Lev1 = " & CStr(vAmpKeys(iAmpIndex)))
            varReturn = objttx.GetEpocsExV("Swep", 0)
            For i = 0 To UBound(varReturn, 2)
                dblStartTime = varReturn(2, i)
                dblEndTime = dblStartTime + lBinWidth
                'dblEndTime = varReturn(3, i)
                dblSwepStartTime = dblStartTime
                For j = 1 To 32
                    Do
                        k = objttx.ReadEventsV(500, "CSPK", j, 0, dblStartTime, dblEndTime, "JUSTTIMES")
                        If k = 0 Then
                            Exit Do
                        End If
                        
                        arrHistTmp(j - 1) = CLng(arrHistTmp(j - 1)) + CLng(k)
                        
'                        varChanData = objttx.ParseEvInfoV(0, k, 6)
'                        For l = 0 To (k - 1)
'                            Worksheets.Item("Settings").Cells(iAmpIndex + 3, 1).Value = j
'                            Worksheets.Item("Settings").Cells(iAmpIndex + 3, 2).Value = varChanData(0)
'                        Next
                        
                        If k < 500 Then
                            Exit Do
                        Else
                            varChanData = objttx.ParseEvInfoV(k - 1, 1, 6)
                            dblStartTime = varChanData(0) + (1 / 100000)
                        End If
                    Loop
                    dblStartTime = dblSwepStartTime
                Next
            Next
            For j = 1 To 32
                theWorksheets(j - 1).Cells((UBound(vAmpKeys) + 3) - iAmpIndex, iFreqIndex + 2).Value = arrHistTmp(j - 1)
                arrHistTmp(j - 1) = 0
            Next
        Next
    Next

    Call writeAxes(theWorksheets, vFreqKeys, vAmpKeys)

    Call objttx.CloseTank
    Call objttx.ReleaseServer
End Sub

Function buildWorksheetArray() As Variant
    Dim theWorksheets(31)
   Dim strWsname As String
   Dim intWSNum As Long

    Dim i As Integer

    For i = 1 To Worksheets.Count
        strWsname = Worksheets.Item(i).Name
        If Left(strWsname, 4) = "Site" Then
            If IsNumeric(Right(strWsname, Len(strWsname) - 4)) Then
                intWSNum = CInt(Right(strWsname, Len(strWsname) - 4))
                If intWSNum < 33 And intWSNum > 0 Then
                    Set theWorksheets(intWSNum - 1) = Worksheets.Item(i)
                End If
            End If
        End If
    Next

    For i = 0 To 31
        If IsEmpty(theWorksheets(i)) Then
            If i > 0 Then
                Set theWorksheets(i) = Worksheets.Add(, theWorksheets(i - 1), 1, xlWorksheet)
            Else
                Set theWorksheets(i) = Worksheets.Add(, Worksheets.Item(Worksheets.Count), 1, xlWorksheet)
            End If
            theWorksheets(i).Name = "Site" & CStr(i + 1)
        End If
    Next
    buildWorksheetArray = theWorksheets
End Function

Sub writeAxes(theWorksheets As Variant, colLabels As Variant, rowLabels As Variant)
    Dim i As Long
    Dim j As Long
        
    For i = 0 To UBound(theWorksheets)
        For j = 0 To UBound(rowLabels)
            theWorksheets(i).Cells((UBound(rowLabels) + 3) - j, 1).Value = rowLabels(j)
        Next
        For j = 0 To UBound(colLabels)
            theWorksheets(i).Cells(2, j + 2).Value = colLabels(j)
        Next
    Next

End Sub

Sub deleteWorksheets()
   Dim strWsname As String
   Dim intWSNum As Long

    Dim i As Integer
    
    i = Worksheets.Count
    
    Do
        If i = 0 Then
            Exit Do
        End If
        strWsname = Worksheets.Item(i).Name
        If Left(strWsname, 4) = "Site" Then
            If IsNumeric(Right(strWsname, Len(strWsname) - 4)) Then
                intWSNum = CInt(Right(strWsname, Len(strWsname) - 4))
                If intWSNum < 33 And intWSNum > 0 Then
                    Worksheets.Item(i).Delete
                End If
            End If
        End If
        i = i - 1
    Loop
End Sub
