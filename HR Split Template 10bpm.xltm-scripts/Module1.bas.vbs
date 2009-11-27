Attribute VB_Name = "Module1"
Const maxRow = 166

Sub moveHR()
Attribute moveHR.VB_ProcData.VB_Invoke_Func = "t\n14"
    Application.Calculation = xlCalculationManual
    
    Dim iSrcRow As Integer

    Dim iHRlt280Offset As Integer
    Dim iHR280Offset As Integer
    Dim iHR2900Offset As Integer
    Dim iHR300Offset As Integer
    Dim iHR310Offset As Integer
    Dim iHR320Offset As Integer
    Dim iHR330Offset As Integer
    Dim iHR340Offset As Integer
    Dim iHR350Offset As Integer
    Dim iHR360Offset As Integer
    Dim iHRgt360Offset As Integer
    
    Dim theWS As Worksheet
    Set theWS = ActiveWorkbook.Worksheets("All HR")
    
    iSrcRow = 1
    While iSrcRow < maxRow
        If theWS.Range("FF" & iSrcRow).Value <> "" Then
            If theWS.Cells(iSrcRow, 1).Value < 280 Then
                ActiveWorkbook.Worksheets("<280").Range("A" & (iHRlt280Offset + 2) & ":FF" & (iHRlt280Offset + 2)).Value = theWS.Range("A" & iSrcRow & ":FF" & iSrcRow).Value
                iHRlt280Offset = iHRlt280Offset + 1
            ElseIf theWS.Cells(iSrcRow, 1).Value < 290 Then
                ActiveWorkbook.Worksheets("280-290").Range("A" & (iHR280Offset + 2) & ":FF" & (iHR280Offset + 2)).Value = theWS.Range("A" & iSrcRow & ":FF" & iSrcRow).Value
                iHR280Offset = iHR280Offset + 1
            ElseIf theWS.Cells(iSrcRow, 1).Value < 300 Then
                ActiveWorkbook.Worksheets("290-300").Range("A" & (iHR290Offset + 2) & ":FF" & (iHR290Offset + 2)).Value = theWS.Range("A" & iSrcRow & ":FF" & iSrcRow).Value
                iHR290Offset = iHR290Offset + 1
            ElseIf theWS.Cells(iSrcRow, 1).Value < 310 Then
                ActiveWorkbook.Worksheets("300-310").Range("A" & (iHR300Offset + 2) & ":FF" & (iHR300Offset + 2)).Value = theWS.Range("A" & iSrcRow & ":FF" & iSrcRow).Value
                iHR300Offset = iHR300Offset + 1
            ElseIf theWS.Cells(iSrcRow, 1).Value < 320 Then
                ActiveWorkbook.Worksheets("310-320").Range("A" & (iHR310Offset + 2) & ":FF" & (iHR310Offset + 2)).Value = theWS.Range("A" & iSrcRow & ":FF" & iSrcRow).Value
                iHR310Offset = iHR310Offset + 1
            ElseIf theWS.Cells(iSrcRow, 1).Value < 330 Then
                ActiveWorkbook.Worksheets("320-330").Range("A" & (iHR320Offset + 2) & ":FF" & (iHR320Offset + 2)).Value = theWS.Range("A" & iSrcRow & ":FF" & iSrcRow).Value
                iHR320Offset = iHR320Offset + 1
            ElseIf theWS.Cells(iSrcRow, 1).Value < 340 Then
                ActiveWorkbook.Worksheets("330-340").Range("A" & (iHR330Offset + 2) & ":FF" & (iHR330Offset + 2)).Value = theWS.Range("A" & iSrcRow & ":FF" & iSrcRow).Value
                iHR330Offset = iHR330Offset + 1
            ElseIf theWS.Cells(iSrcRow, 1).Value < 350 Then
                ActiveWorkbook.Worksheets("340-350").Range("A" & (iHR340Offset + 2) & ":FF" & (iHR340Offset + 2)).Value = theWS.Range("A" & iSrcRow & ":FF" & iSrcRow).Value
                iHR340Offset = iHR340Offset + 1
            ElseIf theWS.Cells(iSrcRow, 1).Value < 360 Then
                ActiveWorkbook.Worksheets("350-360").Range("A" & (iHR350Offset + 2) & ":FF" & (iHR350Offset + 2)).Value = theWS.Range("A" & iSrcRow & ":FF" & iSrcRow).Value
                iHR350Offset = iHR350Offset + 1
            Else
                ActiveWorkbook.Worksheets(">360").Range("A" & (iHR360Offset + 2) & ":FF" & (iHR360Offset + 2)).Value = theWS.Range("A" & iSrcRow & ":FF" & iSrcRow).Value
                iHR360Offset = iHR360Offset + 1
            End If
        End If
        iSrcRow = iSrcRow + 1
    Wend
    
    Application.Calculation = xlAutomatic
End Sub
