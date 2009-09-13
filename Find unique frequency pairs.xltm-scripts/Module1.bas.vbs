Attribute VB_Name = "Module1"
Sub detectUnique()
    Dim i
    Dim pair As Dictionary
    Set pair = New Dictionary
    Dim val As String
    
    Dim firstVar
    Dim secondVar
    
    i = 2
    While Worksheets("Sheet1").Range("F" & i) <> ""
        firstVar = Worksheets("Sheet1").Range("F" & i).Value
        firstVar = CLng(Left(firstVar, Len(firstVar) - 2))
        secondVar = Worksheets("Sheet1").Range("G" & i).Value
        secondVar = CLng(Left(secondVar, Len(secondVar) - 2))
        If firstVar > secondVar Then
            val = firstVar & "," & secondVar
        Else
            val = secondVar & "," & firstVar
        End If
        
        If Not pair.Exists(val) Then
            Call pair.Add(val, 1)
        Else
            pair(val) = pair(val) + 1
        End If
        i = i + 1
    Wend
    
    Dim theKeys
    theKeys = pair.Keys
    
    For i = 0 To UBound(theKeys)
        firstVar = Left(theKeys(i), InStr(theKeys(i), ",") - 1)
        secondVar = Right(theKeys(i), Len(theKeys(i)) - InStr(theKeys(i), ","))
        
        Worksheets("Sheet1").Range("J" & i + 2).Value = firstVar
        Worksheets("Sheet1").Range("K" & i + 2).Value = secondVar
        
        Worksheets("Sheet1").Range("M" & i + 2).Value = pair(theKeys(i))
    Next
        
    
End Sub
