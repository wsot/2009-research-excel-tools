Attribute VB_Name = "linkedListProcessing"
        
'*** Safely sets a variable to a value regardless of whether the value itself is a variable or an object.
Sub LinkedList_SetSafe(ByRef vVariable, ByRef vValue)
        If VarType(vValue) = vbObject Then
                Set vVariable = vValue
        Else
                vVariable = vValue
        End If
End Sub

'*** Safely sets an array index to a value regardless of whether the value itself is a variable or an object.
Sub LinkedList_SetArraySafe(ByRef avArray, lIndex, ByRef vValue)
        If VarType(vValue) = vbObject Then
                Set avArray(lIndex) = vValue
        Else
                avArray(lIndex) = vValue
        End If
End Sub

'*** Reverses an array.
Function LinkedList_ReverseArray(avArray)

        Dim lArraySize, lItem
        lArraySize = UBound(avArray)
        
        ReDim avReversedArray(lArraySize)
        
        For lItem = 0 To lArraySize
                LinkedList_SetArraySafe avReversedArray, lItem, avArray(lArraySize - lItem)
        Next

        LinkedList_ReverseArray = avReversedArray

End Function


