Attribute VB_Name = "clsLinkedListSublist"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
        Public Value
        Public Sublist

        Public Sub SetValueSafe(vNewValue)
                LinkedList_SetSafe Value, vNewValue
        End Sub

        'Get an item at a specific index.
        Public Property Get Item(iIndex)
        
                'Base case.
                If iIndex = 1 Then
                        LinkedList_SetSafe Item, Me.Value
                        Exit Property
                End If

                'Exception case.
                If IsEmpty(Me.Sublist) Then
                        Err.Raise -9, "clsLinkedList.Item", "Subscript out of range"
                End If
                
                'Recursive case.
                LinkedList_SetSafe Item, Me.Sublist.Item(iIndex - 1)
        
        End Property
        
        'Set an item at a specific index - this function handles the case where the item being set is an object.
        Public Property Set Item(iIndex, oItem)
        
                'Base case.
                If iIndex = 1 Then
                        Set Me.Value = oItem
                        Exit Property
                End If

                'Exception case.
                If IsEmpty(Me.Sublist) Then
                        Err.Raise -1, "clsLinkedList.Item", "Subscript out of range"
                End If
                
                'Recursive case
                Set Me.Sublist.Item(iIndex - 1) = oItem
                
        End Property
        
        'Let an item at a specific index - this function handles the case where the item being let is not an object.
        Public Property Let Item(iIndex, vItem)
        
                'Base case.
                If iIndex = 1 Then
                        Me.Value = vItem
                        Exit Property
                End If

                'Exception case.
                If IsEmpty(Me.Sublist) Then
                        Err.Raise -1, "clsLinkedList.Item", "Subscript out of range"
                End If
                
                'Recursive case.
                Me.Sublist.Item(iIndex - 1) = vItem
        
        End Property
        
        'Count the number of items in the list.
        Public Property Get Count()
        
                'Base case.
                If IsEmpty(Sublist) Then
                        Count = 1
                        Exit Property
                End If
                
                'Recursive case.
                Count = 1 + Sublist.Count
        
        End Property

        'Insert an item at a specific index in the list.
        Public Sub InsertAfter(vItem, iIndex)
        
                Dim oNewSublist

                'First base case, allowing an append by specifying an insert at list length + 1.
                If iIndex = 1 And IsEmpty(Me.Sublist) Then
                        Me.Append vItem
                        Exit Sub
                End If

                'Second base case, doing an insert after the specified index item.
                If iIndex = 1 Then
                
                        'Create a new item and insert it after this item but before any subitems.
                        Set oNewSublist = New clsLinkedListSublist
                        oNewSublist.SetValueSafe vItem
                        
                        If Not IsEmpty(Me.Sublist) Then Set oNewSublist.Sublist = Me.Sublist
                        Set Me.Sublist = oNewSublist
                
                        Exit Sub
                        
                End If
                
                'Exception case.
                If IsEmpty(Me.Sublist) Then
                        Err.Raise -1, "clsLinkedList.Insert", "Subscript out of range"
                End If
                
                'Recursive case.
                Me.Sublist.InsertAfter vItem, iIndex - 1
        
        End Sub
        
        'Remove the item following a specific index from the list.
        Public Function RemoveFollowingItem(iIndex)

                'First base case in which the following item is the final item in the list.
                If iIndex = 1 And IsEmpty(Me.Sublist.Sublist) Then
                        LinkedList_SetSafe RemoveFollowingItem, Me.Sublist.Value
                        Me.Sublist = Empty
                        Exit Function
                End If
                
                'Second base case.
                If iIndex = 1 Then
                        LinkedList_SetSafe RemoveFollowingItem, Me.Sublist.Value
                        Set Me.Sublist = Me.Sublist.Sublist             'Omit the next item.
                        Exit Function
                End If

                'Exception case - the next item is the last one, so we can't call 'RemoveFollowingItem' on it.
                If IsEmpty(Me.Sublist.Sublist) Then
                        Err.Raise -1, "clsLinkedList.Remove", "Subscript out of range"
                End If
                
                'Recursive case.
                LinkedList_SetSafe RemoveFollowingItem, Me.Sublist.RemoveFollowingItem(iIndex - 1)

        End Function
        
        'Add an item onto the end of the list.
        Public Sub Append(vItem)
        
                Dim oNewSublist
        
                'Base case.
                If IsEmpty(Me.Sublist) Then
                
                        'Create a new item and append it as a sublist.
                        Set oNewSublist = New clsLinkedListSublist
                        oNewSublist.SetValueSafe vItem
                        Set Me.Sublist = oNewSublist
                        
                        Exit Sub

                End If
                
                'Recursive case.
                Me.Sublist.Append vItem

        End Sub
        
        Public Property Get ToReverseArray()

                Dim avArray, lArraySize

                'Base case.
                If IsEmpty(Me.Sublist) Then
                        ReDim avArray(0)
                        LinkedList_SetArraySafe avArray, 0, Me.Value
                        ToReverseArray = avArray
                        Exit Property
                End If
                
                'Recursive case.
                avArray = Me.Sublist.ToReverseArray             'Get array of remaining values.
                lArraySize = UBound(avArray)
                ReDim Preserve avArray(lArraySize + 1)                          'Resize the array to make room for this value.
                LinkedList_SetArraySafe avArray, lArraySize + 1, Me.Value
                ToReverseArray = avArray
                Exit Property
                        
        End Property



