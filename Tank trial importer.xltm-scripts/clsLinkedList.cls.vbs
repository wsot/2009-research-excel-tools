Attribute VB_Name = "clsLinkedList"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'<Module:Completion>
Option Explicit


'*******************************************************************
'*
'*   clsLinkedList: Single linked list implementation
'*
'*   Author:            Stephen Bottomley
'*   Last Edit:         17-Apr-2007
'*
'*   Instantiate with:
'*      Dim oList
'*      Set oList = New clsLinkedList
'*
'*******************************************************************
'*
'*   Property Item(iIndex)
'*      Provides access to an item at the specified index.
'*      Get, Set and Let are all supported.
'*
'*   Property Count [read-only]
'*      Returns the number of items in the list.
'*
'*   Sub Insert(vItem, iIndex)
'*      Inserts an item before the specified index.
'*      Specify last index + 1 to append instead.
'*
'*   Sub Append(vItem)
'*      Appends an item to the end of the list.
'*
'*   Function Remove(iIndex)
'*      Removes an item at the specified index and returns
'*      it as the result of the function.
'*
'*   Property ToArray [read-only]
'*      Returns the list as an array. Naive implementation
'*      that is not optimised. Setting array items does not
'*      change items in the list.
'*
'*******************************************************************
'*
'*   Notes:
'*      Implementation is recursive throughout.
'*      List is not optimised, it's a basic sll.
'*      Error handling is by runtime error.
'*
'*******************************************************************
'*   To extend this class, create a private class member as a clsLinkedList
'*   and then add the following functions, where m_oLinkedList is your
'*   private member list:
'*******************************************************************
'Public Default Property Get Item(iIndex)       :       Set Item = m_oLinkedList.Item(iIndex)           :       End Property
'Public Property Set Item(iIndex, oItem)        :       Set m_oLinkedList.Item(iIndex) = oItem          :       End Property
'Public Property Get Count                                      :       Count = m_oLinkedList.Count                                     :       End Property
'Public Sub Insert(oItem, iIndex)                       :       m_oLinkedList.Insert oItem, iIndex                      :       End Sub
'Public Sub Append(oItem)                                       :       m_oLinkedList.Append oItem                                      :       End Sub
'Public Function Remove(iIndex)                         :       Set Remove = m_oLinkedList.Remove(iIndex)       :       End Function
'Public Property Get ToArray                            :       ToArray = m_oLinkedList.ToArray                         :       End Property
'*******************************************************************


        Private m_oSublist

        'Public Default Property Get Item(iIndex)
        Public Property Get Item(iIndex)

                'Exception case.
                If IsEmpty(m_oSublist) Then
                        Err.Raise -9, "clsLinkedList.Item", "Subscript out of range"
                End If

                'Recursive case.
                LinkedList_SetSafe Item, m_oSublist.Item(iIndex)

        End Property
        
        Public Property Set Item(iIndex, oItem)
        
                'Exception case.
                If IsEmpty(m_oSublist) Then
                        Err.Raise -9, "clsLinkedList.Item", "Subscript out of range"
                End If

                'Recursive case.
                Set m_oSublist.Item(iIndex) = oItem
        
        End Property
        
        Public Property Let Item(iIndex, vItem)

                'Exception case.
                If IsEmpty(m_oSublist) Then
                        Err.Raise -9, "clsLinkedList.Item", "Subscript out of range"
                End If
                
                'Recursive case.
                m_oSublist.Item(iIndex) = vItem

        End Property
        
        Public Property Get Count()
        
                If IsEmpty(m_oSublist) Then
                        'Base case.
                        Count = 0
                Else
                        'Recursive case.
                        Count = m_oSublist.Count
                End If
        
        End Property

        Public Sub Insert(vItem, iIndex)
        
                Dim oNewSublist
                
                'First base case in which the user is using an insert at index 1 of an empty collection to insert the first item.
                If IsEmpty(m_oSublist) And iIndex = 1 Then
                        Set m_oSublist = New clsLinkedListSublist
                        m_oSublist.SetValueSafe vItem
                        Exit Sub
                End If
                
                'Second base case in which the user is doing an insert at index 1, before any existing list items.
                If iIndex = 1 Then
                        Set oNewSublist = New clsLinkedListSublist
                        oNewSublist.SetValueSafe vItem
                        Set oNewSublist.Sublist = m_oSublist            'Insert the new item before existing items.
                        Set m_oSublist = oNewSublist                            'The list now starts with the new item.
                        Exit Sub
                End If

                'Exception case.
                If IsEmpty(m_oSublist) Then
                        Err.Raise -9, "clsLinkedList.Insert", "Subscript out of range"
                End If
                
                'Recursive case.
                m_oSublist.InsertAfter vItem, iIndex - 1
                        
        End Sub
        
        Public Function Remove(iIndex)
        
                'Exception case.
                If IsEmpty(m_oSublist) Then
                        Err.Raise -9, "clsLinkedList.Remove", "Subscript out of range"
                End If
                
                'First base case in which the remove creates an empty list.
                If m_oSublist.Count = 1 And iIndex = 1 Then
                        LinkedList_SetSafe Remove, m_oSublist.Value     'Return the removed value.
                        m_oSublist = Empty
                        Exit Function
                End If
                
                'Second base case in which we remove the first item but this does not leave an empty list.
                If iIndex = 1 Then
                        LinkedList_SetSafe Remove, m_oSublist.Value     'Return the removed value.
                        Set m_oSublist = m_oSublist.Sublist
                        Exit Function
                End If

                'Recursive case.
                LinkedList_SetSafe Remove, m_oSublist.RemoveFollowingItem(iIndex - 1)   'Return the removed value.
                        
        End Function
        
        Public Sub Append(vItem)
        
                'Base case in which the user is using an append on an empty collection to insert the first item.
                If IsEmpty(m_oSublist) Then
                        Set m_oSublist = New clsLinkedListSublist
                        m_oSublist.SetValueSafe vItem
                        Exit Sub
                End If
                
                'Recursive case.
                m_oSublist.Append vItem
        
        End Sub
        
        Public Property Get ToArray()
        
                'Base case.
                If IsEmpty(m_oSublist) Then
                        ToArray = Array()
                End If
                
                'Recursive case. The array comes back in reverse (as it was created by recursion) so reverse it.
                ToArray = LinkedList_ReverseArray(m_oSublist.ToReverseArray)
        
        End Property
        

