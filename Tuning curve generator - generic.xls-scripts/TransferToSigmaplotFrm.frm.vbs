Attribute VB_Name = "TransferToSigmaplotFrm"
Attribute VB_Base = "0{E30AE40D-C587-41CF-B991-6246CBA46E99}{7A6A0DFC-F54C-4A66-B137-8944F5245E1A}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub Cancel_Click()
    doImport = False
    Unload Me        'Unloads the UserForm.
End Sub

Private Sub SelectAll_Click()
    Dim i As Integer
    For i = 0 To (HeadingList.ListCount - 1)
        HeadingList.Selected(i) = True
    Next
End Sub

Private Sub DeselectAll_Click()
    Dim i As Integer
    For i = 0 To (HeadingList.ListCount - 1)
        HeadingList.Selected(i) = False
    Next
End Sub

Private Sub TransferButton_Click()
    Set dHeadingsSelected = Nothing
    Set dHeadingsSelected = New Dictionary
    
    Dim i As Integer
    For i = 0 To (HeadingList.ListCount - 1)
        If HeadingList.Selected(i) = True Then
            If Not dHeadingsSelected.Exists(HeadingList.List(i)) Then
                Call dHeadingsSelected.Add(HeadingList.List(i), i)
            End If
        End If
    Next
    
    doImport = True
    Unload Me        'Unloads the UserForm.
End Sub

Private Sub UserForm_Activate()
    Dim vHeadings
    vHeadings = dHeadingList.Keys
    
    Dim i As Integer
    
    For i = 0 To UBound(vHeadings)
        Call HeadingList.AddItem(vHeadings(i), i)
        HeadingList.Selected(i) = True
    Next
End Sub
