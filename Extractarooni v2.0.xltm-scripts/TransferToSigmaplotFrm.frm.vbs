Attribute VB_Name = "TransferToSigmaplotFrm"
Attribute VB_Base = "0{6CC7A1D2-8399-40F3-A764-E0B1CD592E9F}{10A466CC-EE6C-43F5-B7C4-51E882207960}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub Cancel_Click()
    doImport = False
    
    Dim iHeadingIndex As Integer
    iHeadingIndex = 2
    Dim i As Integer
    For i = 0 To (HeadingList.ListCount - 1)
        If HeadingList.Selected(i) = True Then
            Worksheets("Variables (do not edit)").Range("J" & iHeadingIndex).Value = HeadingList.List(i)
            iHeadingIndex = iHeadingIndex + 1
        End If
    Next
    
    Worksheets("Variables (do not edit)").Range("J" & iHeadingIndex).Value = ""
    
    Unload Me        'Unloads the UserForm.
End Sub

Private Sub HeadingList_Click()

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
    
    Dim iHeadingIndex As Integer
    iHeadingIndex = 2
    Dim i As Integer
    For i = 0 To (HeadingList.ListCount - 1)
        If HeadingList.Selected(i) = True Then
            If Not dHeadingsSelected.Exists(HeadingList.List(i)) Then
                Call dHeadingsSelected.Add(HeadingList.List(i), i)
                Worksheets("Variables (do not edit)").Range("J" & iHeadingIndex).Value = HeadingList.List(i)
                iHeadingIndex = iHeadingIndex + 1
            End If
        End If
    Next
    
    Worksheets("Variables (do not edit)").Range("J" & iHeadingIndex).Value = ""
    
    doImport = True
    Unload Me        'Unloads the UserForm.
End Sub

Private Sub UserForm_Activate()
    Dim vHeadings
    Dim bAllSelected As Boolean

    Set dHeadingsSelected = Nothing
    Set dHeadingsSelected = New Dictionary

    vHeadings = dHeadingList.Keys
    
    Dim iListIndex As Integer
    iListIndex = 2
    If Worksheets("Variables (do not edit)").Range("J" & CStr(iListIndex)).Value = "" Then
        bAllSelected = True
    Else
        bAllSelected = False
        While Worksheets("Variables (do not edit)").Range("J" & CStr(iListIndex)).Value <> ""
            If Not dHeadingsSelected.Exists(Worksheets("Variables (do not edit)").Range("J" & CStr(iListIndex)).Value) Then
                Call dHeadingsSelected.Add(Worksheets("Variables (do not edit)").Range("J" & CStr(iListIndex)).Value, 1)
            End If
            iListIndex = iListIndex + 1
        Wend
    End If
    
    Dim i As Integer
    
    For i = 0 To UBound(vHeadings)
        Call HeadingList.AddItem(vHeadings(i), i)
        If bAllSelected = True Then
            HeadingList.Selected(i) = True
        Else
            If dHeadingsSelected.Exists(vHeadings(i)) Then
                HeadingList.Selected(i) = True
            End If
        End If
    Next
    
End Sub

Sub getTDTTank(ByRef tankFilename, ByRef blockName)
    tankFilename = ""
    blockName = ""

    docFullName = Doc.FullName

    Dim objFS, objFolder
    Set objFS = CreateObject("Scripting.FileSystemObject")
'   fso.GetParentFolderName(docFullName)
    Set objFolder = objFS.GetFolder(objFS.GetFolder(objFS.GetParentFolderName(docFullName)))

    Dim Subfolders, objSubfolder
    Dim BlockFolders, objBlockFolder
    Dim Files, objFile
    Set Subfolders = objFolder.Subfolders
    For Each objSubfolder In Subfolders
        Set Files = objSubfolder.Files
        For Each objFile In Files
            If objFile.Name = "desktop.ini" Then
                Set BlockFolders = objSubfolder.Subfolders
                For Each objBlockFolder In BlockFolders
                    If objBlockFolder.Name <> "TempBlk" Then
                        blockName = objBlockFolder.Name
                    End If
                Next
                tankFilename = objFS.GetFolder(objFS.GetParentFolderName(docFullName)) & "\" & objSubfolder.Name
                Exit For
            End If
        Next
        If tankFilename <> "" Then
            Exit For
        End If
    Next

    Set objFile = Nothing
    Set objFolder = Nothing
    Set objFS = Nothing

End Sub
