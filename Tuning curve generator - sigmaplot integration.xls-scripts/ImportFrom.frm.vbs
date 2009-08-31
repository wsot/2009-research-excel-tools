Attribute VB_Name = "ImportFrom"
Attribute VB_Base = "0{58A79F4D-F238-4C8F-9198-4D89468D9005}{A0F39646-6135-4A2B-846C-C384DD17FF2D}"
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

Private Sub ImportButton_Click()
    If BlockSelect1.ActiveBlock <> "" Then
        theServer = BlockSelect1.UseServer
        theTank = BlockSelect1.UseTank
        theBlock = BlockSelect1.ActiveBlock
        doImport = True
        Unload Me
    Else
        MsgBox ("Please select a block to import")
    End If
End Sub


Private Sub TankSelect1_TankChanged(ActTank As String, ActServer As String)
    BlockSelect1.UseServer = ActServer
    BlockSelect1.UseTank = ActTank
    Call BlockSelect1.Refresh
End Sub

Private Sub UserForm_Activate()
    If theServer <> "" Then
        TankSelect1.UseServer = theServer
        BlockSelect1.UseServer = theServer
        If theTank <> "" Then
            TankSelect1.ActiveTank = theTank
            BlockSelect1.UseTank = theTank
            If theBlock <> "" Then
                BlockSelect1.ActiveBlock = theBlock
            End If
            BlockSelect1.Refresh
        End If
        TankSelect1.Refresh
    End If
End Sub

Private Sub BlockSelect1_BlockChanged(ActBlock As String, ActTank As String, ActServer As String)
    
    Const EVTYPE_STRON = &H101
    
    Dim objttx As Object
    Set objttx = CreateObject("TTank.X")
    
    If objttx.ConnectServer(ActServer, "Me") <> CLng(1) Then
        MsgBox ("Connecting to server " & theServer & " failed.")
        Exit Sub
    End If
    
    If objttx.OpenTank(ActTank, "R") <> CLng(1) Then
        MsgBox ("Connecting to tank " & theTank & " on server " & theServer & " failed .")
        Call objttx.ReleaseServer
        Exit Sub
    End If
    
    If objttx.SelectBlock(ActBlock) <> CLng(1) Then
        MsgBox ("Connecting to block " & theBlock & " in tank " & theTank & " on server " & theServer & " failed.")
        Call objttx.CloseTank
        Call objttx.ReleaseServer
        Exit Sub
    End If
       
    Dim arrEventCodes() As Long
    
    arrEventCodes = objttx.GetEventCodes(EVTYPE_STRON)
    Dim i As Integer
    
    Dim sOrigXAxis As String
    Dim sOrigYAxis As String
    Dim vOrigOtherGroupings As Variant
    sOrigXAxis = XAxis.Value
    sOrigYAxis = YAxis.Value
    
    Dim bMatchXAxis As Boolean
    bMatchXAxis = False
    Dim bMatchYAxis As Boolean
    bMatchYAxis = False
    
    Call XAxis.Clear
    Call YAxis.Clear
    Call OtherGroupings.Clear
       
    For i = 0 To UBound(arrEventCodes)
        Call XAxis.AddItem(objttx.CodeToString(arrEventCodes(i)), i)
        If CStr(sOrigXAxis) = "" And objttx.CodeToString(arrEventCodes(i)) = "Frq1" Then
            XAxis.Value = "Frq1"
            bMatchXAxis = True
        ElseIf CStr(objttx.CodeToString(arrEventCodes(i))) = CStr(sOrigXAxis) Then
            XAxis.Value = CStr(sOrigXAxis)
            bMatchXAxis = True
        End If
        Call YAxis.AddItem(objttx.CodeToString(arrEventCodes(i)), i)
        If CStr(sOrigYAxis) = "" And objttx.CodeToString(arrEventCodes(i)) = "Lev1" Then
            YAxis.Value = "Lev1"
            bMatchYAxis = True
        ElseIf CStr(objttx.CodeToString(arrEventCodes(i))) = CStr(sOrigYAxis) Then
            YAxis.Value = CStr(sOrigYAxis)
            bMatchYAxis = True
        End If
        Call OtherGroupings.AddItem(objttx.CodeToString(arrEventCodes(i)), i)
    Next
    
    Call XAxis.AddItem("Channel", i)
    Call YAxis.AddItem("Channel", i)
    Call OtherGroupings.AddItem("Channel", i)

    
    If bMatchXAxis = False Then
        XAxis.Value = XAxis.List(0, 0)
    End If
    If bMatchYAxis = False Then
        YAxis.Value = YAxis.List(0, 0)
    End If

    Call objttx.CloseTank
    Call objttx.ReleaseServer
End Sub

