Attribute VB_Name = "ImportFrom"
Attribute VB_Base = "0{7210468A-1FBC-490B-A8F4-04869F32902D}{19D21B76-34CB-4281-8B78-2AE208ECF8C4}"
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
