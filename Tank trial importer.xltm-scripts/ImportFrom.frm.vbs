Attribute VB_Name = "ImportFrom"
Attribute VB_Base = "0{251D5547-60E8-4FC9-9EE8-EA7702A8F073}{B7CCECDC-4C63-45BA-B597-312B8F2D0C81}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Dim objTTX As Object


Private Sub Cancel_Click()
    Unload Me        'Unloads the UserForm.
End Sub

Private Sub ImportButton_Click()
    If BlockSelect1.ActiveBlock <> "" Then
        doImport = True
    
        'set global variables to the selected block information
        theServer = BlockSelect1.UseServer
        theTank = BlockSelect1.UseTank
        theBlock = BlockSelect1.ActiveBlock
        
        Worksheets("Variables (do not edit)").Range("B1").Value = BlockSelect1.UseServer
        Worksheets("Variables (do not edit)").Range("B2").Value = BlockSelect1.UseTank
        Worksheets("Variables (do not edit)").Range("B3").Value = BlockSelect1.ActiveBlock
                
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
    
    Set objTTX = CreateObject("TTank.X") 'establish connection to TDT Tank engine

    'when the form loads, if tanks etc were already selected then re-select them
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
    
    'try to read parameters from the spreadsheet variables
    If theServer = "" Then
        theServer = Worksheets("Variables (do not edit)").Range("B1").Value
        theTank = Worksheets("Variables (do not edit)").Range("B2").Value
        theBlock = Worksheets("Variables (do not edit)").Range("B3").Value
        Select Case testSettings(theServer, theTank, theBlock)
            Case ConnectSuccess:
                TankSelect1.UseServer = theServer
                TankSelect1.ActiveTank = theTank
                BlockSelect1.UseServer = theServer
                BlockSelect1.UseTank = theTank
                BlockSelect1.ActiveBlock = theBlock
                TankSelect1.Refresh
                BlockSelect1.Refresh
            Case BlockSelectFail:
                TankSelect1.UseServer = theServer
                TankSelect1.ActiveTank = theTank
                TankSelect1.Refresh
        End Select
    End If
        
End Sub

'test the connection settings to see if it is possible to connect to the server/tank/block
Function testSettings(ActServer, ActTank, ActBlock)
    
    If objTTX.ConnectServer(ActServer, "Me") <> CLng(1) Then
        testSettings = ServerConnectFail
        Exit Function
    ElseIf objTTX.OpenTank(ActTank, "R") <> CLng(1) Then
        objTTX.ReleaseServer
        testSettings = TankConnectFail
        Exit Function
    ElseIf objTTX.SelectBlock(ActBlock) <> CLng(1) Then
        objTTX.CloseTank
        objTTX.ReleaseServer
        testSettings = BlockConnectFail
    End If
    
End Function
