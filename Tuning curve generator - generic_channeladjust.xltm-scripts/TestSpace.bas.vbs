Attribute VB_Name = "TestSpace"
Sub testFunction()
    Dim tmpDriveDetParams As DriveDetection
    Set tmpDriveDetParams = New DriveDetection
    
    Call tmpDriveDetParams.readDriveDetection(Worksheets("Settings"), "A27")
    Call tmpDriveDetParams.readDriveDetection(thisWorkbook.Worksheets("Settings"), "A27", outputWorkbook.Worksheets("Settings"))
'    If Not tmpDriveDetParams.readDriveDetection(thisWorkbook.Worksheets("Settings"), "A27", outputWorkbook.Worksheets("Settings")) Then
        'loadConfigParams = False
    'End If
    
    Call tmpDriveDetParams.readDriveDetection(Nothing, "A27")
End Sub
