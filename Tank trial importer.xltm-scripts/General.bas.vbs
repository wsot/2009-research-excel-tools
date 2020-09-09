Attribute VB_Name = "General"
Public Function WorksheetExists(ByVal WorksheetName As String, Optional theWB As Variant) As Boolean
    
    On Error Resume Next
        WorksheetExists = (Worksheets(WorksheetName).Name <> "")
    On Error GoTo 0

End Function



