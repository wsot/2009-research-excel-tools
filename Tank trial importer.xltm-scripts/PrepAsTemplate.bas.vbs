Attribute VB_Name = "PrepAsTemplate"
Sub PrepAsTemplate()
    Worksheets("Beat points from LabChart").UsedRange.Clear
    Worksheets("Trial points from LabChart").Range("A2:D500").Clear
    Worksheets("Output").Range("A2:N173").Clear
    Worksheets("Deadzones").UsedRange.Clear
    Worksheets("Interpolations").UsedRange.Clear
    Worksheets("Overbeats").UsedRange.Clear
    Worksheets("AbberantBeats").UsedRange.Clear
    Worksheets("Output").Range("O2:Q173").Clear
    Worksheets("HR detection").Range("A3:AU82").Address
End Sub
