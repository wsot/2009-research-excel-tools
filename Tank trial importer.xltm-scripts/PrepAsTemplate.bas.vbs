Attribute VB_Name = "PrepAsTemplate"
Sub PrepAsTemplate()
    Application.Calculation = xlCalculationManual
    Worksheets("Beat points from LabChart").UsedRange.Clear
    Worksheets("Trial points from LabChart").Range("A2:D500").Clear
    Worksheets("Output").Range("A2:N173").Clear
    Worksheets("Deadzones").UsedRange.Clear
    Worksheets("Interpolations").UsedRange.Clear
    Worksheets("Overbeats").UsedRange.Clear
    Worksheets("Abberant Beats").UsedRange.Clear
    Worksheets("Output").Range("O2:Q173").Clear
    Worksheets("HR detection").Range("A3:BA82").Clear
    Worksheets("-84-4s HRs").UsedRange.Clear
    Worksheets("-4-0s HRs").UsedRange.Clear
    Worksheets("5-9s HRs").UsedRange.Clear
End Sub
