Attribute VB_Name = "DriveDetection"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
    Option Explicit
    
    Public IgnoreInitialTime As Double 'Initial time to ignore in detecting
    Public Gen_MinSpikesPerEpocInFirstN As Double 'Absolute minimum spikes per epoc in intital time window duration
    Public Gen_FirstNDur As Double 'Duration of initial time window (s)
    Public Onset_BinWidth As Double 'Bin Width (s) for 'onset detection' (First bit compared to subsequent bins)
    Public Onset_ReqMultiple As Double 'Onset spike must be at least x times larger than following spikes for onset spike detection
    Public Onset_NumComparBins As Integer 'Number of subsequent bins to compare onset spike to
    Public Onset_MinSpikesPerEpocInComparBins As Double 'Minimum total spikes (per epoc) in comparison bins (the number of bins specified above)
    Public Diff_Threshold As Double 'During-tone vs outside-tone activity difference threshold (ratio inside/outside) for inclusion as 'driven'
    Public Diff_StimDur As Double 'Tone duration (s)
    Public Diff_ITI As Double 'Inter-tone interval (s; including the duration of the tone)

    Public Function readDriveDetection(srcWorksheet As Worksheet, strStartingAddress As String, Optional dstWorksheet As Variant) As Boolean

        Dim drvDetReadOk As Boolean
        drvDetReadOk = True
        
        Dim rngStartPoint As Range
        Set rngStartPoint = srcWorksheet.Range(strStartingAddress)
        
        Dim rngDestPoint As Range
        If Not IsMissing(dstWorksheet) Then
            Set rngDestPoint = dstWorksheet.Range(strStartingAddress)
        End If
        
        Dim vVal As Variant
        
        vVal = rngStartPoint.Offset(4, 1).Value
        If Not checkDataType(vVal, vbDouble, "", , , rngStartPoint.Offset(4, 1).Address) Then
            If Not IsEmpty(rngDestPoint) Then rngDestPoint.Offset(4, 1).Value = vVal
            IgnoreInitialTime = vVal
        Else
            drvDetReadOk = False
        End If

        vVal = rngStartPoint.Offset(2, 1).Value
        If Not checkDataType(vVal, vbDouble, "", , , rngStartPoint.Offset(2, 1).Address) Then
            If Not IsEmpty(rngDestPoint) Then rngDestPoint.Offset(2, 1).Value = vVal
            Gen_MinSpikesPerEpocInFirstN = vVal
        Else
            drvDetReadOk = False
        End If
        
        vVal = rngStartPoint.Offset(3, 1).Value
        If Not checkDataType(vVal, vbDouble, "", , , rngStartPoint.Offset(3, 1).Address) Then
            If Not IsEmpty(rngDestPoint) Then rngDestPoint.Offset(3, 1).Value = vVal
            Gen_FirstNDur = vVal
        Else
            drvDetReadOk = False
        End If
        
        vVal = rngStartPoint.Offset(7, 1).Value
        If Not checkDataType(vVal, vbDouble, "", , , rngStartPoint.Offset(7, 1).Address) Then
            If Not IsEmpty(rngDestPoint) Then rngDestPoint.Offset(7, 1).Value = vVal
            Onset_BinWidth = vVal
        Else
            drvDetReadOk = False
        End If
        
        vVal = rngStartPoint.Offset(8, 1).Value
        If Not checkDataType(vVal, vbDouble, "", , , rngStartPoint.Offset(8, 1).Address) Then
            If Not IsEmpty(rngDestPoint) Then rngDestPoint.Offset(8, 1).Value = vVal
            Onset_ReqMultiple = vVal
        Else
            drvDetReadOk = False
        End If
        
        vVal = rngStartPoint.Offset(9, 1).Value
        If Not checkDataType(vVal, vbInteger, "", , , rngStartPoint.Offset(9, 1).Address) Then
            If Not IsEmpty(rngDestPoint) Then rngDestPoint.Offset(9, 1).Value = vVal
            Onset_NumComparBins = vVal
        Else
            drvDetReadOk = False
        End If

        vVal = rngStartPoint.Offset(10, 1).Value
        If Not checkDataType(vVal, vbDouble, "", , , rngStartPoint.Offset(10, 1).Address) Then
            If Not IsEmpty(rngDestPoint) Then rngDestPoint.Offset(10, 1).Value = vVal
            Onset_MinSpikesPerEpocInComparBins = vVal
        Else
            drvDetReadOk = False
        End If
        
        vVal = rngStartPoint.Offset(13, 1).Value
        If Not checkDataType(vVal, vbDouble, "", , , rngStartPoint.Offset(13, 1).Address) Then
            If Not IsEmpty(rngDestPoint) Then rngDestPoint.Offset(13, 1).Value = vVal
            Diff_Threshold = vVal
        Else
            drvDetReadOk = False
        End If
        
        vVal = rngStartPoint.Offset(14, 1).Value
        If Not checkDataType(vVal, vbDouble, "", , , rngStartPoint.Offset(14, 1).Address) Then
            If Not IsEmpty(rngDestPoint) Then rngDestPoint.Offset(14, 1).Value = vVal
            Diff_StimDur = vVal
        Else
            drvDetReadOk = False
        End If

        vVal = rngStartPoint.Offset(15, 1).Value
        If Not checkDataType(vVal, vbDouble, "", , , rngStartPoint.Offset(15, 1).Address) Then
            If Not IsEmpty(rngDestPoint) Then rngDestPoint.Offset(15, 1).Value = vVal
            Diff_ITI = vVal
        Else
            drvDetReadOk = False
        End If
        
        readDriveDetection = drvDetReadOk
End Function


