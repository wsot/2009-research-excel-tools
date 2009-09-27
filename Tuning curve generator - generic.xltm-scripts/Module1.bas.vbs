Attribute VB_Name = "Module1"
Option Explicit
Global doImport
Global theServer, theTank, theBlock
Global xAxisEp, yAxisEp, arrOtherEp
Global lBinWidth As Double
Global lIgnoreFirstMsec As Double
Global iRowOffset As Integer
Global iColOffset As Integer
Global bReverseX, bReverseY As Boolean

Global dHeadingList As Dictionary
Global dHeadingsSelected As Dictionary
Global bXAxisLog As Boolean

Global thisWorkbook As Workbook
Global outputWorkbook As Workbook
Global plotWorkbook As Workbook

Global plotWhichSheet As String

Global bulkImportRootDir As String

Const marginForGoodTuning = 1#

Dim vXAxisKeys As Variant
Dim vYAxisKeys As Variant

Sub bulkBuildTuningCurves()
'        Dim thisWorkbook As Workbook
    Set thisWorkbook = Application.ActiveWorkbook

'    If IsEmpty(theTank) Then
    theServer = thisWorkbook.Worksheets("Variables (do not edit)").Range("B1").Value
'        theTank = thisWorkbook.Worksheets("Variables (do not edit)").Range("B2").Value
'        theBlock = thisWorkbook.Worksheets("Variables (do not edit)").Range("B3").Value
'    End If

    bulkImportRootDir = thisWorkbook.Worksheets("Settings").Range("B21").Value
    If bulkImportRootDir = "" Then
        MsgBox "If bulk importing, a root data folder must be specified"
        Exit Sub
    ElseIf Not checkPathExists(bulkImportRootDir) Then
        MsgBox "The bulk import path does not exist: " & bulkImportRootDir
        Exit Sub
    End If

    BulkImportFrom.Show
    
    If doImport Then
        Dim specifiedOutputDir As String
        Dim outputDir As String
        Dim outputFilename As String
        specifiedOutputDir = thisWorkbook.Worksheets("Settings").Range("B12").Value
        'outputDir = getOutputDir(specifiedOutputDir, theTank)
                        
        Dim templatePath As String
        templatePath = thisWorkbook.Worksheets("Settings").Range("B16").Value
        
        Dim outputFilePrefix As String
        outputFilePrefix = thisWorkbook.Worksheets("Settings").Range("B11").Value
        
        Dim blnAutoclose As Boolean
        blnAutoclose = thisWorkbook.Worksheets("Settings").Range("B10").Value
        
        Dim blnAutosave As Boolean
        If blnAutoclose Then
            blnAutosave = True
        Else
            blnAutosave = thisWorkbook.Worksheets("Settings").Range("B9").Value
        End If
        
        Dim blnAutoPlot As Boolean
        blnAutoPlot = thisWorkbook.Worksheets("Settings").Range("B5").Value
       
        Dim dBlocks As Dictionary
        Set dBlocks = New Dictionary
        Dim i As Integer
        i = 2
        
        While thisWorkbook.Worksheets("Variables (do not edit)").Range("N" & i).Value <> ""
            If Not dBlocks.Exists(thisWorkbook.Worksheets("Variables (do not edit)").Range("N" & i).Value) Then
                Call dBlocks.Add(thisWorkbook.Worksheets("Variables (do not edit)").Range("N" & i).Value, 0)
            End If
            i = i + 1
        Wend
    
        Dim theBlocks As Variant
        theBlocks = dBlocks.Keys
        
        Application.DisplayAlerts = False
        
'        Dim outputWorkbook As Workbook
        
        For i = LBound(theBlocks) To UBound(theBlocks)
            'Call Worksheets("Totals").UsedRange.ClearContents
            'Call Worksheets("StdDev").UsedRange.ClearContents
            'Call Worksheets("Means").UsedRange.ClearContents
            'Call Worksheets("N").UsedRange.ClearContents
            theTank = Left(theBlocks(i), InStr(theBlocks(i), ":") - 1)
            theBlock = Right(theBlocks(i), Len(theBlocks(i)) - Len(theTank) - 1)
            theTank = bulkImportRootDir & "\" & theTank
            
            If i = 0 Then
                templatePath = getTemplateFilename(templatePath, theTank)
            End If
            Set outputWorkbook = Workbooks.Open(templatePath)
            
            If specifiedOutputDir = "" Then
                outputDir = getOutputDir("", theTank)
                outputFilename = outputDir & "\" & outputFilePrefix & theBlock
            Else
                outputDir = getOutputDir(specifiedOutputDir, theTank)
                If outputDir = "" Then
                    MsgBox ("Output directory " & outputDir & " could not be found." & vbCrLf & "Please update the path and try again")
                    Exit Sub
                End If
                outputFilename = outputDir & "\" & Replace(theTank, "\", ".") & "_" & outputFilePrefix & theBlock
            End If
            
            outputWorkbook.Worksheets("Variables (do not edit)").Range("B2").Value = theTank 'update the block on the worksheet
            outputWorkbook.Worksheets("Variables (do not edit)").Range("B3").Value = theBlock 'update the block on the worksheet
            outputWorkbook.Worksheets("Settings").Range("B18").Value = thisWorkbook.Worksheets("Settings").Range("B18").Value
            Call processImport(False)
            Call detectTunedSegments
            If blnAutosave Then
                Call outputWorkbook.SaveAs(outputFilename, 52)
                If blnAutoPlot Then
                    Set plotWorkbook = outputWorkbook
                    Call transferAllToSigmaplot
                End If
                If blnAutoclose Then
                    Call outputWorkbook.Close
                End If
            End If
        Next
        Application.DisplayAlerts = True
    End If
End Sub

Sub buildTuningCurves()
'        Dim thisWorkbook As Workbook
    Set thisWorkbook = Application.ActiveWorkbook

    If IsEmpty(theTank) Then
        theServer = thisWorkbook.Worksheets("Variables (do not edit)").Range("B1").Value
        theTank = thisWorkbook.Worksheets("Variables (do not edit)").Range("B2").Value
        theBlock = thisWorkbook.Worksheets("Variables (do not edit)").Range("B3").Value
    End If

    ImportFrom.Show
    
    If doImport Then
        Dim outputDir As String
        outputDir = thisWorkbook.Worksheets("Settings").Range("B12").Value
        outputDir = getOutputDir(outputDir, theTank)
               
        If outputDir = "" Then
            MsgBox ("Output directory " & outputDir & " could not be found." & vbCrLf & "Please update the path and try again")
            Exit Sub
        End If
                
        Dim templatePath As String
        templatePath = thisWorkbook.Worksheets("Settings").Range("B16").Value
        templatePath = getTemplateFilename(templatePath, theTank)
        
        Dim outputFilePrefix As String
        outputFilePrefix = thisWorkbook.Worksheets("Settings").Range("B11").Value
        
        Dim blnAutoclose As Boolean
        blnAutoclose = thisWorkbook.Worksheets("Settings").Range("B10").Value
        
        Dim blnAutosave As Boolean
        If blnAutoclose Then
            blnAutosave = True
        Else
            blnAutosave = thisWorkbook.Worksheets("Settings").Range("B9").Value
        End If
        
        Dim blnAutoPlot As Boolean
        blnAutoPlot = thisWorkbook.Worksheets("Settings").Range("B5").Value

        
        Dim dBlocks As Dictionary
        Set dBlocks = New Dictionary
        Dim i As Integer
        i = 2
        
        While thisWorkbook.Worksheets("Variables (do not edit)").Range("N" & i).Value <> ""
            If Not dBlocks.Exists(thisWorkbook.Worksheets("Variables (do not edit)").Range("N" & i).Value) Then
                Call dBlocks.Add(thisWorkbook.Worksheets("Variables (do not edit)").Range("N" & i).Value, 0)
            End If
            i = i + 1
        Wend
    
        Dim theBlocks As Variant
        theBlocks = dBlocks.Keys
        
        Application.DisplayAlerts = False
        
'        Dim outputWorkbook As Workbook
        
        For i = LBound(theBlocks) To UBound(theBlocks)
            Set outputWorkbook = Workbooks.Open(templatePath)
            'Call Worksheets("Totals").UsedRange.ClearContents
            'Call Worksheets("StdDev").UsedRange.ClearContents
            'Call Worksheets("Means").UsedRange.ClearContents
            'Call Worksheets("N").UsedRange.ClearContents
            theBlock = theBlocks(i)
            outputWorkbook.Worksheets("Variables (do not edit)").Range("B3").Value = theBlock 'update the block on the worksheet
            outputWorkbook.Worksheets("Settings").Range("B18").Value = thisWorkbook.Worksheets("Settings").Range("B18").Value
            Call processImport(False)
            If blnAutosave Then
                Call outputWorkbook.SaveAs(outputDir & "\" & outputFilePrefix & theBlock, 52)
                If blnAutoPlot Then
                    Set plotWorkbook = outputWorkbook
                    Call transferAllToSigmaplot
                End If
                If blnAutoclose Then
                    Call outputWorkbook.Close
                End If
            End If
        Next
        Application.DisplayAlerts = True
    End If
End Sub


Sub processImport(importIntoSigmaplot As Boolean)

    'load the bin width for histogram generation
    lBinWidth = thisWorkbook.Worksheets("Settings").Range("B1").Value
    outputWorkbook.Worksheets("Settings").Range("B1").Value = lBinWidth
    
    'load the # of msec to ignore at the start (for filtering stimulation artifact
    lIgnoreFirstMsec = thisWorkbook.Worksheets("Settings").Range("B2").Value
    outputWorkbook.Worksheets("Settings").Range("B2").Value = lIgnoreFirstMsec
    
    'write number of channels to output template
    outputWorkbook.Worksheets("Settings").Range("B3").Value = thisWorkbook.Worksheets("Settings").Range("B3").Value
    
    'used to store the maximum histogram peak for normalisation
    Dim lMaxHistHeight As Double
    lMaxHistHeight = 0
    Dim lMaxHistMeanHeight As Double
    lMaxHistMeanHeight = 0
    
    Dim theWorksheets As Variant 'stores the created worksheets to write to
    Dim arrHistTmp() As Long 'used to store the histogram data for each channel as it is generated
    ReDim arrHistTmp(thisWorkbook.Worksheets("Settings").Range("B3").Value - 1)
    'ReDim arrHistTmp(31)
    
    Dim yCount As Long
    Dim xCount As Long
    Dim zOffsetSize As Long
    
    'offsets to leave space at the top and left of the chart
    iRowOffset = 1
    iColOffset = 0

'    theWorksheets = buildWorksheetArray() 'build the worksheets for writing data
    
    'connect to the tank
    Dim objTTX
    Set objTTX = CreateObject("TTank.X")
    
    If objTTX.ConnectServer(theServer, "Me") <> CLng(1) Then
        MsgBox ("Connecting to server " & theServer & " failed.")
        Exit Sub
    End If
    
    If objTTX.OpenTank(theTank, "R") <> CLng(1) Then
        MsgBox ("Connecting to tank " & theTank & " on server " & theServer & " failed .")
        Call objTTX.ReleaseServer
        Exit Sub
    End If
    
    If objTTX.SelectBlock(theBlock) <> CLng(1) Then
        MsgBox ("Connecting to block " & theBlock & " in tank " & theTank & " on server " & theServer & " failed.")
        Call objTTX.CloseTank
        Call objTTX.ReleaseServer
        Exit Sub
    End If
    
    'index epochs - required to use filters
    Call objTTX.CreateEpocIndexing
    
    Dim dblStartTime As Double
    Dim dblEndTime As Double
    
    Dim varReturn As Variant
    
'    Dim vXAxisKeys As Variant
'    Dim vYAxisKeys As Variant
    
    vXAxisKeys = BuildEpocList(objTTX, xAxisEp, bReverseX)
    vYAxisKeys = BuildEpocList(objTTX, yAxisEp, bReverseY)
        
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    
    Dim arrOtherEpocKeys() As Variant
    If UBound(arrOtherEp) <> -1 Then
        ReDim arrOtherEpocKeys(UBound(arrOtherEp))
        
        For i = 0 To UBound(arrOtherEp)
            arrOtherEpocKeys(i) = BuildEpocList(objTTX, arrOtherEp(i), False)
        Next
    End If
    
    i = 0
    j = 0
    
    Dim iXAxisIndex As Integer
    Dim iYAxisIndex As Integer
    Dim arrOtherEpocIndex() As Integer
    If UBound(arrOtherEp) <> -1 Then
        ReDim arrOtherEpocIndex(UBound(arrOtherEp))
    End If
        
    Dim varChanData As Variant
    Dim dblSwepStartTime As Double
    
    Dim xAxisSearchString As String
    Dim yAxisSearchString As String
    Dim otherAxisSearchString() As String
    Dim strOtherAxisSearchString As String
    If UBound(arrOtherEp) <> -1 Then
        ReDim otherAxisSearchString(UBound(arrOtherEp))
    End If

    Dim iChanNum As Integer
    iChanNum = 0

    If UBound(arrOtherEp) <> -1 Then
        For i = 0 To UBound(vXAxisKeys)
            If xAxisEp = "Channel" Then
                iChanNum = vXAxisKeys(i)
                xAxisSearchString = ""
            Else
                xAxisSearchString = xAxisEp & " = " & CStr(vXAxisKeys(i)) & " and "
            End If
            For j = 0 To UBound(vYAxisKeys)
                If yAxisEp = "Channel" Then
                    iChanNum = vYAxisKeys(j)
                    yAxisSearchString = ""
                Else
                    yAxisSearchString = yAxisEp & " = " & CStr(vYAxisKeys(j)) & " and "
                End If
                Call processSearch(objTTX, arrOtherEp, arrOtherEpocKeys, 0, xAxisSearchString & yAxisSearchString, i + 1, j + 1, UBound(vYAxisKeys) + 3, iChanNum, "", xCount, yCount, zOffsetSize, lMaxHistHeight, lMaxHistMeanHeight)
            Next
        Next
    End If

'    Call writeAxes(theWorksheets, vXAxisKeys, vYAxisKeys, iColOffset, iRowOffset)

    Call objTTX.CloseTank
    Call objTTX.ReleaseServer
    
    outputWorkbook.Worksheets("Variables (do not edit)").Range("H1").Value = xCount
    outputWorkbook.Worksheets("Variables (do not edit)").Range("H2").Value = yCount
    outputWorkbook.Worksheets("Variables (do not edit)").Range("H3").Value = zOffsetSize
    outputWorkbook.Worksheets("Variables (do not edit)").Range("H4").Value = lMaxHistHeight
    outputWorkbook.Worksheets("Variables (do not edit)").Range("H5").Value = iColOffset
    outputWorkbook.Worksheets("Variables (do not edit)").Range("H6").Value = iRowOffset
    outputWorkbook.Worksheets("Variables (do not edit)").Range("H7").Value = lMaxHistMeanHeight
    
    'If importIntoSigmaplot Then
        'Call transferToSigmaplot(xCount, yCount, zOffsetSize, iColOffset, iRowOffset, lMaxHistHeight)
    'End If
    
End Sub

Sub writeAxes(colLabels As Variant, rowLabels As Variant, iColOffset, iRowOffset, zOffset)
    Dim j As Long
        
    For j = 0 To UBound(rowLabels)
        outputWorkbook.Worksheets("Totals").Cells(iRowOffset + j + 2 + zOffset, iColOffset + 1).Value = rowLabels(j)
        outputWorkbook.Worksheets("StdDev").Cells(iRowOffset + j + 2 + zOffset, iColOffset + 1).Value = rowLabels(j)
        outputWorkbook.Worksheets("Means").Cells(iRowOffset + j + 2 + zOffset, iColOffset + 1).Value = rowLabels(j)
        outputWorkbook.Worksheets("N").Cells(iRowOffset + j + 2 + zOffset, iColOffset + 1).Value = rowLabels(j)
    Next
    For j = 0 To UBound(colLabels)
        outputWorkbook.Worksheets("Totals").Cells(iRowOffset + zOffset + 1, j + 2).Value = colLabels(j)
        outputWorkbook.Worksheets("StdDev").Cells(iRowOffset + zOffset + 1, j + 2).Value = colLabels(j)
        outputWorkbook.Worksheets("Means").Cells(iRowOffset + zOffset + 1, j + 2).Value = colLabels(j)
        outputWorkbook.Worksheets("N").Cells(iRowOffset + zOffset + 1, j + 2).Value = colLabels(j)
    Next

End Sub


Sub TransferToSigmaplot()

    Dim xCount As Long
    Dim yCount As Long
    Dim zOffsetSize As Long
    Dim lMaxHistHeight As Long
    Dim iColOffset As Integer
    Dim iRowOffset As Integer

    xCount = plotWorkbook.Worksheets("Variables (do not edit)").Range("H1").Value
    yCount = plotWorkbook.Worksheets("Variables (do not edit)").Range("H2").Value
    zOffsetSize = plotWorkbook.Worksheets("Variables (do not edit)").Range("H3").Value
    If plotWhichSheet = "Means" Then
        lMaxHistHeight = plotWorkbook.Worksheets("Variables (do not edit)").Range("H7").Value
    Else
        lMaxHistHeight = plotWorkbook.Worksheets("Variables (do not edit)").Range("H4").Value
    End If
    iColOffset = plotWorkbook.Worksheets("Variables (do not edit)").Range("H5").Value
    iRowOffset = plotWorkbook.Worksheets("Variables (do not edit)").Range("H6").Value
    
    Const SAA_FROMVAL = &H406
    Const SAA_TOVAL = &H407
    Const GPM_SETPLOTATTR = &H301
    Const GPM_SETAXISATTR = 1025
    Const SLA_SELECTDIM = 776
    Const SAA_OPTIONS = 1027
    Const GPM_SETAXISATTRSTRING = 1032
    Const SEA_COLORCOL = 1557
    Const SEA_COLORREPEAT = 1555
    Const SLA_CONTOURFILLTYPE = 856
    Const SAA_SELECTLINE = 1034
    Const SEA_THICKNESS = 1537
    Const SEA_COLOR = 1542
    Const SAA_SUB1OPTIONS = 1040

    Dim SPApp As Object
    Set SPApp = CreateObject("SigmaPlot.Application.1")
    SPApp.Visible = True
    Call SPApp.Application.Notebooks.Add
    
    Dim i As Integer
    Dim j As Long
    Dim k As Long

    Dim xPos As Long
    Dim yPos As Long

    Dim spNB As Object
    Dim spWS As Object
    Dim spDT As Object
    Dim spGRPH As Object
    
    Dim strTitle
    
    xPos = iColOffset + 1
    yPos = iRowOffset
    
    Do
        If plotWorkbook.Worksheets(plotWhichSheet).Cells(yPos, xPos).Value <> "" Then
            If dHeadingsSelected.Exists(plotWorkbook.Worksheets(plotWhichSheet).Cells(yPos, xPos).Value) Then
                strTitle = plotWorkbook.Worksheets(plotWhichSheet).Cells(yPos, xPos).Value
                Set spNB = SPApp.Notebooks.Item(SPApp.Notebooks.Count - 1)
                Set spWS = spNB.NotebookItems.Item(spNB.NotebookItems.Count - 1)
                spWS.Name = plotWorkbook.Worksheets(plotWhichSheet).Cells(yPos, xPos).Value
                Set spDT = spWS.DataTable
                
                yPos = yPos
                            
                For j = 0 To (xCount - 1)
                    spDT.Cell(0, j) = plotWorkbook.Worksheets(plotWhichSheet).Cells(yPos + 1, xPos + j + 1).Value
                Next
                
                For j = 0 To (yCount - 1)
                    spDT.Cell(1, j) = plotWorkbook.Worksheets(plotWhichSheet).Cells(yPos + j + 2, xPos).Value
                Next
                
                
                For j = 0 To (xCount - 1)
                    For k = 0 To (yCount - 1)
                        spDT.Cell(3 + k, j) = plotWorkbook.Worksheets(plotWhichSheet).Cells(yPos + k + 2, xPos + j + 1).Value
                    Next
                Next
                
                spDT.Cell(2, 0) = "@rgb(255,255,255)"
                spDT.Cell(2, 1) = "@rgb(0,0,255)"
                spDT.Cell(2, 2) = "@rgb(0,255,255)"
                spDT.Cell(2, 3) = "@rgb(0,255,0)"
                spDT.Cell(2, 4) = "@rgb(255,255,0)"
                spDT.Cell(2, 5) = "@rgb(255,0,0)"
                    
                'Call spNB.NotebookItems.Add(2)
                'Set spGRPH = spNB.NotebookItems.Item(spNB.NotebookItems.Count - 1)
        
                Call SPApp.ActiveDocument.NotebookItems.Add(2)
                Dim ColumnsPerPlot(2, 3)
                ColumnsPerPlot(0, 0) = 0
                ColumnsPerPlot(1, 0) = 0
                ColumnsPerPlot(2, 0) = 31999999
                ColumnsPerPlot(0, 1) = 1
                ColumnsPerPlot(1, 1) = 0
                ColumnsPerPlot(2, 1) = 31999999
                ColumnsPerPlot(0, 2) = 3
                ColumnsPerPlot(1, 2) = 0
                ColumnsPerPlot(2, 2) = 31999999
                ColumnsPerPlot(0, 3) = 3 + (yCount - 1)
                ColumnsPerPlot(1, 3) = 0
                ColumnsPerPlot(2, 3) = 31999999
                
                Dim PlotColumnCountArray()
                ReDim PlotColumnCountArray(0)
                
                PlotColumnCountArray(0) = 4
                Call SPApp.ActiveDocument.CurrentPageItem.CreateWizardGraph("Contour Plot", "Filled Contour Plot", "XY Many Z", ColumnsPerPlot, PlotColumnCountArray, "Worksheet Columns", "Standard Deviation", "Degrees", 0#, 360#, , "Standard Deviation", True)
                Call SPApp.ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).SelectObject
            
                SPApp.ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Name = strTitle
                SPApp.ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Axes(0).Name = "Attenuation"
                SPApp.ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Axes(1).Name = "Frequency"
                
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, 3)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_OPTIONS, 10)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_OPTIONS, 51380225)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_OPTIONS, 12583938)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTRSTRING, SAA_FROMVAL, "0")
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTRSTRING, SAA_TOVAL, CStr(lMaxHistHeight))
                
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, 3)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, 2)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_THICKNESS, 10)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, 4)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_COLOR, &HFFFFFF)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_COLORCOL, 2)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_COLORREPEAT, 4)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_CONTOURFILLTYPE, 1)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, 5)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_COLOR, &HFFFFFF)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_COLORCOL, 2)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_COLORREPEAT, 4)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, 4)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SUB1OPTIONS, 1298)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SUB1OPTIONS, 3889)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, 1)
        
                Call SPApp.ActiveDocument.NotebookItems.Add(2)
                ColumnsPerPlot(0, 0) = 0
                ColumnsPerPlot(1, 0) = 0
                ColumnsPerPlot(2, 0) = 31999999
                ColumnsPerPlot(0, 1) = 1
                ColumnsPerPlot(1, 1) = 0
                ColumnsPerPlot(2, 1) = 31999999
                ColumnsPerPlot(0, 2) = 3
                ColumnsPerPlot(1, 2) = 0
                ColumnsPerPlot(2, 2) = 31999999
                ColumnsPerPlot(0, 3) = 3 + (yCount - 1)
                ColumnsPerPlot(1, 3) = 0
                ColumnsPerPlot(2, 3) = 31999999
                
                ReDim PlotColumnCountArray(0)
                
                PlotColumnCountArray(0) = 4
                Call SPApp.ActiveDocument.CurrentPageItem.CreateWizardGraph("Contour Plot", "Filled Contour Plot", "XY Many Z", ColumnsPerPlot, PlotColumnCountArray, "Worksheet Columns", "Standard Deviation", "Degrees", 0#, 360#, , "Standard Deviation", True)
                Call SPApp.ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).SelectObject
            
                SPApp.ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Name = "Site y"
                SPApp.ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Axes(0).Name = "Attenuation"
                SPApp.ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Axes(1).Name = "Frequency"
                        
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, 3)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, 2)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_THICKNESS, 10)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, 4)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_COLOR, &HFFFFFF)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_COLORCOL, 2)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_COLORREPEAT, 4)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_CONTOURFILLTYPE, 1)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, 5)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_COLOR, &HFFFFFF)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_COLORCOL, 2)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_COLORREPEAT, 4)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, 4)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SUB1OPTIONS, 1298)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SUB1OPTIONS, 3889)
                Call SPApp.ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, 1)
        
                Call spNB.NotebookItems.Add(1)
            End If
            yPos = yPos + zOffsetSize
        Else
            Exit Do
        End If
    Loop
End Sub


Function BuildEpocList(objTTX, AxisEp, bReverseOrder)
    'build list of epocs for the given axis epoc name
    
    Dim AxisList As Dictionary
    Set AxisList = New Dictionary
    
    Dim dblStartTime As Double
    Dim varReturn As Variant
    
    Dim i As Integer
    Dim j As Integer
    
    If AxisEp = "Channel" Then
        For i = 1 To thisWorkbook.Worksheets("Settings").Range("B3").Value
            Call AxisList.Add(i, 0)
        Next
    Else
        Do
            i = objTTX.ReadEventsV(500, AxisEp, 0, 0, dblStartTime, 0#, "ALL")
            If i = 0 Then
                Exit Do
            End If
            
            varReturn = objTTX.ParseEvInfoV(0, i, 0)
            For j = 0 To (i - 1)
                If Not AxisList.Exists(varReturn(6, j)) Then
                    Call AxisList.Add(varReturn(6, j), "")
                End If
                dblStartTime = varReturn(5, j) + (1 / 100000)
            Next
            
            If i < 500 Then
                Exit Do
            End If
        Loop
    End If
    
    
    
    If bReverseOrder Then
        Dim returnArr()
        Dim tempArr As Variant
        tempArr = AxisList.Keys
        ReDim returnArr(UBound(tempArr))

        For i = 0 To UBound(tempArr)
            returnArr(i) = tempArr(UBound(tempArr) - i)
        Next
        BuildEpocList = returnArr
    Else
        BuildEpocList = AxisList.Keys
    End If

End Function


Function processSearch(ByRef objTTX, ByRef arrOtherEp, ByRef arrOtherEpocKeys, iOtherEpocNum, strSearchString As String, xOffset, yOffset, zOffset, iChanNum, strTitle, ByRef xCount, ByRef yCount, ByRef zOffsetSize, ByRef lMaxHistHeight, ByRef lMaxHistMeanHeight)
    Dim i As Integer
    Dim j As Integer
    Dim strAddedSearchString As String
    Dim strFilter As String
    Dim strAddedTitle As String
    
    For i = 0 To UBound(arrOtherEpocKeys(iOtherEpocNum))
        If arrOtherEp(iOtherEpocNum) <> "Channel" Then
            'add to search string
            strAddedSearchString = strSearchString & arrOtherEp(iOtherEpocNum) & " = " & CStr(arrOtherEpocKeys(iOtherEpocNum)(i)) & " and "
            strAddedTitle = strTitle & arrOtherEp(iOtherEpocNum) & " = " & CStr(arrOtherEpocKeys(iOtherEpocNum)(i)) & ", "
        Else
            strAddedSearchString = strSearchString
            strAddedTitle = strTitle & "Channel = " & CStr(arrOtherEpocKeys(iOtherEpocNum)(i)) & ", "
            iChanNum = arrOtherEpocKeys(iOtherEpocNum)(i)
        End If
        If iOtherEpocNum < UBound(arrOtherEp) Then
            'there are still more epocs to add to the search
            Call processSearch(objTTX, arrOtherEp, arrOtherEpocKeys, iOtherEpocNum + 1, strAddedSearchString, xOffset, yOffset, (zOffset * UBound(arrOtherEpocKeys(iOtherEpocNum))) + i, iChanNum, strAddedTitle, xCount, yCount, zOffsetSize, lMaxHistHeight, lMaxHistMeanHeight)
        Else
            'we have reached the end of the list of epocs - can actually do a search now
            If Right(strAddedSearchString, 5) = " and " Then 'this should always be the case - should be a trailing 'and' to remove
                strFilter = Left(strAddedSearchString, Len(strAddedSearchString) - 5)
            Else
                strFilter = strAddedSearchString
            End If
            Call objTTX.SetFilterWithDescEx(strFilter)
            
            If xOffset = 1 And yOffset = 1 Then
                outputWorkbook.Worksheets("Totals").Cells(iRowOffset + (i * zOffset), iColOffset + 1).Value = Left(strAddedTitle, Len(strAddedTitle) - 2)
                outputWorkbook.Worksheets("N").Cells(iRowOffset + (i * zOffset), iColOffset + 1).Value = Left(strAddedTitle, Len(strAddedTitle) - 2)
                outputWorkbook.Worksheets("Means").Cells(iRowOffset + (i * zOffset), iColOffset + 1).Value = Left(strAddedTitle, Len(strAddedTitle) - 2)
                outputWorkbook.Worksheets("StdDev").Cells(iRowOffset + (i * zOffset), iColOffset + 1).Value = Left(strAddedTitle, Len(strAddedTitle) - 2)
                Call writeAxes(vXAxisKeys, vYAxisKeys, iColOffset, iRowOffset, (i * zOffset))
            End If

            Call writeResults(objTTX, xOffset, yOffset, i * zOffset, iChanNum, lMaxHistHeight, lMaxHistMeanHeight)
            If xOffset > xCount Then
                xCount = xOffset
            End If
            If yOffset > yCount Then
                yCount = yOffset
            End If
            zOffsetSize = zOffset
        End If
    Next

End Function

Sub writeResults(ByRef objTTX, xOffset, yOffset, zOffset, iChanNum, ByRef lMaxHistHeight, ByRef lMaxHistMeanHeight)
    Dim varReturn As Variant
    Dim varChanData As Variant
    
    Dim dblStartTime As Double
    Dim dblEndTime As Double
    Dim dblSwepStartTime As Double
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Dim histTmp As Long
    Dim histVariance As Double
    Dim histStddev As Double
    Dim histMean As Double
    Dim nSweps As Long
    nSweps = 0

    Dim swepVals()

    varReturn = objTTX.GetEpocsExV("Swep", 0)
    If IsArray(varReturn) Then
        ReDim swepVals(UBound(varReturn, 2))
        nSweps = UBound(varReturn, 2) + 1
        For i = 0 To UBound(varReturn, 2)
            dblStartTime = varReturn(2, i) + lIgnoreFirstMsec
            dblEndTime = dblStartTime + lBinWidth + lIgnoreFirstMsec
            dblSwepStartTime = dblStartTime
            Do
                k = objTTX.ReadEventsV(500, "CSPK", iChanNum, 0, dblStartTime, dblEndTime, "JUSTTIMES")
                If k = 0 Then
                    Exit Do
                End If
    
                histTmp = CLng(histTmp) + CLng(k)
                swepVals(i) = CLng(swepVals(i)) + CLng(k)
                If k < 500 Then
                    Exit Do
                Else
                    varChanData = objTTX.ParseEvInfoV(k - 1, 1, 6)
                    dblStartTime = varChanData(0) + (1 / 100000)
                End If
                
            Loop
            dblStartTime = dblSwepStartTime
        Next
        
        histMean = CDbl(histTmp) / CDbl((UBound(swepVals) + 1))
        histVariance = 0#
        
        For i = 0 To UBound(swepVals)
            histVariance = histVariance + (histMean - CDbl(swepVals(i))) ^ 2
        Next
        histStddev = (histVariance / UBound(swepVals)) ^ 0.5
                
        If xAxisEp = "Channel" Then
            outputWorkbook.Worksheets("Totals").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Value = histTmp
            outputWorkbook.Worksheets("Means").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Value = histMean
            outputWorkbook.Worksheets("StdDev").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Value = histStddev
            outputWorkbook.Worksheets("N").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Value = nSweps
        ElseIf yAxisEp = "Channel" Then
            outputWorkbook.Worksheets("Totals").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Value = histTmp
            outputWorkbook.Worksheets("Means").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Value = histMean
            outputWorkbook.Worksheets("StdDev").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Value = histStddev
            outputWorkbook.Worksheets("N").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Value = nSweps
        Else
            outputWorkbook.Worksheets("Totals").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Value = histTmp
            outputWorkbook.Worksheets("Means").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Value = histMean
            outputWorkbook.Worksheets("StdDev").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Value = histStddev
            outputWorkbook.Worksheets("N").Cells(yOffset + iRowOffset + zOffset + 1, xOffset + iColOffset + 1).Value = nSweps
        End If
        If histMean > lMaxHistMeanHeight Then
            lMaxHistMeanHeight = histMean
        End If
        If histTmp > lMaxHistHeight Then
            lMaxHistHeight = histTmp
        End If
    End If
    
End Sub
Sub buildTuningCurvesIntoSigmaplot()
'    ImportFrom.Show
    
'    If doImport Then
'        Call processImport(True)
'    End If
    Call TransferToSigmaplot
End Sub
Sub transferToSigmaplotButton()
    plotWhichSheet = plotWorkbook.Worksheets("Settings").Range("B18").Value
    Set plotWorkbook = Application.ActiveWorkbook

    Dim zOffsetSize As Long
    Dim iColOffset As Integer
    Dim iRowOffset As Integer

    zOffsetSize = plotWorkbook.Worksheets("Variables (do not edit)").Range("H3").Value
    iColOffset = plotWorkbook.Worksheets("Variables (do not edit)").Range("H5").Value
    iRowOffset = plotWorkbook.Worksheets("Variables (do not edit)").Range("H6").Value

    Dim xPos As Long
    Dim yPos As Long
   
    xPos = iColOffset + 1
    yPos = iRowOffset
    
    Set dHeadingList = New Dictionary
    
    Do
        If plotWorkbook.Worksheets(plotWhichSheet).Cells(yPos, xPos).Value <> "" Then
            If Not dHeadingList.Exists(plotWorkbook.Worksheets(plotWhichSheet).Cells(yPos, xPos).Value) Then
                Call dHeadingList.Add(plotWorkbook.Worksheets(plotWhichSheet).Cells(yPos, xPos).Value, 0)
            End If
            yPos = yPos + zOffsetSize
        Else
            Exit Do
        End If
    Loop
    
    TransferToSigmaplotFrm.Show
    If doImport Then
        Call TransferToSigmaplot
    End If

End Sub

Sub transferAllToSigmaplot()
    plotWhichSheet = plotWorkbook.Worksheets("Settings").Range("B18").Value

    Dim zOffsetSize As Long
    Dim iColOffset As Integer
    Dim iRowOffset As Integer

    zOffsetSize = plotWorkbook.Worksheets("Variables (do not edit)").Range("H3").Value
    iColOffset = plotWorkbook.Worksheets("Variables (do not edit)").Range("H5").Value
    iRowOffset = plotWorkbook.Worksheets("Variables (do not edit)").Range("H6").Value

    Dim xPos As Long
    Dim yPos As Long
   
    xPos = iColOffset + 1
    yPos = iRowOffset
    
    Set dHeadingsSelected = New Dictionary
    
    Do
        If plotWorkbook.Worksheets(plotWhichSheet).Cells(yPos, xPos).Value <> "" Then
            If Not dHeadingsSelected.Exists(plotWorkbook.Worksheets(plotWhichSheet).Cells(yPos, xPos).Value) Then
                Call dHeadingsSelected.Add(plotWorkbook.Worksheets(plotWhichSheet).Cells(yPos, xPos).Value, 0)
            End If
            yPos = yPos + zOffsetSize
        Else
            Exit Do
        End If
    Loop
    
    If doImport Then
        Call TransferToSigmaplot
    End If
End Sub

Sub transferCandidatesToSigmaplot()
    plotWhichSheet = plotWorkbook.Worksheets("Settings").Range("B18").Value

'    Dim zOffsetSize As Long
'    Dim iColOffset As Integer
'    Dim iRowOffset As Integer

'    zOffsetSize = plotWorkbook.Worksheets("Variables (do not edit)").Range("H3").Value
'    iColOffset = plotWorkbook.Worksheets("Variables (do not edit)").Range("H5").Value
'    iRowOffset = plotWorkbook.Worksheets("Variables (do not edit)").Range("H6").Value

'    Dim xPos As Long
'    Dim yPos As Long
   
'    xPos = iColOffset + 1
'    yPos = iRowOffset
    
    Dim iRow As Integer
    iRow = 2
    
    Set dHeadingsSelected = New Dictionary
    
    Do
        If plotWorkbook.Worksheets("Likely tuned channels").Cells(iRow, 1).Value <> "" Then
            If Not dHeadingsSelected.Exists(plotWorkbook.Worksheets("Likely tuned channels").Cells(iRow, 1).Value) Then
                Call dHeadingsSelected.Add(plotWorkbook.Worksheets("Likely tuned channels").Cells(iRow, 1).Value, 0)
            End If
            iRow = iRow + 1
        Else
            Exit Do
        End If
    Loop
    
    If doImport Then
        Call TransferToSigmaplot
    End If
End Sub
Function getOutputDir(theOutputDir, fileOnTargetDrive) As String
    
    Dim objFS As FileSystemObject
    Set objFS = CreateObject("Scripting.FileSystemObject")
    
    If theOutputDir = "" Then
        theOutputDir = objFS.GetParentFolderName(fileOnTargetDrive)
    ElseIf Right(Left(theOutputDir, 2), 1) <> ":" Then
        Dim theDrive As String
        theDrive = objFS.GetDriveName(fileOnTargetDrive)
        theOutputDir = theDrive & theOutputDir
    End If
    
    If objFS.FolderExists(theOutputDir) Then
        getOutputDir = theOutputDir
    Else
        getOutputDir = ""
    End If
    
    Set objFS = Nothing
End Function

Function getTemplateFilename(templateName, fileOnTargetDrive) As String
    
    Dim objFS As FileSystemObject
    Set objFS = CreateObject("Scripting.FileSystemObject")
    
    If templateName = "" Then
        getTemplateFilename = ""
    Else
        If Right(Left(templateName, 2), 1) <> ":" Then
            Dim theDrive As String
            theDrive = objFS.GetDriveName(fileOnTargetDrive)
            getTemplateFilename = theDrive & templateName
        End If
        
        If Not objFS.FileExists(getTemplateFilename) Then
            getTemplateFilename = ""
        End If
    End If
    
    Set objFS = Nothing
    
End Function

Function checkPathExists(thePath As String) As Boolean
    
    
    If thePath = "" Then
        checkPathExists = False
    Else
        Dim objFS As FileSystemObject
        Set objFS = CreateObject("Scripting.FileSystemObject")
        If Not objFS.FolderExists(thePath) Then
            checkPathExists = False
        Else
            checkPathExists = True
        End If
        Set objFS = Nothing
    End If

End Function


Sub detectTunedSegments()
    If outputWorkbook Is Nothing Then
        Set outputWorkbook = Application.ActiveWorkbook
    End If
    Dim iOutputOffset As Integer
    iOutputOffset = 2
    
    outputWorkbook.Worksheets("Likely Tuned Channels").UsedRange.Clear

    Dim zOffsetSize As Long
    Dim iColOffset As Integer
    Dim iRowOffset As Integer

    Dim xCount As Integer
    Dim yCount As Integer

    zOffsetSize = outputWorkbook.Worksheets("Variables (do not edit)").Range("H3").Value
    iColOffset = outputWorkbook.Worksheets("Variables (do not edit)").Range("H5").Value
    iRowOffset = outputWorkbook.Worksheets("Variables (do not edit)").Range("H6").Value

    xCount = outputWorkbook.Worksheets("Variables (do not edit)").Range("H1").Value
    yCount = outputWorkbook.Worksheets("Variables (do not edit)").Range("H2").Value

    Dim xPos As Long
    Dim yPos As Long
   
    xPos = iColOffset + 1
    yPos = iRowOffset

    Dim dRowTotal As Double
    Dim dFirstRowTotal As Double
    
    Dim iRow As Integer
    Dim iCol As Integer
    Dim blnLooksGood As Boolean
    
    Do
        dRowTotal = 0#
        dFirstRowTotal = 0#
        If outputWorkbook.Worksheets("Means").Cells(yPos, xPos).Value <> "" Then
            blnLooksGood = True
            For iRow = (yPos + 2) To (yPos + yCount + 2) 'only want to look at the first 2 rows - after than there is no real guarantees
                For iCol = (xPos + 1) To (xPos + xCount + 1)
                    dRowTotal = dRowTotal + outputWorkbook.Worksheets("Means").Cells(iRow, iCol).Value
                Next
                If iRow > (yPos + 2) Then 'can only compare to previous row if not first row
                    If (dRowTotal * marginForGoodTuning) > dFirstRowTotal Then
                        blnLooksGood = False
                        Exit For
                    End If
                Else
                    dFirstRowTotal = dRowTotal
                End If
                dRowTotal = 0
            Next
            If blnLooksGood Then
                outputWorkbook.Worksheets("Likely Tuned Channels").Cells(iOutputOffset, 1).Value = outputWorkbook.Worksheets("Means").Cells(yPos, xPos).Value
                outputWorkbook.Worksheets("Likely Tuned Channels").Cells(iOutputOffset, 2).Value = yPos
                iOutputOffset = iOutputOffset + 1
            End If
            
            yPos = yPos + zOffsetSize
        Else
            Exit Do
        End If
    Loop
    
End Sub
