Attribute VB_Name = "Module1"
Option Explicit

'Global maxPreTrialTime As Double
'Global minSpikes As Double
    
'Global exIntCountGT As Integer
'Global exIntBeatsGT As Integer
'Global exLongestIntDurGT As Integer
'Global exLongestIntBeatsGT As Integer

Global pLess05FC As FormatCondition
Global pLess10FC As FormatCondition
Global excludedTrialCell As Range

Dim neuralByClass As Dictionary
Dim neuralByDate As Dictionary
Dim neuralByAcclim As Dictionary


'Global clusterByDate As Boolean
'Global clusterByStimParams As Boolean

Sub aggregrate_results()
    Dim exclusionInfo As Variant
    
    Dim animalID As String
    Dim experimentDate As String
    Dim experimentTag As String

    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Dim objFS As FileSystemObject
        
    Dim thisWorkbook As Workbook
    
    Dim pathToData As String
    
    'get the root folder under which all data is housed
    Dim rootFolder As Folder
       
    Dim AnimalFolders As Folders
    Dim objAnimalFolder As Folder
    
    Dim experimentFolders As Folders
    Dim objExpFolder As Folder
    
    Dim Files As Files
    Dim objFile As File
    
    Dim strExcelFilename As String
    Dim strExcelPathname As String
    
    Dim blnCurrFolderIsTrial As Boolean
       
    Dim thisAnimalWorksheet As Worksheet
    Dim thisAnimalTrialsRow As Long
    Dim workbookToProcess As Workbook
        
    Set thisWorkbook = ActiveWorkbook
    
'    maxPercOfBeatsInt = thisWorkbook.Worksheets("Controller").Cells(3, 2).Value
'    maxSingleIntSamples = thisWorkbook.Worksheets("Controller").Cells(4, 2).Value
'    maxSingleIntBeats = thisWorkbook.Worksheets("Controller").Cells(5, 2).Value

    Set objFS = CreateObject("Scripting.FileSystemObject")
    
    pathToData = objFS.GetDriveName(thisWorkbook.FullName) & thisWorkbook.Worksheets("Controller").Cells(19, 2).Value
    Set rootFolder = objFS.GetFolder(pathToData)
    
    blnCurrFolderIsTrial = False
        
    Call deleteOldWorksheets(thisWorkbook)
    
    Set AnimalFolders = rootFolder.Subfolders
    For Each objAnimalFolder In AnimalFolders 'cycle through the folder for each animal
        exclusionInfo = checkForExclusion(objAnimalFolder)
        If exclusionInfo(0) = "folder" Then
            Stop
        Else
            thisAnimalTrialsRow = 3
            Set thisAnimalWorksheet = Nothing
            animalID = objAnimalFolder.Name
                        
            Set experimentFolders = objAnimalFolder.Subfolders
            For Each objExpFolder In experimentFolders 'go through the experiments within an animal folder
                blnCurrFolderIsTrial = False
                exclusionInfo = checkForExclusion(objExpFolder)
                If Not (exclusionInfo(0) <> "" And exclusionInfo(1) = "") Then 'check if the exclusion includes a message, or is only for some types of trial
                    experimentDate = Left(objExpFolder.Name, 10)
                    experimentTag = objExpFolder.Name
                    blnCurrFolderIsTrial = False 'assume until file detected this is not an experiment
                    strExcelFilename = ""
                    Set Files = objExpFolder.Files
                    For Each objFile In Files 'inside the experiment file, check for a file ending with ".adicht" to check it is a test
                        If UCase(Right(objFile.Name, 7)) = ".ADICHT" Then
                            blnCurrFolderIsTrial = True
                        End If
                        If UCase(Right(objFile.Name, 5)) = ".XLSM" And Left(objFile.Name, 1) <> "~" Then 'locate the spreadsheet file with trial data
                            strExcelFilename = objFile.Name
                            strExcelPathname = objFile.Path
                        End If
                    Next
                    
                    If blnCurrFolderIsTrial Then
                        If strExcelFilename <> "" Then
                            'open the workbook to read data from
                            Set workbookToProcess = Workbooks.Open(strExcelPathname)
                            If thisAnimalWorksheet Is Nothing Then
                                Call thisWorkbook.Worksheets("Trials").Copy(, thisWorkbook.Worksheets("Trials"))
                                Set thisAnimalWorksheet = thisWorkbook.Worksheets("Trials (2)")
                                thisAnimalWorksheet.Name = animalID
                            End If
                            Call copyTrials(workbookToProcess, experimentDate, experimentTag, exclusionInfo, thisAnimalWorksheet, thisAnimalTrialsRow, strExcelPathname)
                            Call workbookToProcess.Close
                        End If
                    End If
                End If
            Next
        End If
    Next
            
    Set objFile = Nothing
    Set Files = Nothing
                    
    Set experimentFolders = Nothing
    Set objExpFolder = Nothing
    
    Set AnimalFolders = Nothing
    Set objAnimalFolder = Nothing
               
    Set rootFolder = Nothing
    Set objFS = Nothing
    
    thisWorkbook.Worksheets("Controller").Range("G4").Value = Now()
    
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
End Sub
Function copyTrials(workbookToProcess As Workbook, experimentDate As String, experimentTag As String, exclusionInfo As Variant, thisAnimalWorksheet As Worksheet, ByRef thisAnimalTrialsRow As Long, strExcelPathname As String)
    Dim exclusionReason As String
    Dim strTrialString As String
    Dim iTrialNum As Integer
    Dim lNeuroSourceRow As Long
    Dim lNeuroOffset As Long
    Dim lSourceRow As Long
    lNeuroSourceRow = 1
    lSourceRow = 2
    
    If workbookToProcess.Worksheets("Neural Data").Range("A1").Value <> "" Then
        While workbookToProcess.Worksheets("Output").Cells(lSourceRow, 1) <> ""
            iTrialNum = workbookToProcess.Worksheets("Output").Range("B" & lSourceRow).Value
            Do
                strTrialString = workbookToProcess.Worksheets("Neural Data").Range("A" & lNeuroSourceRow).Value
                If lNeuroSourceRow > 50000 Then
                    Exit Do
                End If
                
                If LCase(Left(strTrialString, Len("trial"))) = "trial" Then
                    If CInt(Right(strTrialString, Len(strTrialString) - Len("trial") - 1) = iTrialNum) Then
                        Exit Do
                    End If
                End If
                lNeuroSourceRow = lNeuroSourceRow + 1
            Loop
    
            If lNeuroSourceRow < 50000 Then
                For lNeuroOffset = 0 To 31 'step through the channels
                    If workbookToProcess.Worksheets("Neural Data").Range("C" & lNeuroSourceRow + 2 + (lNeuroOffset * 2)).Value > 0 Then
                        thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 1).Value = strExcelPathname
                        thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 2).Value = workbookToProcess.Worksheets("Variables (do not edit)").Range("B2").Value
                        thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 3).Value = workbookToProcess.Worksheets("Variables (do not edit)").Range("B3").Value
                        thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 4).Value = experimentDate
                        thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 5).Value = experimentTag
                        thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 6).Value = exclusionInfo(0)
                        thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 7).Value = exclusionInfo(1)
                        thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 8).Value = exclusionInfo(2)
                        thisAnimalWorksheet.Range(thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 9), thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 22)).Value = workbookToProcess.Worksheets("Output").Range("A" & lSourceRow & ":N" & lSourceRow).Value
                               
                        'freq
                        thisAnimalWorksheet.Range("X" & thisAnimalTrialsRow).Value = workbookToProcess.Worksheets("Neural Data").Range("D" & lNeuroSourceRow).Value
                        'attn 1
                        thisAnimalWorksheet.Range("Y" & thisAnimalTrialsRow).Value = workbookToProcess.Worksheets("Neural Data").Range("B" & lNeuroSourceRow + 6).Value
                        'attn 1 1-4 count
                        thisAnimalWorksheet.Range("Z" & thisAnimalTrialsRow).Value = workbookToProcess.Worksheets("Neural Data").Range("B" & lNeuroSourceRow + 7).Value
                        'attn 1 5-8 count
                        thisAnimalWorksheet.Range("AA" & thisAnimalTrialsRow).Value = workbookToProcess.Worksheets("Neural Data").Range("B" & lNeuroSourceRow + 8).Value
                        
                        'attn 2
                        thisAnimalWorksheet.Range("AB" & thisAnimalTrialsRow).Value = workbookToProcess.Worksheets("Neural Data").Range("B" & lNeuroSourceRow + 10).Value
                        'attn 2 1-4 count
                        thisAnimalWorksheet.Range("AC" & thisAnimalTrialsRow).Value = workbookToProcess.Worksheets("Neural Data").Range("B" & lNeuroSourceRow + 11).Value
                        'attn 2 5-8 count
                        thisAnimalWorksheet.Range("AD" & thisAnimalTrialsRow).Value = workbookToProcess.Worksheets("Neural Data").Range("B" & lNeuroSourceRow + 12).Value
                        
                        'attn 3
                        thisAnimalWorksheet.Range("AE" & thisAnimalTrialsRow).Value = workbookToProcess.Worksheets("Neural Data").Range("B" & lNeuroSourceRow + 14).Value
                        'attn 3 1-4 count
                        thisAnimalWorksheet.Range("AF" & thisAnimalTrialsRow).Value = workbookToProcess.Worksheets("Neural Data").Range("B" & lNeuroSourceRow + 15).Value
                        'attn 3 5-8 count
                        thisAnimalWorksheet.Range("AG" & thisAnimalTrialsRow).Value = workbookToProcess.Worksheets("Neural Data").Range("B" & lNeuroSourceRow + 16).Value
                        
                        'channel
                        thisAnimalWorksheet.Range("AJ" & thisAnimalTrialsRow).Value = workbookToProcess.Worksheets("Neural Data").Range("D" & lNeuroSourceRow + 2 + (lNeuroOffset * 2)).Value
                        
                        'pre total
                        thisAnimalWorksheet.Range("AK" & thisAnimalTrialsRow).Value = workbookToProcess.Worksheets("Neural Data").Range("E" & lNeuroSourceRow + 2 + (lNeuroOffset * 2)).Value
                        'post total
                        thisAnimalWorksheet.Range("AL" & thisAnimalTrialsRow).Value = workbookToProcess.Worksheets("Neural Data").Range("E" & lNeuroSourceRow + 2 + (lNeuroOffset * 2) + 1).Value
    
                        'pre mean
                        thisAnimalWorksheet.Range("AM" & thisAnimalTrialsRow).Value = workbookToProcess.Worksheets("Neural Data").Range("H" & lNeuroSourceRow + 2 + (lNeuroOffset * 2)).Value
                        'post mean
                        thisAnimalWorksheet.Range("AN" & thisAnimalTrialsRow).Value = workbookToProcess.Worksheets("Neural Data").Range("H" & lNeuroSourceRow + 2 + (lNeuroOffset * 2) + 1).Value
    
                        'pre stddev
                        thisAnimalWorksheet.Range("AO" & thisAnimalTrialsRow).Value = workbookToProcess.Worksheets("Neural Data").Range("K" & lNeuroSourceRow + 2 + (lNeuroOffset * 2)).Value
                        'post mean
                        thisAnimalWorksheet.Range("AP" & thisAnimalTrialsRow).Value = workbookToProcess.Worksheets("Neural Data").Range("K" & lNeuroSourceRow + 2 + (lNeuroOffset * 2) + 1).Value
    
                        '1-4 total amp1
                        thisAnimalWorksheet.Range("AR" & thisAnimalTrialsRow).Value = workbookToProcess.Worksheets("Neural Data").Range("Y" & lNeuroSourceRow + 2 + (lNeuroOffset * 2)).Value
                        '1-4 total amp2
                        thisAnimalWorksheet.Range("AS" & thisAnimalTrialsRow).Value = workbookToProcess.Worksheets("Neural Data").Range("AA" & lNeuroSourceRow + 2 + (lNeuroOffset * 2)).Value
                        '1-4 total amp3
                        thisAnimalWorksheet.Range("AT" & thisAnimalTrialsRow).Value = workbookToProcess.Worksheets("Neural Data").Range("AC" & lNeuroSourceRow + 2 + (lNeuroOffset * 2)).Value
                        
                        '5-8 total amp1
                        thisAnimalWorksheet.Range("AV" & thisAnimalTrialsRow).Value = workbookToProcess.Worksheets("Neural Data").Range("Y" & lNeuroSourceRow + 2 + (lNeuroOffset * 2) + 1).Value
                        '5-8 total amp2
                        thisAnimalWorksheet.Range("AW" & thisAnimalTrialsRow).Value = workbookToProcess.Worksheets("Neural Data").Range("AA" & lNeuroSourceRow + 2 + (lNeuroOffset * 2) + 1).Value
                        '5-8 total amp3
                        thisAnimalWorksheet.Range("AX" & thisAnimalTrialsRow).Value = workbookToProcess.Worksheets("Neural Data").Range("AC" & lNeuroSourceRow + 2 + (lNeuroOffset * 2) + 1).Value
                        
                        
                        'post mean
                        thisAnimalWorksheet.Range("AP" & thisAnimalTrialsRow).Value = workbookToProcess.Worksheets("Neural Data").Range("K" & lNeuroSourceRow + 2 + (lNeuroOffset * 2) + 1).Value
    
    
                        
                        'thisAnimalWorksheet.Range(thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 9), thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 22)).Value = workbookToProcess.Worksheets("Output").Range("A" & lSourceRow & ":N" & lSourceRow).Value
        
                        thisAnimalTrialsRow = thisAnimalTrialsRow + 1
                    End If
                Next
            End If
            lSourceRow = lSourceRow + 1
        Wend
    End If
End Function

Sub processTrials()
    Dim exclusionInfo As Variant

    Dim templateFilename As String
    Dim validTrialCount As Integer
    
    Dim animalID As String
    Dim experimentDate As String
    Dim experimentTag As String

    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Dim objFS As FileSystemObject
        
    Dim thisWorkbook As Workbook
    
    Dim pathToData As String
    
    Dim theDict As Dictionary
    
    Dim thisAnimalWorksheet As Worksheet
    Dim thisAnimalSummarySheet As Worksheet
    Dim thisAnimalSummarySheetRow As Long
    Dim outputWorkbook As Workbook
    Dim workbookToProcess As Workbook
    
    Dim outputFilename As String
        
    Set thisWorkbook = ActiveWorkbook
    
    templateFilename = "\Code current\Excel tools\aggregate neural results output.xltm"
    Set objFS = CreateObject("Scripting.FileSystemObject")
    templateFilename = objFS.GetDriveName(thisWorkbook.FullName) & templateFilename 'get the drive letter for the template
    
    pathToData = objFS.GetDriveName(thisWorkbook.FullName) & thisWorkbook.Worksheets("Controller").Cells(19, 2).Value
    
    Set pLess05FC = thisWorkbook.Worksheets("Controller").Range("B11").FormatConditions(1)
    Set pLess10FC = thisWorkbook.Worksheets("Controller").Range("B12").FormatConditions(1)
    
    Set excludedTrialCell = thisWorkbook.Worksheets("Controller").Range("B17")
    
    Dim iSourceWorksheetOffset As Integer
    Dim sourceWorksheet As Worksheet
    
    Dim iPass As Integer
 
    Dim outputByDate As Workbook
    Dim outputByAcclim As Workbook
    Dim outputByClass As Workbook
    
    Dim iColHeadersForHRLine As Integer
        
        For iSourceWorksheetOffset = 1 To (thisWorkbook.Worksheets.Count)
            If thisWorkbook.Worksheets(iSourceWorksheetOffset).Name <> "Controller" And thisWorkbook.Worksheets(iSourceWorksheetOffset).Name <> "Trials" Then 'check if this is actually a data sheet
                Set sourceWorksheet = thisWorkbook.Worksheets(iSourceWorksheetOffset)
                
                Set neuralByClass = New Dictionary
                Set neuralByDate = New Dictionary
                Set neuralByAcclim = New Dictionary

                validTrialCount = 0
                Set thisAnimalWorksheet = Nothing
                Set thisAnimalSummarySheet = Nothing
                animalID = sourceWorksheet.Name
                
                Call parseTrials(sourceWorksheet, animalID)
                For iPass = 2 To 2
                    thisAnimalSummarySheetRow = 2
                    Select Case iPass
                        Case 0:
                            'clusterByStimParams = True
                            'clusterByDate = False
                            If outputByAcclim Is Nothing Then
                                Set outputByAcclim = Workbooks.Open(templateFilename)
                            End If
                            Set outputWorkbook = outputByAcclim
                            Set theDict = neuralByAcclim
                        Case 1:
                            'clusterByStimParams = False
                            'clusterByDate = True
                            If outputByDate Is Nothing Then
                                Set outputByDate = Workbooks.Open(templateFilename)
                            End If
                            Set outputWorkbook = outputByDate
                            Set theDict = neuralByDate
                        Case 2:
                            'clusterByStimParams = False
                            'clusterByDate = True
                            If outputByClass Is Nothing Then
                                Set outputByClass = Workbooks.Open(templateFilename)
                            End If
                            Set outputWorkbook = outputByClass
                            Set theDict = neuralByClass
                    End Select

                    Call outputWorkbook.Worksheets("Summary template").Copy(, outputWorkbook.Worksheets("Output template"))
                    Set thisAnimalSummarySheet = outputWorkbook.Worksheets("Summary template (2)")
                    thisAnimalSummarySheet.Name = animalID & " summary"
                    
                    thisAnimalSummarySheet.Cells(1, 1).Value = "Cluster"
                    thisAnimalSummarySheet.Cells(1, 2).Value = "Included trials"
                    thisAnimalSummarySheet.Cells(1, 3).Value = "Excluded trials"
                    thisAnimalSummarySheet.Cells(1, 4).Value = "% trials increase in spikes"
                    thisAnimalSummarySheet.Cells(1, 5).Value = "Mean spike change"
                    thisAnimalSummarySheet.Cells(1, 6).Value = "spike std dtv"
                    thisAnimalSummarySheet.Cells(1, 7).Value = "T score"
                    thisAnimalSummarySheet.Cells(1, 8).Value = "p value"
                    thisAnimalSummarySheet.Range("A1:R1").Font.Bold = True
                                        
                                        
                    Call outputWorkbook.Worksheets("Output template").Copy(, outputWorkbook.Worksheets("Output template"))
                    Set thisAnimalWorksheet = outputWorkbook.Worksheets("Output template (2)")
                    thisAnimalWorksheet.Name = animalID
                    Call outputTrials(theDict, thisAnimalWorksheet, thisAnimalSummarySheet, thisAnimalSummarySheetRow, sourceWorksheet)
                Next
            End If
        Next
        If Not outputByClass Is Nothing Then
            outputFilename = pathToData & "\neural aggregate by class.xlsx"
            Call outputByClass.SaveAs(outputFilename)
            Call outputByClass.Close
        End If
        If Not outputByDate Is Nothing Then
            outputFilename = pathToData & "\neural aggregate by date.xlsx"
            Call outputByDate.SaveAs(outputFilename)
            Call outputByDate.Close
        End If
        If Not outputByAcclim Is Nothing Then
            outputFilename = pathToData & "\neural aggregate by acclim.xlsx"
            Call outputByAcclim.SaveAs(outputFilename)
            Call outputByAcclim.Close
        End If
    
    Set objFS = Nothing
    
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic

End Sub

'Function parseTrials(outputDict As Dictionary, sourceWorksheet As Workbook, experimentDate As String, experimentTag As String, exclusionInfo As Variant)
Function parseTrials(sourceWorksheet As Worksheet, strAnimal As String)
    Dim experimentDate As String
    Dim experimentTag As String
    Dim iChannel As Integer

    'Dim neuralByDateInfo As String
    Dim neuralByAcclimInfo As String

    Dim i As Integer
    i = 3

    While sourceWorksheet.Range("A" & i).Value <> ""
        experimentDate = sourceWorksheet.Range("D" & i).Value
        experimentTag = sourceWorksheet.Range("E" & i).Value
        iChannel = sourceWorksheet.Range("AJ" & i).Value
        
        If InStr(1, experimentTag, "acclimatisation", vbTextCompare) > 0 Then
            neuralByAcclimInfo = "Acclimatisation"
        Else
            neuralByAcclimInfo = "Trials"
        End If
        
        Select Case strAnimal
            Case "111_140":
                Call addToDict(neuralByClass, neuralByAcclimInfo, "DCN", i)
            Case "111_141":
                Select Case iChannel:
                    Case 4, 5, 6, 7, 8, 9, 25:
                        Call addToDict(neuralByClass, neuralByAcclimInfo, "VCN", i)
                    Case 10, 11, 12, 13, 14, 15, 25:
                        Call addToDict(neuralByClass, neuralByAcclimInfo, "DCN", i)
                    Case 26, 27, 28, 29, 30, 31, 32:
                        Call addToDict(neuralByClass, neuralByAcclimInfo, "Octopus", i)
                End Select
            Case "112_1024":
                Select Case iChannel:
                    Case 17, 18, 19, 20, 21, 22, 23, 8, 9, 10, 11, 12, 13, 14, 15, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32:
                    'Case 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32:
                        Call addToDict(neuralByClass, neuralByAcclimInfo, "VCN", i)
                    'Case 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16:
                    '    Call addToDict(neuralByClass, experimentDate, "DCN", i)
                End Select
            Case "123_1164":
                'Select Case iChannel:
'                    Case 21:
'                        Call addToDict(neuralByClass, experimentDate, "DCN", i)
'                    Case 7, 15:
                        Call addToDict(neuralByClass, neuralByAcclimInfo, "AVCN", i)
'                    Case 5, 6, 8, 9, 10, 11, 13:
'                        Call addToDict(neuralByClass, experimentDate, "VCN", i)
                'End Select
                'Call addToDict(neuralByClass, experimentDate, "DCN", i)
                'Call addToDict(neuralByClass, experimentDate, "VCN", i)
                'Call addToDict(neuralByClass, experimentDate, "Octopus", i)
        End Select

        
        Call addToDict(neuralByDate, experimentDate, CStr(iChannel), i)
        Call addToDict(neuralByAcclim, neuralByAcclimInfo, CStr(iChannel), i)
        i = i + 1
    Wend
End Function



Sub outputTrials(theDict As Dictionary, thisAnimalWorksheet As Worksheet, thisAnimalSummarySheet As Worksheet, ByRef thisAnimalSummarySheetRow, ByRef sourceWorksheet As Worksheet)
    Dim arrChannels As Variant
    arrChannels = theDict.Keys
    Dim iChanNum As Integer
        
    Dim formatCond As FormatCondition
    
    Dim dictSublevel As Dictionary
    Dim arrSubLevels As Variant
    Dim iSubLevelNum As Integer
    
    Dim arrTrials As Variant
    
    Dim TotalVal As Double
    Dim TotalSD As Double
    Dim TotalIterator As Integer
    Dim TotalnInSoFar As Integer

    Dim thisVal As Double
    Dim thisSD As Double
    Dim thisIterator As Integer
    Dim thisnInSoFar As Integer
    
    Dim lRowNum As Long
    Dim lTrialNum As Long
    
    Dim iExcelOffset As Long
    iExcelOffset = 1
    Dim iPrevExcelOffset As Long
    Dim iMaxExcelOffset As Long
    
    Dim iThisAnimalSummarySheetStartingRow As Integer
    iThisAnimalSummarySheetStartingRow = CInt(thisAnimalSummarySheetRow)
    
    Dim changeSum As Double
    Dim changeSumSqr As Double
    Dim nInMeanSoFar As Integer
    Dim nExcluded As Integer
    Dim diff As Double
    Dim tStat As Double
        
    Dim changeMean As Double
    Dim changeVar As Double
    
    'Dim dblTotal14 As Double
    'Dim dblTotal58 As Double
    Dim dblTotal1420 As Double
    Dim dblTotal1425 As Double
    Dim dblTotal1430 As Double
    
    Dim dblTotal5820 As Double
    Dim dblTotal5825 As Double
    Dim dblTotal5830 As Double
    
    Dim lTotalAttn1420 As Long
    Dim lTotalAttn1425 As Long
    Dim lTotalAttn1430 As Long
    
    Dim lTotalAttn5820 As Long
    Dim lTotalAttn5825 As Long
    Dim lTotalAttn5830 As Long
    
    Dim myChart As ChartObject
    Dim lChartPos As Long
    
    Dim i As Integer
    
    Dim iSummaryCol As Integer

    For iChanNum = 0 To UBound(arrChannels)
        thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "Channel " & arrChannels(iChanNum)

        thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Bold = True
        iExcelOffset = iExcelOffset + 1
    
        Set dictSublevel = theDict(arrChannels(iChanNum))
        arrSubLevels = dictSublevel.Keys
        For iSubLevelNum = 0 To UBound(arrSubLevels)
            dblTotal1420 = 0
            dblTotal1425 = 0
            dblTotal1430 = 0
            
            dblTotal5820 = 0
            dblTotal5825 = 0
            dblTotal5830 = 0
            
            lTotalAttn1420 = 0
            lTotalAttn1425 = 0
            lTotalAttn1430 = 0
            
            lTotalAttn5820 = 0
            lTotalAttn5825 = 0
            lTotalAttn5830 = 0
            'thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 1).Style.Name = "Normal"

            thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 1).Font.Bold = True
            
            thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = arrSubLevels(iSubLevelNum)
            thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Bold = True
            iExcelOffset = iExcelOffset + 1
            
            thisAnimalWorksheet.Range("A" & iExcelOffset, "H" & iExcelOffset).Font.Italic = True
            thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "Tag"
            thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = "Trial Number"
            thisAnimalWorksheet.Cells(iExcelOffset, 3).Value = "Date"
            thisAnimalWorksheet.Cells(iExcelOffset, 4).Value = "Atten 1"
            thisAnimalWorksheet.Cells(iExcelOffset, 5).Value = "Atten 1 1-4 count"
            thisAnimalWorksheet.Cells(iExcelOffset, 6).Value = "Atten 1 5-8 count"
            thisAnimalWorksheet.Cells(iExcelOffset, 7).Value = "Atten 2"
            thisAnimalWorksheet.Cells(iExcelOffset, 8).Value = "Atten 2 1-4 count"
            thisAnimalWorksheet.Cells(iExcelOffset, 9).Value = "Atten 2 5-8 count"
            thisAnimalWorksheet.Cells(iExcelOffset, 10).Value = "Atten 3"
            thisAnimalWorksheet.Cells(iExcelOffset, 11).Value = "Atten 3 1-4 count"
            thisAnimalWorksheet.Cells(iExcelOffset, 12).Value = "Atten 3 5-8 count"
            thisAnimalWorksheet.Cells(iExcelOffset, 13).Value = "Mean 1-4 spikes"
            thisAnimalWorksheet.Cells(iExcelOffset, 14).Value = "Mean 5-8 spikes"
            thisAnimalWorksheet.Cells(iExcelOffset, 15).Value = "Total 1-4 spikes"
            thisAnimalWorksheet.Cells(iExcelOffset, 16).Value = "Total 5-8 spikes"
            thisAnimalWorksheet.Cells(iExcelOffset, 17).Value = "StdDev 1-4 spikes"
            thisAnimalWorksheet.Cells(iExcelOffset, 18).Value = "StdDev 5-8 spikes"
            thisAnimalWorksheet.Cells(iExcelOffset, 19).Value = "Overall trial exclusion reason"
            iExcelOffset = iExcelOffset + 1
            
            arrTrials = dictSublevel(arrSubLevels(iSubLevelNum))
            
            nExcluded = 0
            nInMeanSoFar = 0
            changeSum = 0
            changeSumSqr = 0
                            
            For lTrialNum = 0 To UBound(arrTrials)
                lRowNum = arrTrials(lTrialNum)
                'If sourceWorksheet.Range("Z" & lRowNum).Value <> 0 And sourceWorksheet.Range("AA" & lRowNum).Value <> 0 Then
                    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 1) = "'" & arrChannels(iChanNum) & ": " & CStr(arrSubLevels(iSubLevelNum))
                    If Not arrSubLevels(iSubLevelNum) = "Acclimatisation" And Not arrSubLevels(iSubLevelNum) = "Trials" Then
                        If InStr(1, sourceWorksheet.Range("E" & lRowNum).Value, "acclimatisation", vbTextCompare) > 0 Then
                            thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 1) = thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 1) & ": Acclimatisation"
                        Else
                            thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 1) = thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 1) & ": Trials"
                        End If
                    End If
                    thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = sourceWorksheet.Range("E" & lRowNum).Value
                    thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = sourceWorksheet.Range("J" & lRowNum).Value
                    thisAnimalWorksheet.Cells(iExcelOffset, 3).Value = sourceWorksheet.Range("D" & lRowNum).Value
                    
                    'amplitudes
                    'thisAnimalWorksheet.Cells(iExcelOffset, 4).Value = sourceWorksheet.Range("Y" & lRowNum).Value
                    'thisAnimalWorksheet.Cells(iExcelOffset, 5).Value = sourceWorksheet.Range("Z" & lRowNum).Value
                    'thisAnimalWorksheet.Cells(iExcelOffset, 6).Value = sourceWorksheet.Range("AA" & lRowNum).Value
                    'thisAnimalWorksheet.Cells(iExcelOffset, 7).Value = sourceWorksheet.Range("AB" & lRowNum).Value
                    'thisAnimalWorksheet.Cells(iExcelOffset, 8).Value = sourceWorksheet.Range("AC" & lRowNum).Value
                    'thisAnimalWorksheet.Cells(iExcelOffset, 9).Value = sourceWorksheet.Range("AD" & lRowNum).Value
                    'thisAnimalWorksheet.Cells(iExcelOffset, 10).Value = sourceWorksheet.Range("AE" & lRowNum).Value
                    'thisAnimalWorksheet.Cells(iExcelOffset, 11).Value = sourceWorksheet.Range("AF" & lRowNum).Value
                    'thisAnimalWorksheet.Cells(iExcelOffset, 12).Value = sourceWorksheet.Range("AG" & lRowNum).Value
                    
                    thisAnimalWorksheet.Cells(iExcelOffset, 13).Value = sourceWorksheet.Range("AM" & lRowNum).Value
                    thisAnimalWorksheet.Cells(iExcelOffset, 14).Value = sourceWorksheet.Range("AN" & lRowNum).Value
                    
                    thisAnimalWorksheet.Cells(iExcelOffset, 15).Value = sourceWorksheet.Range("AK" & lRowNum).Value
                    thisAnimalWorksheet.Cells(iExcelOffset, 16).Value = sourceWorksheet.Range("AL" & lRowNum).Value
        
                    thisAnimalWorksheet.Cells(iExcelOffset, 17).Value = sourceWorksheet.Range("AO" & lRowNum).Value
                    thisAnimalWorksheet.Cells(iExcelOffset, 18).Value = sourceWorksheet.Range("AP" & lRowNum).Value

                    'spike counts
                    'thisAnimalWorksheet.Cells(iExcelOffset, 20).Value = sourceWorksheet.Range("AR" & lRowNum).Value
                    'thisAnimalWorksheet.Cells(iExcelOffset, 21).Value = sourceWorksheet.Range("AS" & lRowNum).Value
                    'thisAnimalWorksheet.Cells(iExcelOffset, 22).Value = sourceWorksheet.Range("AT" & lRowNum).Value
        
                    'thisAnimalWorksheet.Cells(iExcelOffset, 24).Value = sourceWorksheet.Range("AV" & lRowNum).Value
                    'thisAnimalWorksheet.Cells(iExcelOffset, 25).Value = sourceWorksheet.Range("AW" & lRowNum).Value
                    'thisAnimalWorksheet.Cells(iExcelOffset, 26).Value = sourceWorksheet.Range("AX" & lRowNum).Value
        
        
                    If sourceWorksheet.Range("G" & lRowNum).Value = "" Or (sourceWorksheet.Range("G" & lRowNum).Value <> "" And sourceWorksheet.Range("H" & lRowNum).Value >= sourceWorksheet.Range("K" & lRowNum).Value) Then 'check if the data should be excluded
                        nInMeanSoFar = nInMeanSoFar + 1
                        'dblTotal1420 = dblTotal1420 + thisAnimalWorksheet.Cells(iExcelOffset, 15).Value
                        'dblTotal5820 = dblTotal5820 + thisAnimalWorksheet.Cells(iExcelOffset, 16).Value
                        'diff = thisAnimalWorksheet.Cells(iExcelOffset, 14).Value - thisAnimalWorksheet.Cells(iExcelOffset, 13).Value
                        'changeSum = changeSum + diff
                        'changeSumSqr = changeSumSqr + diff ^ 2
                        
                        Select Case sourceWorksheet.Range("Y" & lRowNum).Value
                            Case 20:
                                    thisAnimalWorksheet.Cells(iExcelOffset, 4) = 20
                                    thisAnimalWorksheet.Cells(iExcelOffset, 5).Value = sourceWorksheet.Range("Z" & lRowNum).Value
                                    thisAnimalWorksheet.Cells(iExcelOffset, 6).Value = sourceWorksheet.Range("AA" & lRowNum).Value
                                    thisAnimalWorksheet.Cells(iExcelOffset, 20).Value = sourceWorksheet.Range("AR" & lRowNum).Value
                                    thisAnimalWorksheet.Cells(iExcelOffset, 24).Value = sourceWorksheet.Range("AV" & lRowNum).Value
                            Case 25:
                                    thisAnimalWorksheet.Cells(iExcelOffset, 7) = 25
                                    thisAnimalWorksheet.Cells(iExcelOffset, 8).Value = sourceWorksheet.Range("Z" & lRowNum).Value
                                    thisAnimalWorksheet.Cells(iExcelOffset, 9).Value = sourceWorksheet.Range("AA" & lRowNum).Value
                                    thisAnimalWorksheet.Cells(iExcelOffset, 21).Value = sourceWorksheet.Range("AR" & lRowNum).Value
                                    thisAnimalWorksheet.Cells(iExcelOffset, 25).Value = sourceWorksheet.Range("AV" & lRowNum).Value
                            Case 30:
                                    thisAnimalWorksheet.Cells(iExcelOffset, 10) = 30
                                    thisAnimalWorksheet.Cells(iExcelOffset, 11).Value = sourceWorksheet.Range("Z" & lRowNum).Value
                                    thisAnimalWorksheet.Cells(iExcelOffset, 12).Value = sourceWorksheet.Range("AA" & lRowNum).Value
                                    thisAnimalWorksheet.Cells(iExcelOffset, 22).Value = sourceWorksheet.Range("AR" & lRowNum).Value
                                    thisAnimalWorksheet.Cells(iExcelOffset, 26).Value = sourceWorksheet.Range("AV" & lRowNum).Value
                        End Select
                        
                        Select Case sourceWorksheet.Range("AB" & lRowNum).Value
                            Case 20:
                                    thisAnimalWorksheet.Cells(iExcelOffset, 4) = 20
                                    thisAnimalWorksheet.Cells(iExcelOffset, 5).Value = sourceWorksheet.Range("AC" & lRowNum).Value
                                    thisAnimalWorksheet.Cells(iExcelOffset, 6).Value = sourceWorksheet.Range("AD" & lRowNum).Value
                                    thisAnimalWorksheet.Cells(iExcelOffset, 20).Value = sourceWorksheet.Range("AS" & lRowNum).Value
                                    thisAnimalWorksheet.Cells(iExcelOffset, 24).Value = sourceWorksheet.Range("AW" & lRowNum).Value
                            Case 25:
                                    thisAnimalWorksheet.Cells(iExcelOffset, 7) = 25
                                    thisAnimalWorksheet.Cells(iExcelOffset, 8).Value = sourceWorksheet.Range("AC" & lRowNum).Value
                                    thisAnimalWorksheet.Cells(iExcelOffset, 9).Value = sourceWorksheet.Range("AD" & lRowNum).Value
                                    thisAnimalWorksheet.Cells(iExcelOffset, 21).Value = sourceWorksheet.Range("AS" & lRowNum).Value
                                    thisAnimalWorksheet.Cells(iExcelOffset, 25).Value = sourceWorksheet.Range("AW" & lRowNum).Value
                            Case 30:
                                    thisAnimalWorksheet.Cells(iExcelOffset, 10) = 30
                                    thisAnimalWorksheet.Cells(iExcelOffset, 11).Value = sourceWorksheet.Range("AC" & lRowNum).Value
                                    thisAnimalWorksheet.Cells(iExcelOffset, 12).Value = sourceWorksheet.Range("AD" & lRowNum).Value
                                    thisAnimalWorksheet.Cells(iExcelOffset, 22).Value = sourceWorksheet.Range("AS" & lRowNum).Value
                                    thisAnimalWorksheet.Cells(iExcelOffset, 26).Value = sourceWorksheet.Range("AW" & lRowNum).Value
                            End Select


                        Select Case sourceWorksheet.Range("AE" & lRowNum).Value
                            Case 20:
                                    thisAnimalWorksheet.Cells(iExcelOffset, 4) = 20
                                    thisAnimalWorksheet.Cells(iExcelOffset, 5).Value = sourceWorksheet.Range("AF" & lRowNum).Value
                                    thisAnimalWorksheet.Cells(iExcelOffset, 6).Value = sourceWorksheet.Range("AG" & lRowNum).Value
                                    thisAnimalWorksheet.Cells(iExcelOffset, 20).Value = sourceWorksheet.Range("AT" & lRowNum).Value
                                    thisAnimalWorksheet.Cells(iExcelOffset, 24).Value = sourceWorksheet.Range("AX" & lRowNum).Value
                            Case 25:
                                    thisAnimalWorksheet.Cells(iExcelOffset, 7) = 25
                                    thisAnimalWorksheet.Cells(iExcelOffset, 8).Value = sourceWorksheet.Range("AF" & lRowNum).Value
                                    thisAnimalWorksheet.Cells(iExcelOffset, 9).Value = sourceWorksheet.Range("AG" & lRowNum).Value
                                    thisAnimalWorksheet.Cells(iExcelOffset, 21).Value = sourceWorksheet.Range("AT" & lRowNum).Value
                                    thisAnimalWorksheet.Cells(iExcelOffset, 25).Value = sourceWorksheet.Range("AX" & lRowNum).Value
                            Case 30:
                                    thisAnimalWorksheet.Cells(iExcelOffset, 10) = 30
                                    thisAnimalWorksheet.Cells(iExcelOffset, 11).Value = sourceWorksheet.Range("AF" & lRowNum).Value
                                    thisAnimalWorksheet.Cells(iExcelOffset, 12).Value = sourceWorksheet.Range("AG" & lRowNum).Value
                                    thisAnimalWorksheet.Cells(iExcelOffset, 22).Value = sourceWorksheet.Range("AT" & lRowNum).Value
                                    thisAnimalWorksheet.Cells(iExcelOffset, 26).Value = sourceWorksheet.Range("AX" & lRowNum).Value
                            End Select

                        lTotalAttn1420 = lTotalAttn1420 + thisAnimalWorksheet.Cells(iExcelOffset, 5).Value
                        lTotalAttn5820 = lTotalAttn5820 + thisAnimalWorksheet.Cells(iExcelOffset, 6).Value
                        dblTotal1420 = dblTotal1420 + thisAnimalWorksheet.Cells(iExcelOffset, 24).Value
                        dblTotal5820 = dblTotal5820 + thisAnimalWorksheet.Cells(iExcelOffset, 20).Value
                        
                        lTotalAttn1425 = lTotalAttn1425 + thisAnimalWorksheet.Cells(iExcelOffset, 8).Value
                        lTotalAttn5825 = lTotalAttn5825 + thisAnimalWorksheet.Cells(iExcelOffset, 9).Value
                        dblTotal1425 = dblTotal1425 + thisAnimalWorksheet.Cells(iExcelOffset, 25).Value
                        dblTotal5825 = dblTotal5825 + thisAnimalWorksheet.Cells(iExcelOffset, 21).Value
                        
                        lTotalAttn1430 = lTotalAttn1430 + thisAnimalWorksheet.Cells(iExcelOffset, 11).Value
                        lTotalAttn5830 = lTotalAttn5830 + thisAnimalWorksheet.Cells(iExcelOffset, 12).Value
                        dblTotal1430 = dblTotal1430 + thisAnimalWorksheet.Cells(iExcelOffset, 26).Value
                        dblTotal5830 = dblTotal5830 + thisAnimalWorksheet.Cells(iExcelOffset, 22).Value

'                        Select Case thisAnimalWorksheet.Cells(iExcelOffset, 4).Value
'                            Case 20:
'                                lTotalAttn1420 = lTotalAttn1420 + thisAnimalWorksheet.Cells(iExcelOffset, 5).Value
'                                lTotalAttn5820 = lTotalAttn5820 + thisAnimalWorksheet.Cells(iExcelOffset, 6).Value
'                                dblTotal1420 = dblTotal1420 + thisAnimalWorksheet.Cells(iExcelOffset, 24).Value
'                                dblTotal5820 = dblTotal5820 + thisAnimalWorksheet.Cells(iExcelOffset, 20).Value

'                            Case 25:
'                                lTotalAttn1425 = lTotalAttn1425 + thisAnimalWorksheet.Cells(iExcelOffset, 5).Value
'                                lTotalAttn5825 = lTotalAttn5825 + thisAnimalWorksheet.Cells(iExcelOffset, 6).Value
'                                dblTotal1425 = dblTotal1425 + thisAnimalWorksheet.Cells(iExcelOffset, 24).Value
'                                dblTotal5825 = dblTotal5825 + thisAnimalWorksheet.Cells(iExcelOffset, 20).Value

'                            Case 30:
'                                lTotalAttn1430 = lTotalAttn1430 + thisAnimalWorksheet.Cells(iExcelOffset, 5).Value
'                                lTotalAttn5830 = lTotalAttn5830 + thisAnimalWorksheet.Cells(iExcelOffset, 6).Value
'                                dblTotal1430 = dblTotal1430 + thisAnimalWorksheet.Cells(iExcelOffset, 24).Value
'                                dblTotal5830 = dblTotal5830 + thisAnimalWorksheet.Cells(iExcelOffset, 20).Value
'                        End Select
                            
'                        Select Case thisAnimalWorksheet.Cells(iExcelOffset, 7).Value
'                            Case 20:
'                                lTotalAttn1420 = lTotalAttn1420 + thisAnimalWorksheet.Cells(iExcelOffset, 8).Value
'                                lTotalAttn5820 = lTotalAttn5820 + thisAnimalWorksheet.Cells(iExcelOffset, 9).Value
'                                dblTotal1420 = dblTotal1420 + thisAnimalWorksheet.Cells(iExcelOffset, 25).Value
'                                dblTotal5820 = dblTotal5820 + thisAnimalWorksheet.Cells(iExcelOffset, 21).Value

'                            Case 25:
'                                lTotalAttn1425 = lTotalAttn1425 + thisAnimalWorksheet.Cells(iExcelOffset, 8).Value
'                                lTotalAttn5825 = lTotalAttn5825 + thisAnimalWorksheet.Cells(iExcelOffset, 9).Value
'                                dblTotal1425 = dblTotal1425 + thisAnimalWorksheet.Cells(iExcelOffset, 25).Value
'                                dblTotal5825 = dblTotal5825 + thisAnimalWorksheet.Cells(iExcelOffset, 21).Value

'                            Case 30:
'                                lTotalAttn1430 = lTotalAttn1430 + thisAnimalWorksheet.Cells(iExcelOffset, 8).Value
'                                lTotalAttn5830 = lTotalAttn5830 + thisAnimalWorksheet.Cells(iExcelOffset, 9).Value
'                                dblTotal1430 = dblTotal1430 + thisAnimalWorksheet.Cells(iExcelOffset, 25).Value
'                                dblTotal5830 = dblTotal5830 + thisAnimalWorksheet.Cells(iExcelOffset, 21).Value
'                        End Select
                            
'                        Select Case thisAnimalWorksheet.Cells(iExcelOffset, 10).Value
'                            Case 20:
'                                lTotalAttn1420 = lTotalAttn1420 + thisAnimalWorksheet.Cells(iExcelOffset, 11).Value
'                                lTotalAttn5820 = lTotalAttn5820 + thisAnimalWorksheet.Cells(iExcelOffset, 12).Value
'                                dblTotal1420 = dblTotal1420 + thisAnimalWorksheet.Cells(iExcelOffset, 26).Value
'                                dblTotal5820 = dblTotal5820 + thisAnimalWorksheet.Cells(iExcelOffset, 22).Value

'                            Case 25:
'                                lTotalAttn1425 = lTotalAttn1425 + thisAnimalWorksheet.Cells(iExcelOffset, 11).Value
'                                lTotalAttn5825 = lTotalAttn5825 + thisAnimalWorksheet.Cells(iExcelOffset, 12).Value
'                                dblTotal1425 = dblTotal1425 + thisAnimalWorksheet.Cells(iExcelOffset, 26).Value
'                                dblTotal5825 = dblTotal5825 + thisAnimalWorksheet.Cells(iExcelOffset, 22).Value

'                            Case 30:
'                                lTotalAttn1430 = lTotalAttn1430 + thisAnimalWorksheet.Cells(iExcelOffset, 11).Value
'                                lTotalAttn5830 = lTotalAttn5830 + thisAnimalWorksheet.Cells(iExcelOffset, 12).Value
'                                dblTotal1430 = dblTotal1430 + thisAnimalWorksheet.Cells(iExcelOffset, 26).Value
'                                dblTotal5830 = dblTotal5830 + thisAnimalWorksheet.Cells(iExcelOffset, 22).Value
'                        End Select
                            
                        
                            'Dim lTotalAttn1420 As Long
                            'Dim lTotalAttn1425 As Long
                            'Dim lTotalAttn1430 As Long
    
                            'Dim lTotalAttn5820 As Long
                            'Dim lTotalAttn5825 As Long
                            'Dim lTotalAttn5830 As Long
                        
                        
                    ElseIf sourceWorksheet.Range("G" & lRowNum).Value <> "" Then
                        nExcluded = nExcluded + 1
                        thisAnimalWorksheet.Cells(iExcelOffset, 19).Value = sourceWorksheet.Range("G" & lRowNum).Value
                        thisAnimalWorksheet.Range("A" & iExcelOffset, "AZ" & iExcelOffset).Interior.Color = excludedTrialCell.Interior.Color
                        thisAnimalWorksheet.Range("A" & iExcelOffset, "AZ" & iExcelOffset).Interior.ColorIndex = excludedTrialCell.Interior.ColorIndex
                        thisAnimalWorksheet.Range("A" & iExcelOffset, "AZ" & iExcelOffset).Font.Color = excludedTrialCell.Font.Color
                        thisAnimalWorksheet.Range("A" & iExcelOffset, "AZ" & iExcelOffset).Font.ColorIndex = excludedTrialCell.Font.ColorIndex
                    End If
                'End If
                iExcelOffset = iExcelOffset + 1
            Next
    
'            If nInMeanSoFar > 0 Then
'                changeMean = changeSum / nInMeanSoFar
'            End If
    
 '           If nInMeanSoFar > 1 Then
 '               changeVar = (changeSumSqr - (changeSum ^ 2 / nInMeanSoFar)) / (nInMeanSoFar - 1)
 '               If changeVar <> 0 Then
 '                   tStat = (changeMean / (changeVar ^ 0.5) / (nInMeanSoFar ^ 0.5))
 '               Else
 '                   tStat = 10000
 '               End If
 '           End If
            
            iExcelOffset = iExcelOffset + 1
            
            thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 2) = nInMeanSoFar
            thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 3) = nExcluded
            
            
            thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "N included:"
            thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Italic = True
            thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = nInMeanSoFar
            iExcelOffset = iExcelOffset + 1
            thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "N excluded:"
            thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Italic = True
            thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = nExcluded
            iExcelOffset = iExcelOffset + 1
    
            If nInMeanSoFar > 0 Then
            
                lChartPos = iExcelOffset
                iExcelOffset = iExcelOffset + 1
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "1-4 total 20db:"
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Italic = True
                thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = dblTotal1420
                thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 5) = dblTotal1420
                
                thisAnimalWorksheet.Cells(iExcelOffset, 4).Value = "1-4 mean 20db:"
                thisAnimalWorksheet.Cells(iExcelOffset, 4).Font.Italic = True
                thisAnimalWorksheet.Cells(iExcelOffset, 5).Value = dblTotal1420 / lTotalAttn1420
                
                iExcelOffset = iExcelOffset + 1
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "1-4 total 25db:"
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Italic = True
                thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = dblTotal1425
                thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 6) = dblTotal1425
                
                thisAnimalWorksheet.Cells(iExcelOffset, 4).Value = "1-4 mean 25db:"
                thisAnimalWorksheet.Cells(iExcelOffset, 4).Font.Italic = True
                thisAnimalWorksheet.Cells(iExcelOffset, 5).Value = dblTotal1425 / lTotalAttn1425
                
                iExcelOffset = iExcelOffset + 1
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "1-4 total 30db:"
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Italic = True
                thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = dblTotal1430
                thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 7) = dblTotal1430
                
                thisAnimalWorksheet.Cells(iExcelOffset, 4).Value = "1-4 mean 30db:"
                thisAnimalWorksheet.Cells(iExcelOffset, 4).Font.Italic = True
                thisAnimalWorksheet.Cells(iExcelOffset, 5).Value = dblTotal1430 / lTotalAttn1430
                
                iExcelOffset = iExcelOffset + 2
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "5-8 total 20db:"
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Italic = True
                thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = dblTotal5820
                thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 8) = dblTotal5820
                
                thisAnimalWorksheet.Cells(iExcelOffset, 4).Value = "5-8 mean 20db:"
                thisAnimalWorksheet.Cells(iExcelOffset, 4).Font.Italic = True
                thisAnimalWorksheet.Cells(iExcelOffset, 5).Value = dblTotal5820 / lTotalAttn5820
                
                iExcelOffset = iExcelOffset + 1
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "5-8 total 25db:"
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Italic = True
                thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = dblTotal5825
                thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 9) = dblTotal5825
                
                thisAnimalWorksheet.Cells(iExcelOffset, 4).Value = "5-8 mean 25db:"
                thisAnimalWorksheet.Cells(iExcelOffset, 4).Font.Italic = True
                thisAnimalWorksheet.Cells(iExcelOffset, 5).Value = dblTotal5825 / lTotalAttn5825
                
                iExcelOffset = iExcelOffset + 1
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "5-8 total 30db:"
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Italic = True
                thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = dblTotal5830
                thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 10) = dblTotal5830
                
                thisAnimalWorksheet.Cells(iExcelOffset, 4).Value = "5-8 mean 30db:"
                thisAnimalWorksheet.Cells(iExcelOffset, 4).Font.Italic = True
                thisAnimalWorksheet.Cells(iExcelOffset, 5).Value = dblTotal5830 / lTotalAttn5830
                

                'Dim myChart As ChartObject
                Set myChart = thisAnimalWorksheet.ChartObjects.Add(thisAnimalWorksheet.Cells(lChartPos, 8).Left, thisAnimalWorksheet.Cells(lChartPos, 8).Top, 300, 200)
                myChart.Chart.ChartType = xlLine
                'myChart.Chart.SeriesCollection(1).Name
                myChart.Chart.SeriesCollection.NewSeries
                myChart.Chart.SeriesCollection(1).Name = "Stim 1-4"
                myChart.Chart.SeriesCollection(1).Format.Line.Weight = 1#
                myChart.Chart.SeriesCollection(1).XValues = Array(20, 25, 30)
                'myChart.Chart.Legend.Delete
                myChart.Chart.SeriesCollection(1).Values = thisAnimalWorksheet.Range("$E$" & lChartPos + 1 & ":$E$" & lChartPos + 3)
'                myChart.Chart.ChartType = xlLine
                myChart.Chart.SeriesCollection.NewSeries
                myChart.Chart.SeriesCollection(2).Name = "Stim 5-8"
                myChart.Chart.SeriesCollection(2).Format.Line.Weight = 1#
                myChart.Chart.SeriesCollection(2).XValues = Array(20, 25, 30)
                myChart.Chart.SeriesCollection(2).Values = thisAnimalWorksheet.Range("$E$" & lChartPos + 4 & ":$E$" & lChartPos + 6)

                iExcelOffset = iExcelOffset + 2
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "1-4 20db count:"
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Italic = True
                thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = lTotalAttn1420
                thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 12) = lTotalAttn1420

                iExcelOffset = iExcelOffset + 1
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "1-4 25db count:"
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Italic = True
                thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = lTotalAttn1425
                thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 13) = lTotalAttn1425
                
                iExcelOffset = iExcelOffset + 1
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "1-4 30db count:"
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Italic = True
                thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = lTotalAttn1430
                thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 14) = lTotalAttn1430
                iExcelOffset = iExcelOffset + 1
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "1-4 total count:"
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Italic = True
                thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = lTotalAttn1430 + lTotalAttn1425 + lTotalAttn1420
                thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 17) = lTotalAttn1430 + lTotalAttn1425 + lTotalAttn1420


                iExcelOffset = iExcelOffset + 2
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "5-8 20db count:"
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Italic = True
                thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = lTotalAttn5820
                thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 15) = lTotalAttn5820
                iExcelOffset = iExcelOffset + 1
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "5-8 25db count:"
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Italic = True
                thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = lTotalAttn5825
                thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 16) = lTotalAttn5825
                iExcelOffset = iExcelOffset + 1
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "5-8 30db count:"
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Italic = True
                thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = lTotalAttn5830
                thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 17) = lTotalAttn5830
                iExcelOffset = iExcelOffset + 1
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "5-8 total count:"
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Italic = True
                thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = lTotalAttn5830 + lTotalAttn5825 + lTotalAttn5820
                thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 17) = lTotalAttn5830 + lTotalAttn5825 + lTotalAttn5820

                If nInMeanSoFar > 1 Then
'                    thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "Variance:"
'                    thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = changeVar
'                    iExcelOffset = iExcelOffset + 1
                    'thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "Standard Deviation:"
                    'thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = changeVar ^ 0.5
                    'thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 6).Value = changeVar ^ 0.5
                    'iExcelOffset = iExcelOffset + 1
                    'thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "Std. Error of Mean:"
                    'thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = ((changeVar / nInMeanSoFar) ^ 0.5)
                    'iExcelOffset = iExcelOffset + 1
                    'thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "T-statistic:"
                    'thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = tStat
                    'thisAnimalWorksheet.Cells(iExcelOffset, 2).NumberFormat = "0.000"
                    'thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 7).Value = tStat
                    'thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 7).NumberFormat = "0.000"
                    'iExcelOffset = iExcelOffset + 1
                    'thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "P-value:"
                    'thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Italic = True
                    'thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = "=TDIST(ABS(B" & CStr(iExcelOffset - 1) & ")," & CStr(nInMeanSoFar - 1) & ",1)"
                    'Call thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions.Delete
                    'Call thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions.Add(xlCellValue, xlLessEqual, ".05")
                    'thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(1).Font.Color = pLess05FC.Font.Color
                    'thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(1).Font.ColorIndex = pLess05FC.Font.ColorIndex
                    'thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(1).Interior.Color = pLess05FC.Interior.Color
                    'thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(1).Interior.ColorIndex = pLess05FC.Interior.ColorIndex
                    'Call thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions.Add(xlCellValue, xlLessEqual, ".1")
                    'thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(2).Font.Color = pLess10FC.Font.Color
                    'thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(2).Font.ColorIndex = pLess10FC.Font.ColorIndex
                    'thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(2).Interior.Color = pLess10FC.Interior.Color
                    'thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(2).Interior.ColorIndex = pLess10FC.Interior.ColorIndex
                    'thisAnimalWorksheet.Cells(iExcelOffset, 2).NumberFormat = "0.000"
                    
                    'thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 8).Value = "=TDIST(ABS(" & thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 7).Address & ")," & CStr(nInMeanSoFar - 1) & ",1)"
                    'Call thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 8).FormatConditions.Delete
                    'Call thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 8).FormatConditions.Add(xlCellValue, xlLessEqual, ".05")
                    'thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 8).FormatConditions(1).Font.Color = pLess05FC.Font.Color
                    'thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 8).FormatConditions(1).Font.ColorIndex = pLess05FC.Font.ColorIndex
                    'thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 8).FormatConditions(1).Interior.Color = pLess05FC.Interior.Color
                    'thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 8).FormatConditions(1).Interior.ColorIndex = pLess05FC.Interior.ColorIndex
                    'Call thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 8).FormatConditions.Add(xlCellValue, xlLessEqual, ".1")
                    'thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 8).FormatConditions(2).Font.Color = pLess10FC.Font.Color
                    'thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 8).FormatConditions(2).Font.ColorIndex = pLess10FC.Font.ColorIndex
                    'thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 8).FormatConditions(2).Interior.Color = pLess10FC.Interior.Color
                    'thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 8).FormatConditions(2).Interior.ColorIndex = pLess10FC.Interior.ColorIndex
                    'thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 8).NumberFormat = "0.000"
                Else
                    thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "Additional stats could not be calculated (N=1)"
                    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 6) = "=NA()"
                    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 7) = "=NA()"
                    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 8) = "=NA()"
                End If
            End If
            
            iExcelOffset = iExcelOffset + 2
            

                
            thisAnimalSummarySheetRow = thisAnimalSummarySheetRow + 1
        Next
        iExcelOffset = iExcelOffset + 1
    Next
    
    thisAnimalSummarySheetRow = thisAnimalSummarySheetRow + 2
End Sub


Sub deleteOldWorksheets(thisWorkbook As Workbook)
    Dim i As Integer
    i = 1
    While thisWorkbook.Worksheets.Count > 2
        If thisWorkbook.Worksheets(i).Name <> "Controller" And thisWorkbook.Worksheets(i).Name <> "Output template" And thisWorkbook.Worksheets(i).Name <> "Trials" Then
            thisWorkbook.Worksheets(i).Delete
        Else
            i = i + 1
        End If
    Wend
End Sub

Sub setUpStyles(thisWorkbook As Workbook)
    Dim currStyle As Style
    Set currStyle = Nothing
    
    On Error Resume Next
    Set currStyle = thisWorkbook.Styles("Heading")
    On Error GoTo 0
    If currStyle Is Nothing Then
        Call thisWorkbook.Styles.Add("Head", thisWorkbook.Styles("Normal"))
    End If
    
    thisWorkbook.Styles("Heading").Font.Bold = True
    
    On Error Resume Next
    Set currStyle = thisWorkbook.Styles("ParamSet")
    On Error GoTo 0
    If currStyle Is Nothing Then
        Call thisWorkbook.Styles.Add("ParamSet", thisWorkbook.Styles("Normal"))
    End If
    
    thisWorkbook.Styles("ParamSet").Font.Bold = True
    
End Sub

Function checkForExclusion(objFolder As Folder) As Variant
    Dim exclusionInfo(2) As String
    
    exclusionInfo(0) = ""
    checkForExclusion = False
    Dim Files As Files
    Dim objFile As File

    Set Files = objFolder.Files

    Dim tmpStr1 As String
    Dim tmpStr2 As String
    Dim iLenOfPrefix As Integer
    iLenOfPrefix = Len("exclude from neural data - ")

    For Each objFile In Files
        If LCase(objFile.Name) = "exclude from neural data.txt" Then
            exclusionInfo(0) = "folder"
            Exit For
        ElseIf LCase(Left(objFile.Name, iLenOfPrefix)) = "exclude from neural data - " Then
            exclusionInfo(0) = "neural"
            'exclude from results aggregration - all.txt
            tmpStr1 = Right(LCase(objFile.Name), Len(objFile.Name) - iLenOfPrefix)
            tmpStr2 = Left(tmpStr1, Len(tmpStr1) - 4)
            
            If LCase(Left(tmpStr2, Len("partial"))) = "partial" Then
               exclusionInfo(2) = readPartialFromFile(objFile)
               exclusionInfo(1) = readCommentFromFile(objFile)
               'tmpStr2 = Right(tmpStr2, Len(tmpStr2) - Len("partial") - 1)
            End If
            
            'If LCase(tmpStr2) = "with message" Then
                'exclusionInfo(1) = readCommentFromFile(objFile)
            'End If
            
'            Select Case tmpStr2
'                Case "all":
'                    exclusionInfo(0) = "all"
'                Case "all with message":
'                    exclusionInfo(0) = "all"
'                    exclusionInfo(1) = readCommentFromFile(objFile)
'                Case "acoustic":
'                    exclusionInfo(0) = "Acoustic"
'                Case "acoustic with message":
'                    exclusionInfo(0) = "Acoustic"
'                    exclusionInfo(1) = readCommentFromFile(objFile)
'                Case "electical":
'                    exclusionInfo(0) = "Electrical"
'                Case "electrical with message":
'                    exclusionInfo(0) = "Electrical"
'                    exclusionInfo(1) = readCommentFromFile(objFile)
'                'exclude from neural data - partial electrical with message.txt
'            End Select
            Exit For
        End If
    Next
    
    checkForExclusion = exclusionInfo

End Function

Function readCommentFromFile(objFile As File) As String
    Dim ts As TextStream
    Set ts = objFile.OpenAsTextStream
    Dim strBuffer As String
    
    Do
        If ts.AtEndOfStream Then
            Exit Do
        End If
        
        strBuffer = ts.ReadLine
        
        If Not LCase(Left(strBuffer, Len("exclude after:"))) = "exclude after:" Then
            readCommentFromFile = strBuffer
            Exit Do
        End If
    Loop
    ts.Close
End Function

Function readPartialFromFile(objFile As File) As String
    Dim ts As TextStream
    Set ts = objFile.OpenAsTextStream
    Dim strBuffer As String
    
    Do
        If ts.AtEndOfStream Then
            Exit Do
        End If
        
        strBuffer = ts.ReadLine
        
        If LCase(Left(strBuffer, Len("exclude after:"))) = "exclude after:" Then
            'While not Exclude after: 4800
            readPartialFromFile = Right(strBuffer, Len(strBuffer) - Len("exclude after:"))
            Exit Do
        End If
    Loop
    ts.Close
End Function

Function addToDict(ByRef objDict As Dictionary, entryInfo As String, chanNum As String, iRow As Integer)
    Dim paramArr As Variant
    Dim iParamOffset As Integer
    
    If Not objDict.Exists(chanNum) Then
        Dim newDict As Dictionary
        Set newDict = New Dictionary
        Call objDict.Add(chanNum, newDict)
    End If
    
    If Not objDict(chanNum).Exists(entryInfo) Then
        Call objDict(chanNum).Add(entryInfo, Array())
    End If
                     
    paramArr = objDict(chanNum)(entryInfo)
                     
    ReDim Preserve paramArr(UBound(paramArr) + 1)
    iParamOffset = UBound(paramArr)
    paramArr(iParamOffset) = iRow
                     
    objDict(chanNum)(entryInfo) = paramArr
End Function
