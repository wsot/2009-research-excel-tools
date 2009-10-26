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
    
    Dim iColHeadersForHRLine As Integer
        
        For iSourceWorksheetOffset = 1 To (thisWorkbook.Worksheets.Count)
            If thisWorkbook.Worksheets(iSourceWorksheetOffset).Name <> "Controller" And thisWorkbook.Worksheets(iSourceWorksheetOffset).Name <> "Trials" Then 'check if this is actually a data sheet
                Set sourceWorksheet = thisWorkbook.Worksheets(iSourceWorksheetOffset)
                
                Set neuralByDate = New Dictionary
                Set neuralByAcclim = New Dictionary

                validTrialCount = 0
                Set thisAnimalWorksheet = Nothing
                Set thisAnimalSummarySheet = Nothing
                animalID = sourceWorksheet.Name
                
                Call parseTrials(sourceWorksheet)
                For iPass = 0 To 1
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
        If Not outputByDate Is Nothing Then
            outputFilename = pathToData & "\neural aggregate by date.xlsx"
            Call outputByDate.SaveAs(outputFilename)
            Call outputByDate.Close
        End If
        If Not outputByAcclim Is Nothing Then
            outputFilename = pathToData & "\aggregate by acclim.xlsx"
            Call outputByAcclim.SaveAs(outputFilename)
            Call outputByAcclim.Close
        End If
    
    Set objFS = Nothing
    
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic

End Sub

'Function parseTrials(outputDict As Dictionary, sourceWorksheet As Workbook, experimentDate As String, experimentTag As String, exclusionInfo As Variant)
Function parseTrials(sourceWorksheet As Worksheet)
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
        
        Call addToDict(neuralByDate, experimentDate, iChannel, i)
        Call addToDict(neuralByAcclim, neuralByAcclimInfo, iChannel, i)
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
        
    Dim iSummaryCol As Integer

    For iChanNum = 0 To UBound(arrChannels)
        thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "Channel " & arrChannels(iChanNum)

        thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Bold = True
        iExcelOffset = iExcelOffset + 1
    
        Set dictSublevel = theDict(arrChannels(iChanNum))
        arrSubLevels = dictSublevel.Keys
        For iSubLevelNum = 0 To UBound(arrSubLevels)
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
                
                thisAnimalWorksheet.Cells(iExcelOffset, 4).Value = sourceWorksheet.Range("Y" & lRowNum).Value
                thisAnimalWorksheet.Cells(iExcelOffset, 5).Value = sourceWorksheet.Range("Z" & lRowNum).Value
                thisAnimalWorksheet.Cells(iExcelOffset, 6).Value = sourceWorksheet.Range("AA" & lRowNum).Value
                thisAnimalWorksheet.Cells(iExcelOffset, 7).Value = sourceWorksheet.Range("AB" & lRowNum).Value
                thisAnimalWorksheet.Cells(iExcelOffset, 8).Value = sourceWorksheet.Range("AC" & lRowNum).Value
                thisAnimalWorksheet.Cells(iExcelOffset, 9).Value = sourceWorksheet.Range("AD" & lRowNum).Value
                thisAnimalWorksheet.Cells(iExcelOffset, 10).Value = sourceWorksheet.Range("AE" & lRowNum).Value
                thisAnimalWorksheet.Cells(iExcelOffset, 11).Value = sourceWorksheet.Range("AF" & lRowNum).Value
                thisAnimalWorksheet.Cells(iExcelOffset, 12).Value = sourceWorksheet.Range("AG" & lRowNum).Value
                
                thisAnimalWorksheet.Cells(iExcelOffset, 13).Value = sourceWorksheet.Range("AM" & lRowNum).Value
                thisAnimalWorksheet.Cells(iExcelOffset, 14).Value = sourceWorksheet.Range("AN" & lRowNum).Value
                
                thisAnimalWorksheet.Cells(iExcelOffset, 15).Value = sourceWorksheet.Range("AK" & lRowNum).Value
                thisAnimalWorksheet.Cells(iExcelOffset, 16).Value = sourceWorksheet.Range("AL" & lRowNum).Value
    
                thisAnimalWorksheet.Cells(iExcelOffset, 17).Value = sourceWorksheet.Range("AO" & lRowNum).Value
                thisAnimalWorksheet.Cells(iExcelOffset, 18).Value = sourceWorksheet.Range("AP" & lRowNum).Value
    
                If sourceWorksheet.Range("G" & lRowNum).Value = "" Or (sourceWorksheet.Range("G" & lRowNum).Value <> "" And sourceWorksheet.Range("H" & lRowNum).Value >= sourceWorksheet.Range("K" & lRowNum).Value) Then 'check if the data should be excluded
                    nInMeanSoFar = nInMeanSoFar + 1
                    diff = thisAnimalWorksheet.Cells(iExcelOffset, 14).Value - thisAnimalWorksheet.Cells(iExcelOffset, 13).Value
                    changeSum = changeSum + diff
                    changeSumSqr = changeSumSqr + diff ^ 2
                ElseIf sourceWorksheet.Range("G" & lRowNum).Value <> "" Then
                    nExcluded = nExcluded + 1
                    thisAnimalWorksheet.Cells(iExcelOffset, 19).Value = sourceWorksheet.Range("G" & lRowNum).Value
                    thisAnimalWorksheet.Range("A" & iExcelOffset, "AZ" & iExcelOffset).Interior.Color = excludedTrialCell.Interior.Color
                    thisAnimalWorksheet.Range("A" & iExcelOffset, "AZ" & iExcelOffset).Interior.ColorIndex = excludedTrialCell.Interior.ColorIndex
                    thisAnimalWorksheet.Range("A" & iExcelOffset, "AZ" & iExcelOffset).Font.Color = excludedTrialCell.Font.Color
                    thisAnimalWorksheet.Range("A" & iExcelOffset, "AZ" & iExcelOffset).Font.ColorIndex = excludedTrialCell.Font.ColorIndex
                End If
    
                iExcelOffset = iExcelOffset + 1
            Next
    
            If nInMeanSoFar > 0 Then
                changeMean = changeSum / nInMeanSoFar
            End If
    
            If nInMeanSoFar > 1 Then
                changeVar = (changeSumSqr - (changeSum ^ 2 / nInMeanSoFar)) / (nInMeanSoFar - 1)
                If changeVar <> 0 Then
                    tStat = (changeMean / (changeVar ^ 0.5) / (nInMeanSoFar ^ 0.5))
                Else
                    tStat = 10000
                End If
            End If
            
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
                iExcelOffset = iExcelOffset + 1
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "Mean change:"
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Italic = True
                thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = changeMean
                thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 5) = changeMean
                iExcelOffset = iExcelOffset + 1
                If nInMeanSoFar > 1 Then
                    thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "Variance:"
                    thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = changeVar
                    iExcelOffset = iExcelOffset + 1
                    thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "Standard Deviation:"
                    thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = changeVar ^ 0.5
                    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 6).Value = changeVar ^ 0.5
                    iExcelOffset = iExcelOffset + 1
                    thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "Std. Error of Mean:"
                    thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = ((changeVar / nInMeanSoFar) ^ 0.5)
                    iExcelOffset = iExcelOffset + 1
                    thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "T-statistic:"
                    thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = tStat
                    thisAnimalWorksheet.Cells(iExcelOffset, 2).NumberFormat = "0.000"
                    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 7).Value = tStat
                    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 7).NumberFormat = "0.000"
                    iExcelOffset = iExcelOffset + 1
                    thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "P-value:"
                    thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Italic = True
                    thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = "=TDIST(ABS(B" & CStr(iExcelOffset - 1) & ")," & CStr(nInMeanSoFar - 1) & ",1)"
                    Call thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions.Delete
                    Call thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions.Add(xlCellValue, xlLessEqual, ".05")
                    thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(1).Font.Color = pLess05FC.Font.Color
                    thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(1).Font.ColorIndex = pLess05FC.Font.ColorIndex
                    thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(1).Interior.Color = pLess05FC.Interior.Color
                    thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(1).Interior.ColorIndex = pLess05FC.Interior.ColorIndex
                    Call thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions.Add(xlCellValue, xlLessEqual, ".1")
                    thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(2).Font.Color = pLess10FC.Font.Color
                    thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(2).Font.ColorIndex = pLess10FC.Font.ColorIndex
                    thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(2).Interior.Color = pLess10FC.Interior.Color
                    thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(2).Interior.ColorIndex = pLess10FC.Interior.ColorIndex
                    thisAnimalWorksheet.Cells(iExcelOffset, 2).NumberFormat = "0.000"
                    
                    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 8).Value = "=TDIST(ABS(" & thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 7).Address & ")," & CStr(nInMeanSoFar - 1) & ",1)"
                    Call thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 8).FormatConditions.Delete
                    Call thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 8).FormatConditions.Add(xlCellValue, xlLessEqual, ".05")
                    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 8).FormatConditions(1).Font.Color = pLess05FC.Font.Color
                    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 8).FormatConditions(1).Font.ColorIndex = pLess05FC.Font.ColorIndex
                    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 8).FormatConditions(1).Interior.Color = pLess05FC.Interior.Color
                    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 8).FormatConditions(1).Interior.ColorIndex = pLess05FC.Interior.ColorIndex
                    Call thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 8).FormatConditions.Add(xlCellValue, xlLessEqual, ".1")
                    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 8).FormatConditions(2).Font.Color = pLess10FC.Font.Color
                    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 8).FormatConditions(2).Font.ColorIndex = pLess10FC.Font.ColorIndex
                    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 8).FormatConditions(2).Interior.Color = pLess10FC.Interior.Color
                    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 8).FormatConditions(2).Interior.ColorIndex = pLess10FC.Interior.ColorIndex
                    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 8).NumberFormat = "0.000"
                Else
                    thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "Additional stats could not be calculated (N=1)"
                    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 6) = "=NA()"
                    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 7) = "=NA()"
                    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 8) = "=NA()"
                End If
            End If
            
            iExcelOffset = iExcelOffset + 2
            
            If 0 = 1 Then
                Dim myChart As ChartObject
                Dim chartOffset As Integer
                Dim chartHeight As Integer
                If iThisAnimalSummarySheetStartingRow > 2 Then
                    'chartOffset = (iThisAnimalSummarySheetStartingRow) * 15.5 + (UBound(arrParamSets) + 2) * 15.5
                    chartOffset = thisAnimalSummarySheet.Range("A" & iThisAnimalSummarySheetStartingRow + UBound(arrSubLevels) + 7 & ":A" & "A" & iThisAnimalSummarySheetStartingRow + UBound(arrSubLevels) + 7 + 19).Top
                    chartHeight = thisAnimalSummarySheet.Range("A" & iThisAnimalSummarySheetStartingRow + UBound(arrSubLevels) + 7 & ":A" & "A" & iThisAnimalSummarySheetStartingRow + UBound(arrSubLevels) + 7 + 19).Height
                Else
                    chartOffset = thisAnimalSummarySheet.Range("A" & UBound(arrSubLevels) + 7 & ":A" & UBound(arrSubLevels) + 7 + 19).Top
                    chartHeight = thisAnimalSummarySheet.Range("A" & UBound(arrSubLevels) + 7 & ":A" & UBound(arrSubLevels) + 7 + 19).Height
                    'chartOffset = (UBound(arrParamSets) + 5) * 15.5
                End If
    
                Set myChart = thisAnimalSummarySheet.ChartObjects.Add(((thisAnimalSummarySheetRow - iThisAnimalSummarySheetStartingRow) * 500) + 1, chartOffset, 500, chartHeight)
                myChart.Chart.ChartType = xlLine
                myChart.Chart.SeriesCollection.NewSeries
                myChart.Chart.SeriesCollection(1).Name = thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 1).Value & " (N=" & nInMeanSoFar & ")"
                myChart.Chart.SeriesCollection(1).Format.Line.Weight = 1#
                myChart.Chart.SeriesCollection(1).XValues = thisAnimalSummarySheet.Range("=$U$1:$EU$1")
                myChart.Chart.Legend.Delete
                myChart.Chart.SeriesCollection(1).Values = thisAnimalSummarySheet.Range("$U$" & thisAnimalSummarySheetRow & ":$EU$" & thisAnimalSummarySheetRow)
                myChart.Chart.SeriesCollection(1).HasErrorBars = True
                '1.96 Standard deviation
                'myChart.Chart.SeriesCollection(1).ErrorBar Direction:=xlY, Include:=xlBoth, _
                '    Type:=xlErrorBarTypeCustom, Amount:=thisAnimalSummarySheet.Range("$PC$" & thisAnimalSummarySheetRow & ":$UC$" & thisAnimalSummarySheetRow), MinusValues:=thisAnimalSummarySheet.Range("$PC$" & thisAnimalSummarySheetRow & ":$UC$" & thisAnimalSummarySheetRow)
    
                '1 Standard deviation
    '                myChart.Chart.SeriesCollection(1).ErrorBar Direction:=xlY, Include:=xlBoth, _
    '                    Type:=xlErrorBarTypeCustom, Amount:=thisAnimalSummarySheet.Range("$EW$" & thisAnimalSummarySheetRow & ":$JW$" & thisAnimalSummarySheetRow), MinusValues:=thisAnimalSummarySheet.Range("$EW$" & thisAnimalSummarySheetRow & ":$JW$" & thisAnimalSummarySheetRow)
                '2 SEM
                                myChart.Chart.SeriesCollection(1).ErrorBar Direction:=xlY, Include:=xlBoth, _
                    Type:=xlErrorBarTypeCustom, Amount:=thisAnimalSummarySheet.Range("$JZ$" & thisAnimalSummarySheetRow & ":$OZ$" & thisAnimalSummarySheetRow), MinusValues:=thisAnimalSummarySheet.Range("$JZ$" & thisAnimalSummarySheetRow & ":$OZ$" & thisAnimalSummarySheetRow)
    
    
                myChart.Chart.ChartTitle.Characters.Font.Size = 12
                myChart.Chart.Axes(xlValue).MinimumScale = 0.85
                myChart.Chart.Axes(xlValue).MaximumScale = 1.15
            End If
                
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

Function addToDict(ByRef objDict As Dictionary, entryInfo As String, chanNum As Integer, iRow As Integer)
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
