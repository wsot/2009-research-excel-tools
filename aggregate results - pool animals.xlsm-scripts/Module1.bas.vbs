Attribute VB_Name = "Module1"
Option Explicit

Const HRDetOffset = 24
Const HRDetCols = 18
Const HRDetHROffset = 5
Const HRDetInterpProp = 10
Const HRDetLongIntSamp = 12
Const HRDetLongIntBeat = 14
Const HRDetStdDev = 7

Global maxPercOfBeatsInt As Double
Global maxSingleIntSamples As Double
Global maxSingleIntBeats As Double
    
'Global exIntCountGT As Integer
'Global exIntBeatsGT As Integer
'Global exLongestIntDurGT As Integer
'Global exLongestIntBeatsGT As Integer

Global pLess05FC As FormatCondition
Global pLess10FC As FormatCondition
Global percOutside1585FC As FormatCondition
Global percOutside2575FC As FormatCondition
Global excludedTrialCell As Range

Dim allTrials() As Variant

Dim trialTypesByDate As Dictionary
Dim trialTypesByStimParamsFull As Dictionary
Dim trialTypesByDateStimParamsFull As Dictionary
Dim trialTypesByStimParamsNoAmp As Dictionary


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
    
    maxPercOfBeatsInt = thisWorkbook.Worksheets("Controller").Cells(3, 2).Value
    maxSingleIntSamples = thisWorkbook.Worksheets("Controller").Cells(4, 2).Value
    maxSingleIntBeats = thisWorkbook.Worksheets("Controller").Cells(5, 2).Value

    
    Set objFS = CreateObject("Scripting.FileSystemObject")
    
    pathToData = objFS.GetDriveName(thisWorkbook.FullName) & thisWorkbook.Worksheets("Controller").Cells(19, 2).Value
    Set rootFolder = objFS.GetFolder(pathToData)
    
    blnCurrFolderIsTrial = False
        
    Call deleteOldWorksheets(thisWorkbook)
    
    Set AnimalFolders = rootFolder.Subfolders
    For Each objAnimalFolder In AnimalFolders 'cycle through the folder for each animal
        exclusionInfo = checkForExclusion(objAnimalFolder)
        If Not exclusionInfo(0) = "folder" Then
            thisAnimalTrialsRow = 3
            Set thisAnimalWorksheet = Nothing
            animalID = objAnimalFolder.Name
                        
            Set experimentFolders = objAnimalFolder.Subfolders
            For Each objExpFolder In experimentFolders 'go through the experiments within an animal folder
                blnCurrFolderIsTrial = False
                exclusionInfo = checkForExclusion(objExpFolder)
                If exclusionInfo(1) <> "" Or exclusionInfo(0) <> "all" Or exclusionInfo(2) <> "" Then 'check if the exclusion includes a message, or is only for some types of trial
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
    Dim iSourceRow As Long
    iSourceRow = 2
    While workbookToProcess.Worksheets("Output").Cells(iSourceRow, 1) <> ""
        thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 1).Value = strExcelPathname
        thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 2).Value = workbookToProcess.Worksheets("Variables (do not edit)").Range("B2").Value
        thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 3).Value = workbookToProcess.Worksheets("Variables (do not edit)").Range("B3").Value
        thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 4).Value = experimentDate
        thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 5).Value = experimentTag
        thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 6).Value = exclusionInfo(0)
        thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 7).Value = exclusionInfo(1)
        thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 8).Value = exclusionInfo(2)
        thisAnimalWorksheet.Range(thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 9), thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 22)).Value = workbookToProcess.Worksheets("Output").Range("A" & iSourceRow & ":N" & iSourceRow).Value
        thisAnimalWorksheet.Range(thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 24), thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 94)).Value = workbookToProcess.Worksheets("HR detection").Range("A" & iSourceRow + 1 & ":BS" & iSourceRow + 1).Value
        thisAnimalWorksheet.Range(thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 102), thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 232)).Value = workbookToProcess.Worksheets("HRLine").Range("B" & iSourceRow & ":EB" & iSourceRow).Value
                
        exclusionReason = checkForHRExclusions(thisAnimalWorksheet, CInt(thisAnimalTrialsRow), HRDetOffset)
        If exclusionReason <> "" Then
            thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 96).Value = exclusionReason
        End If
        exclusionReason = checkForHRExclusions(thisAnimalWorksheet, CInt(thisAnimalTrialsRow), HRDetOffset + HRDetCols)
        If exclusionReason <> "" Then
            thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 97).Value = exclusionReason
        End If
        exclusionReason = checkForHRExclusions(thisAnimalWorksheet, CInt(thisAnimalTrialsRow), HRDetOffset + (2 * HRDetCols))
        If exclusionReason <> "" Then
            thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 98).Value = exclusionReason
        End If
        exclusionReason = checkForHRExclusions(thisAnimalWorksheet, CInt(thisAnimalTrialsRow), HRDetOffset + (3 * HRDetCols))
        If exclusionReason <> "" Then
            thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 99).Value = exclusionReason
        End If
        
        'convert when the stimulation cable was detached to being 'no stimulation' trials for comparison
        If exclusionInfo(0) = "Electrical" Then
            If thisAnimalWorksheet.Range("M" & thisAnimalTrialsRow).Value = "Electrical" Then
                If InStr(1, LCase(exclusionInfo(1)), "stimulation cable may have been detached", vbTextCompare) Then
                    thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 6).Value = ""
                    thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 7).Value = ""
                    thisAnimalWorksheet.Cells(thisAnimalTrialsRow, 8).Value = ""
                    
                    thisAnimalWorksheet.Range("N" & thisAnimalTrialsRow).Value = "0* ref 0* @ 400Hz"
                    thisAnimalWorksheet.Range("O" & thisAnimalTrialsRow).Value = "0uA"
                    thisAnimalWorksheet.Range("P" & thisAnimalTrialsRow).Value = "0uA"
                    thisAnimalWorksheet.Range("Q" & thisAnimalTrialsRow).Value = "0uA"
                    
                    thisAnimalWorksheet.Range("R" & thisAnimalTrialsRow).Value = "0* ref 0* @ 400Hz"
                    thisAnimalWorksheet.Range("S" & thisAnimalTrialsRow).Value = "0uA"
                    thisAnimalWorksheet.Range("T" & thisAnimalTrialsRow).Value = "0uA"
                    thisAnimalWorksheet.Range("U" & thisAnimalTrialsRow).Value = "0uA"
                End If
            End If
        End If
        
        thisAnimalTrialsRow = thisAnimalTrialsRow + 1
        iSourceRow = iSourceRow + 1

    Wend
End Function

Sub processTrials()
    Dim exclusionInfo As Variant
    Dim oneAnimalOneSheet As Boolean

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
    
    'get the root folder under which all data is housed
'    Dim rootFolder As Folder
'    Set rootFolder = objFS.GetFolder(objFS.GetFolder(objFS.GetParentFolderName(ActiveWorkbook.FullName)))
    
'    Dim AnimalFolders As Folders
'    Dim objAnimalFolder As Folder
    
'    Dim experimentFolders As Folders
'    Dim objExpFolder As Folder
    
'    Dim Files As Files
'    Dim objFile As File
    
'    Dim strExcelFilename As String
'    Dim strExcelPathname As String
    
'    Dim blnCurrFolderIsTrial As Boolean
    Dim trialTypes As Dictionary
    
    Dim thisAnimalWorksheet As Worksheet
    Dim thisAnimalSummarySheet As Worksheet
    Dim thisAnimalSummarySheetRow As Long
    Dim outputWorkbook As Workbook
    Dim workbookToProcess As Workbook
    
    Dim outputFilename As String
        
    Set thisWorkbook = ActiveWorkbook
    
    templateFilename = "\Code current\Excel tools\aggregate results output.xltm"
    Set objFS = CreateObject("Scripting.FileSystemObject")
    templateFilename = objFS.GetDriveName(thisWorkbook.FullName) & templateFilename 'get the drive letter for the template
    
    pathToData = objFS.GetDriveName(thisWorkbook.FullName) & thisWorkbook.Worksheets("Controller").Cells(19, 2).Value
'    Set rootFolder = objFS.GetFolder(pathToData)
        
    oneAnimalOneSheet = thisWorkbook.Worksheets("Controller").Cells(9, 2).Value
    
    maxPercOfBeatsInt = thisWorkbook.Worksheets("Controller").Cells(3, 2).Value
    maxSingleIntSamples = thisWorkbook.Worksheets("Controller").Cells(4, 2).Value
    maxSingleIntBeats = thisWorkbook.Worksheets("Controller").Cells(5, 2).Value

    'exIntCountGT = CInt(thisWorkbook.Worksheets("Controller").Cells(3, 2).Value)
    'exIntBeatsGT = CInt(thisWorkbook.Worksheets("Controller").Cells(4, 2).Value)
    'exLongestIntDurGT = CInt(thisWorkbook.Worksheets("Controller").Cells(5, 2).Value)
    'exLongestIntBeatsGT = CInt(thisWorkbook.Worksheets("Controller").Cells(6, 2).Value)
    
    Set pLess05FC = thisWorkbook.Worksheets("Controller").Range("B11").FormatConditions(1)
    Set pLess10FC = thisWorkbook.Worksheets("Controller").Range("B12").FormatConditions(1)
    
    Set percOutside1585FC = thisWorkbook.Worksheets("Controller").Range("B14").FormatConditions(1)
    Set percOutside2575FC = thisWorkbook.Worksheets("Controller").Range("B15").FormatConditions(1)
    
    Set excludedTrialCell = thisWorkbook.Worksheets("Controller").Range("B17")
    
    Dim iSourceWorksheetOffset As Integer
    Dim sourceWorksheet As Worksheet
    
    Dim iPass As Integer
    
'    For iPass = 0 To 2
'        Select Case iPass
'            Case 0:
'                clusterByStimParams = True
'                clusterByDate = False
'                outputFilename = pathToData & "\aggregate by stim params.xlsx"
'            Case 1:
'                clusterByStimParams = False
'                clusterByDate = True
'                outputFilename = pathToData & "\aggregate by date.xlsx"
'            Case 2:
'                clusterByStimParams = True
'                clusterByDate = True
'                outputFilename = pathToData & "\aggregate by stim params and date.xlsx"
'            End Select
'        blnCurrFolderIsTrial = False
    
'        clusterByStimParams = thisWorkbook.Worksheets("Controller").Cells(20, 2).Value
'        clusterByDate = thisWorkbook.Worksheets("Controller").Cells(21, 2).Value
            
'        Call deleteOldWorksheets(thisWorkbook)
        
'        Set AnimalFolders = rootFolder.Subfolders
'        For Each objAnimalFolder In AnimalFolders 'cycle through the folder for each animal
'            exclusionInfo = checkForExclusion(objAnimalFolder)
'            If Not exclusionInfo(0) = "folder" Then
        
    Dim outputByStimParamsNoAmp As Workbook
    Dim outputByDateStimParamsFull As Workbook
    Dim outputByDate As Workbook
    Dim outputByStimParamsFull As Workbook
    
    Dim iColHeadersForHRLine As Integer
        
        For iSourceWorksheetOffset = 1 To (thisWorkbook.Worksheets.Count)
            If thisWorkbook.Worksheets(iSourceWorksheetOffset).Name <> "Controller" And thisWorkbook.Worksheets(iSourceWorksheetOffset).Name <> "Trials" Then 'check if this is actually a data sheet
                Set sourceWorksheet = thisWorkbook.Worksheets(iSourceWorksheetOffset)
                
                
                'Dim trialTypesByDate As Dictionary
                'Dim trialTypesByStimParamsFull As Dictionary
                'Dim trialTypesByDateStimParamsFull As Dictionary
                'Dim trialTypesByStimParamsNoAmp As Dictionary
                
                Set trialTypesByDate = New Dictionary
                Call trialTypesByDate.Add("Acclimatisation", New Dictionary)
                Call trialTypesByDate.Add("Acoustic", New Dictionary)
                Call trialTypesByDate.Add("Acoustic Only", New Dictionary)
                Call trialTypesByDate.Add("Electrical", New Dictionary)
                Call trialTypesByDate.Add("No Stim", New Dictionary)
                
                Set trialTypesByStimParamsFull = New Dictionary
                Call trialTypesByStimParamsFull.Add("Acclimatisation", New Dictionary)
                Call trialTypesByStimParamsFull.Add("Acoustic", New Dictionary)
                Call trialTypesByStimParamsFull.Add("Acoustic Only", New Dictionary)
                Call trialTypesByStimParamsFull.Add("Electrical", New Dictionary)
                Call trialTypesByStimParamsFull.Add("No Stim", New Dictionary)

                Set trialTypesByDateStimParamsFull = New Dictionary
                Call trialTypesByDateStimParamsFull.Add("Acclimatisation", New Dictionary)
                Call trialTypesByDateStimParamsFull.Add("Acoustic", New Dictionary)
                Call trialTypesByDateStimParamsFull.Add("Acoustic Only", New Dictionary)
                Call trialTypesByDateStimParamsFull.Add("Electrical", New Dictionary)
                Call trialTypesByDateStimParamsFull.Add("No Stim", New Dictionary)

                Set trialTypesByStimParamsNoAmp = New Dictionary
                Call trialTypesByStimParamsNoAmp.Add("Acclimatisation", New Dictionary)
                Call trialTypesByStimParamsNoAmp.Add("Acoustic", New Dictionary)
                Call trialTypesByStimParamsNoAmp.Add("Acoustic Only", New Dictionary)
                Call trialTypesByStimParamsNoAmp.Add("Electrical", New Dictionary)
                Call trialTypesByStimParamsNoAmp.Add("No Stim", New Dictionary)
                
                
                validTrialCount = 0
                Set thisAnimalWorksheet = Nothing
                Set thisAnimalSummarySheet = Nothing
                animalID = sourceWorksheet.Name
                
'                Call parseTrials(trialTypes, sourceWorksheet)
                Call parseTrials(sourceWorksheet)
                For iPass = 0 To 3
                    thisAnimalSummarySheetRow = 2
                    Select Case iPass
                        Case 0:
                            'clusterByStimParams = True
                            'clusterByDate = False
                            If outputByStimParamsFull Is Nothing Then
                                Set outputByStimParamsFull = Workbooks.Open(templateFilename)
                            End If
                            Set outputWorkbook = outputByStimParamsFull
                            Set trialTypes = trialTypesByStimParamsFull
                        Case 1:
                            'clusterByStimParams = False
                            'clusterByDate = True
                            If outputByDate Is Nothing Then
                                Set outputByDate = Workbooks.Open(templateFilename)
                            End If
                            Set outputWorkbook = outputByDate
                            Set trialTypes = trialTypesByDate
                        Case 2:
                            'clusterByStimParams = True
                            'clusterByDate = True
                            If outputByDateStimParamsFull Is Nothing Then
                                Set outputByDateStimParamsFull = Workbooks.Open(templateFilename)
                            End If
                            Set outputWorkbook = outputByDateStimParamsFull
                            Set trialTypes = trialTypesByDateStimParamsFull
                        Case 3:
                            'clusterByStimParams = True
                            'clusterByDate = True
                            If outputByStimParamsNoAmp Is Nothing Then
                                Set outputByStimParamsNoAmp = Workbooks.Open(templateFilename)
                            End If
                            Set outputWorkbook = outputByStimParamsNoAmp
                            Set trialTypes = trialTypesByStimParamsNoAmp
                    End Select

'                    Set outputWorkbook = Workbooks.Open(templateFilename)
                
                    Call outputWorkbook.Worksheets("Summary template").Copy(, outputWorkbook.Worksheets("Output template"))
                    Set thisAnimalSummarySheet = outputWorkbook.Worksheets("Summary template (2)")
                    thisAnimalSummarySheet.Name = animalID & " summary"
                    
                    thisAnimalSummarySheet.Cells(1, 1).Value = "Cluster"
                    thisAnimalSummarySheet.Cells(1, 2).Value = "HR Included trials"
                    thisAnimalSummarySheet.Cells(1, 3).Value = "HR Excluded trials"
                    thisAnimalSummarySheet.Cells(1, 4).Value = "% trials decrease HR"
                    thisAnimalSummarySheet.Cells(1, 5).Value = "Mean HR change"
                    thisAnimalSummarySheet.Cells(1, 6).Value = "HR StdDev"
                    thisAnimalSummarySheet.Cells(1, 7).Value = "T score"
                    thisAnimalSummarySheet.Cells(1, 8).Value = "p value"
                    thisAnimalSummarySheet.Cells(1, 10).Value = "-84s to -4s HR N"
                    thisAnimalSummarySheet.Cells(1, 11).Value = "-84s to -4s HR StdDev"
                    thisAnimalSummarySheet.Cells(1, 12).Value = "-84s to -4s HR StdDev StdDev"
                    thisAnimalSummarySheet.Cells(1, 13).Value = "-4s to 0s HR N"
                    thisAnimalSummarySheet.Cells(1, 14).Value = "-4s to 0s HR StdDev"
                    thisAnimalSummarySheet.Cells(1, 15).Value = "-4s to 0s HR StdDev StdDev"
                    thisAnimalSummarySheet.Cells(1, 16).Value = "5s to 9s HR N"
                    thisAnimalSummarySheet.Cells(1, 17).Value = "5s to 9s HR StdDev"
                    thisAnimalSummarySheet.Cells(1, 18).Value = "5s to 9s HR StdDev StdDev"
                    thisAnimalSummarySheet.Range("A1:R1").Font.Bold = True
                    
                    For iColHeadersForHRLine = 0 To 130
                        thisAnimalSummarySheet.Cells(1, 21 + iColHeadersForHRLine).Value = Round((iColHeadersForHRLine - 40) / 10, 2)
                    Next
                    
                    If trialTypes("Acclimatisation").Count > 0 Then
                        Call outputWorkbook.Worksheets("Output template").Copy(, outputWorkbook.Worksheets("Output template"))
                        Set thisAnimalWorksheet = outputWorkbook.Worksheets("Output template (2)")
                        thisAnimalWorksheet.Name = animalID & " Acclimatisation"
                        Call outputTrials(trialTypes, "Acclimatisation", thisAnimalWorksheet, thisAnimalSummarySheet, thisAnimalSummarySheetRow, sourceWorksheet)
                    End If
                    If trialTypes("Acoustic Only").Count > 0 Then
                        Call outputWorkbook.Worksheets("Output template").Copy(, outputWorkbook.Worksheets("Output template"))
                        Set thisAnimalWorksheet = outputWorkbook.Worksheets("Output template (2)")
                        thisAnimalWorksheet.Name = animalID & " Acoustic Only"
                        Call outputTrials(trialTypes, "Acoustic Only", thisAnimalWorksheet, thisAnimalSummarySheet, thisAnimalSummarySheetRow, sourceWorksheet)
                    End If
                    If trialTypes("Acoustic").Count > 0 Then
                        Call outputWorkbook.Worksheets("Output template").Copy(, outputWorkbook.Worksheets("Output template"))
                        Set thisAnimalWorksheet = outputWorkbook.Worksheets("Output template (2)")
                        thisAnimalWorksheet.Name = animalID & " Acoustic"
                        Call outputTrials(trialTypes, "Acoustic", thisAnimalWorksheet, thisAnimalSummarySheet, thisAnimalSummarySheetRow, sourceWorksheet)
                    End If
                    If trialTypes("Electrical").Count > 0 Then
                        Call outputWorkbook.Worksheets("Output template").Copy(, outputWorkbook.Worksheets("Output template"))
                        Set thisAnimalWorksheet = outputWorkbook.Worksheets("Output template (2)")
                        thisAnimalWorksheet.Name = animalID & " Electrical"
                        Call outputTrials(trialTypes, "Electrical", thisAnimalWorksheet, thisAnimalSummarySheet, thisAnimalSummarySheetRow, sourceWorksheet)
                    End If
                    If trialTypes("No Stim").Count > 0 Then
                        Call outputWorkbook.Worksheets("Output template").Copy(, outputWorkbook.Worksheets("Output template"))
                        Set thisAnimalWorksheet = outputWorkbook.Worksheets("Output template (2)")
                        thisAnimalWorksheet.Name = animalID & " No Stim"
                        Call outputTrials(trialTypes, "No Stim", thisAnimalWorksheet, thisAnimalSummarySheet, thisAnimalSummarySheetRow, sourceWorksheet)
                    End If
'                    Call outputWorkbook.SaveAs(outputFilename)
'                    Call outputWorkbook.Close
                Next
            End If
        Next
        If Not outputByStimParamsFull Is Nothing Then
            outputFilename = pathToData & "\aggregate by stim params.xlsx"
            Call outputByStimParamsFull.SaveAs(outputFilename)
            Call outputByStimParamsFull.Close
        End If
        If Not outputByDate Is Nothing Then
            outputFilename = pathToData & "\aggregate by date.xlsx"
            Call outputByDate.SaveAs(outputFilename)
            Call outputByDate.Close
        End If
        If Not outputByDateStimParamsFull Is Nothing Then
            outputFilename = pathToData & "\aggregate by stim params and date.xlsx"
            Call outputByDateStimParamsFull.SaveAs(outputFilename)
            Call outputByDateStimParamsFull.Close
        End If
        If Not outputByStimParamsNoAmp Is Nothing Then
            outputFilename = pathToData & "\aggregate by stim params without amp.xlsx"
            Call outputByStimParamsNoAmp.SaveAs(outputFilename)
            Call outputByStimParamsNoAmp.Close
        End If

'    Next
            
'    Set objFile = Nothing
'    Set Files = Nothing
                    
'    Set experimentFolders = Nothing
'    Set objExpFolder = Nothing
    
'    Set AnimalFolders = Nothing
'    Set objAnimalFolder = Nothing
               
'    Set rootFolder = Nothing
    Set objFS = Nothing
    
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic

End Sub

'Function parseTrials(outputDict As Dictionary, sourceWorksheet As Workbook, experimentDate As String, experimentTag As String, exclusionInfo As Variant)
Function parseTrials(sourceWorksheet As Worksheet)
    ReDim allTrials(0)
    Dim experimentDate As String
    Dim experimentTag As String
    Dim exclusionInfo As Variant
    
    Dim trialInfoByDate As String
    Dim trialInfoByStimParamsFull As String
    Dim trialInfoByDateStimParamsFull As String
    Dim trialInfoByStimParamsNoAmp As String

    Dim i As Integer
    i = 3
    
    Dim iParamOffset As Integer
    
    Dim trialInfo As String
    Dim param1 As String
    Dim param1composite As String
    Dim param2 As String
    Dim param2composite As String
    Dim acoAmps(3) As String 'param 1 lower, param 1 upper, param 2 lower, param 2 upper
    Dim elAmps(3) As String 'param 1 lower, param 1 upper, param 2 lower, param 2 upper
    
'    Dim trialInfoByDate As String
'    Dim trialInfoByStimParamsFull As String
'    Dim trialInfoByDateStimParamsFull As String
'    Dim trialInfoByStimParamsNoAmp As String
    
    Dim param1arr As Variant
    Dim param2arr As Variant
    
    Dim param1str As String
    Dim param2str As String
    
    Dim blnExcludeThis As Boolean
    
    Dim trialArr
    Dim paramArr
    
    Dim iCurrBlockNum As Integer
    
    Dim exclusionReason As String
    
    While sourceWorksheet.Range("A" & i).Value <> ""
        If sourceWorksheet.Range("M" & i).Value <> "" Then
            experimentDate = sourceWorksheet.Range("D" & i).Value
            experimentTag = sourceWorksheet.Range("E" & i).Value
            exclusionInfo = Array(sourceWorksheet.Range("F" & i).Value, sourceWorksheet.Range("G" & i).Value, sourceWorksheet.Range("H" & i).Value)
            
            param1composite = ""
            param2composite = ""
        
            param1 = sourceWorksheet.Range("N" & i).Value
            param2 = sourceWorksheet.Range("R" & i).Value
    
    '        If workbookToProcess.Worksheets("Output").Cells(i, 1).Value <> iCurrBlockNum Then
             iCurrBlockNum = sourceWorksheet.Range("I" & i).Value
             Call readAmpArrays(acoAmps, elAmps, param1, param2, sourceWorksheet, iCurrBlockNum, experimentTag)
    '        End If
           
            trialArr = Array()
            ReDim trialArr(15)
            'result array contains 11 elements
            '1:date
            '2:HR -84 to -4s from start
            '3:reason for -84 to -4s exclusion (if excluded)
            '4:HR at -4s to 0s
            '5:reason for -4s to 0s exclusion (if excluded)
            '6:HR at 5 to 9s
            '7:reason for 5 to 9s exclusion (if excluded)
            '8:reason for overall exclusion (from exclusion text file)
            '9:StdDev -84 to -4s from start
            '10:StdDev -4s to 0s from start
            '11:StdDev 5s to 9s from start
            '12: label
            '13: Trial number
            '14: Stim params
            '15: Row in worksheet
            '16: reason for -4 to 9s exclusion (if excluded)
    
    'Const HRDetCols = 16
    'Const HRDetHROffset = 5
    
    'Const HRDetHROffset = 5
    'Const HRDetInterpProp = 10
    'Const HRDetLongIntSamp = 12
    'Const HRDetLongIntBeat = 14
    'Const HRDetStdDev = 7
    
            trialArr(0) = experimentDate
            trialArr(11) = experimentTag
            trialArr(12) = sourceWorksheet.Range("J" & i).Value
            trialArr(14) = i
    
            exclusionReason = checkForHRExclusions(sourceWorksheet, i, HRDetOffset)
            If exclusionReason <> "" Then
                trialArr(1) = "=NA()"
                trialArr(2) = exclusionReason
            Else
                trialArr(1) = sourceWorksheet.Cells(i, HRDetOffset + HRDetHROffset).Value
                trialArr(8) = sourceWorksheet.Cells(i, HRDetOffset + HRDetStdDev).Value
            End If
    
            exclusionReason = checkForHRExclusions(sourceWorksheet, i, HRDetOffset + HRDetCols)
            If exclusionReason <> "" Then
                trialArr(3) = "=NA()"
                trialArr(4) = exclusionReason
            Else
                trialArr(3) = sourceWorksheet.Cells(i, HRDetOffset + HRDetCols + HRDetHROffset).Value
                trialArr(9) = sourceWorksheet.Cells(i, HRDetOffset + HRDetCols + HRDetStdDev).Value
            End If
            exclusionReason = checkForHRExclusions(sourceWorksheet, i, HRDetOffset + (2 * HRDetCols))
            If exclusionReason <> "" Then
                trialArr(5) = "=NA()"
                trialArr(6) = exclusionReason
            Else
                trialArr(5) = sourceWorksheet.Cells(i, HRDetOffset + (HRDetCols * 2) + HRDetHROffset).Value
                trialArr(10) = sourceWorksheet.Cells(i, HRDetOffset + (HRDetCols * 2) + HRDetStdDev).Value
            End If
            exclusionReason = checkForHRExclusions(sourceWorksheet, i, HRDetOffset + (3 * HRDetCols))
            trialArr(15) = exclusionReason
             
            If sourceWorksheet.Range("M" & i).Value = "Acoustic" Then
                If Not ((exclusionInfo(0) = "Acoustic" Or exclusionInfo(0) = "all") And exclusionInfo(1) = "") Then
                     If (exclusionInfo(0) = "Acoustic" Or exclusionInfo(0) = "all") And exclusionInfo(1) <> "" Then
                        If exclusionInfo(2) = "" Then
                            trialArr(7) = exclusionInfo(1)
                        ElseIf CDbl(exclusionInfo(2)) <= sourceWorksheet.Range("K" & i).Value Then  'check if there is a time cutoff which has been passed
                            trialArr(7) = exclusionInfo(1)
                        Else
                            trialArr(7) = ""
                        End If
                     Else
                        trialArr(7) = ""
                     End If

                     'acoustic trial - drop the last 2 letters to remove the Hz
                     If LCase(Right(param1, 2)) = "hz" Then
                         param1composite = Left(param1, Len(param1) - 2)
                     Else
                         param1composite = param1
                     End If
                     If LCase(Right(param2, 2)) = "hz" Then
                         param2composite = Left(param2, Len(param2) - 2)
                     Else
                         param2composite = param2
                     End If
                     
                     param1composite = param1composite & Replace(acoAmps(0), ".", "") & Replace(acoAmps(1), ".", "")
                     param2composite = param2composite & Replace(acoAmps(2), ".", "") & Replace(acoAmps(3), ".", "")
                     
                     param1str = CStr(param1) & " (" & acoAmps(0) & "dB to " & acoAmps(1) & "dB)"
                     param2str = CStr(param2) & " (" & acoAmps(2) & "dB to " & acoAmps(3) & "dB)"
                     
                     'organise the clustering info to generate the grouping value (trialInfo)
                    If CDbl(param1composite) > CDbl(param2composite) Then
                        trialInfoByDateStimParamsFull = experimentDate & ": " & param1str & ", " & param2str
                        trialInfoByStimParamsFull = param1str & ", " & param2str
                        trialInfoByStimParamsNoAmp = CStr(param1) & ", " & CStr(param2)
                        trialArr(13) = param1str & ", " & param2str
                    Else
                        trialInfoByDateStimParamsFull = experimentDate & ": " & param2str & ", " & param1str
                        trialInfoByStimParamsFull = param2str & ", " & param1str
                        trialInfoByStimParamsNoAmp = CStr(param2) & ", " & CStr(param1)
                        trialArr(13) = param2str & ", " & param1str
                    End If
                    trialInfoByDate = experimentDate
                     
'                    If UBound(allTrials) Then
                    ReDim Preserve allTrials(UBound(allTrials) + 1)
'                    Else
'                        ReDim allTrials(0)
'                    End If
                    allTrials(UBound(allTrials)) = trialArr
                    
                    If InStr(1, LCase(experimentTag), "acclimatisation", vbTextCompare) Then
                        Call addToDict(trialTypesByDate, trialInfoByDate, "Acclimatisation", UBound(allTrials))
                        Call addToDict(trialTypesByStimParamsFull, trialInfoByStimParamsFull, "Acclimatisation", UBound(allTrials))
                        Call addToDict(trialTypesByDateStimParamsFull, trialInfoByDateStimParamsFull, "Acclimatisation", UBound(allTrials))
                        Call addToDict(trialTypesByStimParamsNoAmp, trialInfoByStimParamsNoAmp, "Acclimatisation", UBound(allTrials))
                    Else
                        If InStr(1, LCase(experimentTag), "electrical", vbTextCompare) = 0 Then
                            Call addToDict(trialTypesByDate, trialInfoByDate, "Acoustic Only", UBound(allTrials))
                            Call addToDict(trialTypesByStimParamsFull, trialInfoByStimParamsFull, "Acoustic Only", UBound(allTrials))
                            Call addToDict(trialTypesByDateStimParamsFull, trialInfoByDateStimParamsFull, "Acoustic Only", UBound(allTrials))
                            Call addToDict(trialTypesByStimParamsNoAmp, trialInfoByStimParamsNoAmp, "Acoustic Only", UBound(allTrials))
                        End If
                        Call addToDict(trialTypesByDate, trialInfoByDate, "Acoustic", UBound(allTrials))
                        Call addToDict(trialTypesByStimParamsFull, trialInfoByStimParamsFull, "Acoustic", UBound(allTrials))
                        Call addToDict(trialTypesByDateStimParamsFull, trialInfoByDateStimParamsFull, "Acoustic", UBound(allTrials))
                        Call addToDict(trialTypesByStimParamsNoAmp, trialInfoByStimParamsNoAmp, "Acoustic", UBound(allTrials))
                    End If

                End If
            Else 'electrical trial
                If Not ((exclusionInfo(0) = "Electrical" Or exclusionInfo(0) = "all") And exclusionInfo(1) = "") Then
                     If (exclusionInfo(0) = "Electrical" Or exclusionInfo(0) = "all") And exclusionInfo(1) <> "" Then
                        If exclusionInfo(2) = "" Then
                            trialArr(7) = exclusionInfo(1)
                        ElseIf CDbl(exclusionInfo(2)) <= sourceWorksheet.Range("K" & i).Value Then  'check if there is a time cutoff which has been passed
                            trialArr(7) = exclusionInfo(1)
                        Else
                            trialArr(7) = ""
                        End If
                     Else
                        trialArr(7) = ""
                     End If
                     
                    param1arr = Split(param1, " ")
                    param2arr = Split(param2, " ")
                    
                    If Right(param1arr(0), 1) = "*" Then
                        param1composite = param1composite & Left(param1arr(0), Len(param1arr(0)) - 1)
                    Else
                        param1composite = param1composite & param1arr(0)
                    End If
                    If Right(param1arr(2), 1) = "*" Then
                        param1composite = param1composite & Left(param1arr(2), Len(param1arr(2)) - 1)
                    Else
                        param1composite = param1composite & param1arr(2)
                    End If
                    If Not param1composite = "00" Then 'if stim and ref chans are 0, then the freq is irrelevant
                        If Right(param1arr(4), 2) = "Hz" Then
                            param1composite = param1composite & Left(param1arr(4), Len(param1arr(4)) - 2)
                        Else
                            param1composite = param1composite & param1arr(4)
                        End If
                    End If
                    
                    If Right(param2arr(0), 1) = "*" Then
                        param2composite = param2composite & Left(param2arr(0), Len(param2arr(0)) - 1)
                    Else
                        param2composite = param2composite & param2arr(0)
                    End If
                    If Right(param2arr(2), 1) = "*" Then
                        param2composite = param2composite & Left(param2arr(2), Len(param2arr(2)) - 1)
                    Else
                        param2composite = param2composite & param2arr(2)
                    End If
                    If Not param2composite = "00" Then 'if stim and ref chans are 0, then the freq is irrelevant
                        If Right(param2arr(4), 2) = "Hz" Then
                            param2composite = param2composite & Left(param2arr(4), Len(param2arr(4)) - 2)
                        Else
                            param2composite = param2composite & param2arr(4)
                        End If
                    End If
                    
                    param1composite = param1composite & Replace(elAmps(0), ".", "") & Replace(elAmps(1), ".", "")
                    param2composite = param2composite & Replace(elAmps(2), ".", "") & Replace(elAmps(3), ".", "")
                    
                    If param1composite <> "0000" Then
                        param1str = CStr(param1) & " (" & elAmps(0) & "uA to " & elAmps(1) & "uA)"
                    Else
                        param1str = "No stimulation"
                    End If
                    
                    If param2composite <> "0000" Then
                        param2str = CStr(param2) & " (" & elAmps(2) & "uA to " & elAmps(3) & "uA)"
                    Else
                        param2str = "No stimulation"
                    End If
    
    
                    If CDbl(param1composite) > CDbl(param2composite) Then
                        trialInfoByDateStimParamsFull = experimentDate & ": " & param1str & ", " & param2str
                        trialInfoByStimParamsFull = param1str & ", " & param2str
                        trialInfoByStimParamsNoAmp = CStr(param1) & ", " & CStr(param2)
                        trialArr(13) = param1str & ", " & param2str
                    Else
                        trialInfoByDateStimParamsFull = experimentDate & ": " & param2str & ", " & param1str
                        trialInfoByStimParamsFull = param2str & ", " & param1str
                        trialInfoByStimParamsNoAmp = CStr(param2) & ", " & CStr(param1)
                        trialArr(13) = param2str & ", " & param1str
                    End If
                    trialInfoByDate = experimentDate


                    ReDim Preserve allTrials(UBound(allTrials) + 1)
                    allTrials(UBound(allTrials)) = trialArr
                    
                    If param1str = "No stimulation" Then
                        Call addToDict(trialTypesByDate, trialInfoByDate, "No Stim", UBound(allTrials))
                        Call addToDict(trialTypesByStimParamsFull, trialInfoByStimParamsFull, "No Stim", UBound(allTrials))
                        Call addToDict(trialTypesByDateStimParamsFull, trialInfoByDateStimParamsFull, "No Stim", UBound(allTrials))
                        Call addToDict(trialTypesByStimParamsNoAmp, trialInfoByStimParamsNoAmp, "No Stim", UBound(allTrials))
                    Else
                        Call addToDict(trialTypesByDate, trialInfoByDate, "Electrical", UBound(allTrials))
                        Call addToDict(trialTypesByStimParamsFull, trialInfoByStimParamsFull, "Electrical", UBound(allTrials))
                        Call addToDict(trialTypesByDateStimParamsFull, trialInfoByDateStimParamsFull, "Electrical", UBound(allTrials))
                        Call addToDict(trialTypesByStimParamsNoAmp, trialInfoByStimParamsNoAmp, "Electrical", UBound(allTrials))
                    End If
                End If
            End If
        End If
        i = i + 1
    Wend
End Function

Function checkForHRExclusions(sourceWorksheet As Worksheet, i As Integer, horizOffset As Integer) As String
            checkForHRExclusions = ""
            
            If sourceWorksheet.Cells(i, horizOffset + HRDetHROffset).Value = -1 Then
                checkForHRExclusions = "HR not detectable (" & sourceWorksheet.Cells(i, horizOffset + HRDetHROffset).Value & ")"
            ElseIf sourceWorksheet.Cells(i, horizOffset + HRDetInterpProp).Value > maxPercOfBeatsInt And maxPercOfBeatsInt <> -1 Then
                checkForHRExclusions = "Too large % interpolated (" & sourceWorksheet.Cells(i, horizOffset + HRDetInterpProp).Value & ">" & maxPercOfBeatsInt & ")"
            ElseIf sourceWorksheet.Cells(i, horizOffset + HRDetLongIntSamp).Value > maxSingleIntSamples And maxSingleIntSamples <> -1 Then
                checkForHRExclusions = "Longest interpolation too long (" & sourceWorksheet.Cells(i, horizOffset + HRDetLongIntSamp).Value & ">" & maxSingleIntSamples & ")"
            ElseIf sourceWorksheet.Cells(i, horizOffset + HRDetLongIntBeat).Value > maxSingleIntBeats And maxSingleIntBeats <> -1 Then
                checkForHRExclusions = "Longest interpolation too many beats (" & sourceWorksheet.Cells(i, horizOffset + HRDetLongIntBeat).Value & ">" & maxSingleIntBeats & ")"
            End If
End Function

Sub outputTrials(trialTypes As Dictionary, trialType As String, thisAnimalWorksheet As Worksheet, thisAnimalSummarySheet As Worksheet, ByRef thisAnimalSummarySheetRow, ByRef sourceWorksheet As Worksheet)
    Dim arrTrialTypes
    arrTrialTypes = trialTypes.Keys
    
    Dim formatCond As FormatCondition
    
    Dim dictParamSets As Dictionary
    
    Dim arrParamSets
    Dim arrTrials
    Dim arrTrial
    
    Dim TotalHRPlot() As Double
    Dim TotalHRSD() As Double
    Dim TotalHRIterator As Integer
    Dim TotalnInHrSoFar As Integer

    Dim HRPlot() As Double
    Dim HRSD() As Double
    Dim HRIterator As Integer
    Dim nInHrSoFar As Integer
    
    Dim iTrialTypeNum As Integer
    Dim iParamSetNum As Integer
    Dim iTrialNum As Integer
    
    Dim iExcelOffset As Long
    iExcelOffset = 1
    Dim iPrevExcelOffset As Long
    Dim iMaxExcelOffset As Long
    
    Dim iThisAnimalSummarySheetStartingRow As Integer
    iThisAnimalSummarySheetStartingRow = CInt(thisAnimalSummarySheetRow)
    
    Dim meanHRChange As Double
    Dim HRChangeVar As Double
    Dim nInMeanSoFar As Integer
    Dim nExcluded As Integer
    Dim diff As Double
    Dim tStat As Double
    
    Dim nInMeanStdDevSoFar(2) As Integer
    Dim StdDevVar(2) As Double
    Dim meanStdDev(2) As Double
    Dim iVarCycling As Integer
    Dim iSummaryCol As Integer

    Dim currPooledPretrialHRMean As Double
    Dim currPooledPretrialHRStdDev As Double

    Dim currPooledHRChMean As Double
    Dim currPooledHRChCum As Double
    Dim currPooledHRChN As Long
    Dim currPooledHRChNExcl As Long
    Dim currPooledHRChNDec As Long
           
    Dim currPooledVarMean As Variant
    Dim currPooledVarCum As Variant
    Dim currPooledVarN As Variant
        
    Dim HRIncTrials As Integer
    Dim HRDecTrials As Integer
    
    currPooledHRChN = 0
    currPooledHRChMean = 0
    currPooledHRChCum = 0
    currPooledHRChNExcl = 0
    currPooledHRChNDec = 0
    
    currPooledVarMean = Array(0#, 0#, 0#)
    currPooledVarCum = Array(0#, 0#, 0#)
    currPooledVarN = Array(0#, 0#, 0#)
    
    For iTrialTypeNum = 0 To UBound(arrTrialTypes)
        If arrTrialTypes(iTrialTypeNum) = trialType Then
            ReDim TotalHRPlot(130)
            ReDim TotalHRSD(130)

            thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = arrTrialTypes(iTrialTypeNum) & " Trials"
            'thisAnimalWorksheet.Cells(iExcelOffset, 1).Style = "Heading"
            thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Bold = True
            iExcelOffset = iExcelOffset + 1
            Set dictParamSets = trialTypes(arrTrialTypes(iTrialTypeNum))
            arrParamSets = dictParamSets.Keys
            For iParamSetNum = 0 To UBound(arrParamSets)
                thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 1) = arrTrialTypes(iTrialTypeNum) & ": " & arrParamSets(iParamSetNum)
                thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 1).Font.Bold = True
                
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = arrParamSets(iParamSetNum)
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Bold = True
                iExcelOffset = iExcelOffset + 1
                thisAnimalWorksheet.Range("A" & iExcelOffset, "H" & iExcelOffset).Font.Italic = True
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "Tag"
                thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = "Trial Number"
                thisAnimalWorksheet.Cells(iExcelOffset, 3).Value = "Stim params"
                thisAnimalWorksheet.Cells(iExcelOffset, 4).Value = "Date"
                thisAnimalWorksheet.Cells(iExcelOffset, 5).Value = "HR -84s to -4s"
                thisAnimalWorksheet.Cells(iExcelOffset, 6).Value = "HR -4s to 0s"
                thisAnimalWorksheet.Cells(iExcelOffset, 7).Value = "HR 5s to 9s"
                thisAnimalWorksheet.Cells(iExcelOffset, 8).Value = "StdDev -84s to -4s"
                thisAnimalWorksheet.Cells(iExcelOffset, 9).Value = "StdDev -4s to 0s"
                thisAnimalWorksheet.Cells(iExcelOffset, 10).Value = "StdDev 5s to 9s"
                thisAnimalWorksheet.Cells(iExcelOffset, 11).Value = "HR -84s to -4s exclusion reason"
                thisAnimalWorksheet.Cells(iExcelOffset, 12).Value = "HR -4s to 0s exclusion reason"
                thisAnimalWorksheet.Cells(iExcelOffset, 13).Value = "HR 5s to 9s exclusion reason"
                thisAnimalWorksheet.Cells(iExcelOffset, 14).Value = "Overall trial exclusion reason"
                thisAnimalWorksheet.Cells(iExcelOffset, 15).Value = "HR -4s to 9s exclusion reason"
                iExcelOffset = iExcelOffset + 1
                arrTrials = dictParamSets(arrParamSets(iParamSetNum))
                nInMeanSoFar = 0
                nExcluded = 0
                meanHRChange = 0
                HRChangeVar = 0
                HRIncTrials = 0
                HRDecTrials = 0
                nInMeanStdDevSoFar(0) = 0
                nInMeanStdDevSoFar(1) = 0
                nInMeanStdDevSoFar(2) = 0
                StdDevVar(0) = 0#
                StdDevVar(1) = 0#
                StdDevVar(2) = 0#
                meanStdDev(0) = 0#
                meanStdDev(1) = 0#
                meanStdDev(2) = 0#
                
                nInHrSoFar = 0
                ReDim HRPlot(130)
                ReDim HRSD(130)
'                For HRIterator = 0 To 130
'                    HRPlot(HRIterator) = 0
'                Next
                
                For iTrialNum = 0 To UBound(arrTrials)
                    arrTrial = allTrials(arrTrials(iTrialNum))
                                        
                                        
                    thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = arrTrial(11)
                    thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = arrTrial(12)
                    thisAnimalWorksheet.Cells(iExcelOffset, 3).Value = arrTrial(13)
                    thisAnimalWorksheet.Cells(iExcelOffset, 4).Value = arrTrial(0)
                    
                    thisAnimalWorksheet.Cells(iExcelOffset, 5).Value = arrTrial(1)
                    thisAnimalWorksheet.Cells(iExcelOffset, 6).Value = arrTrial(3)
                    thisAnimalWorksheet.Cells(iExcelOffset, 7).Value = arrTrial(5)
                    thisAnimalWorksheet.Cells(iExcelOffset, 8).Value = arrTrial(8)
                    thisAnimalWorksheet.Cells(iExcelOffset, 9).Value = arrTrial(9)
                    thisAnimalWorksheet.Cells(iExcelOffset, 10).Value = arrTrial(10)
                    
                    thisAnimalWorksheet.Cells(iExcelOffset, 11).Value = arrTrial(2)
                    thisAnimalWorksheet.Cells(iExcelOffset, 12).Value = arrTrial(4)
                    thisAnimalWorksheet.Cells(iExcelOffset, 13).Value = arrTrial(6)
                    thisAnimalWorksheet.Cells(iExcelOffset, 14).Value = arrTrial(15)

                    'If arrTrial(15) = "" And arrTrial(7) = "" Then 'check if the data should be excluded
                    If arrTrial(15) = "" And arrTrial(7) = "" And arrTrial(6) = "" And arrTrial(4) = "" Then 'check if the data should be excluded
                        'If Not arrParamSets(iParamSetNum) = "No stimulation, No stimulation" Then 'dont include if no stim - shouldn't be pooled with the rest
                            nInHrSoFar = nInHrSoFar + 1
                            For HRIterator = 0 To 130
                                HRPlot(HRIterator) = HRPlot(HRIterator) + ((sourceWorksheet.Cells(arrTrial(14), HRIterator + 102).Value - HRPlot(HRIterator)) / nInHrSoFar)
                                HRSD(HRIterator) = HRSD(HRIterator) + (sourceWorksheet.Cells(arrTrial(14), HRIterator + 102).Value ^ 2)
                            Next
                        'End If
                    End If


                    If arrTrial(4) <> "" Or arrTrial(6) <> "" Or arrTrial(7) <> "" Then
                        thisAnimalWorksheet.Cells(iExcelOffset, 14).Value = arrTrial(7)
                        thisAnimalWorksheet.Range("A" & iExcelOffset, "AZ" & iExcelOffset).Interior.Color = excludedTrialCell.Interior.Color
                        thisAnimalWorksheet.Range("A" & iExcelOffset, "AZ" & iExcelOffset).Interior.ColorIndex = excludedTrialCell.Interior.ColorIndex
                        thisAnimalWorksheet.Range("A" & iExcelOffset, "AZ" & iExcelOffset).Font.Color = excludedTrialCell.Font.Color
                        thisAnimalWorksheet.Range("A" & iExcelOffset, "AZ" & iExcelOffset).Font.ColorIndex = excludedTrialCell.Font.ColorIndex
                    End If
                    iExcelOffset = iExcelOffset + 1
                    
                    If arrTrial(4) = "" And arrTrial(6) = "" And arrTrial(7) = "" Then
                        'contribute to the mean
                        nInMeanSoFar = nInMeanSoFar + 1
                        diff = arrTrial(5) - arrTrial(3)
                        
                        
                        meanHRChange = meanHRChange + ((diff - meanHRChange) / CDbl(nInMeanSoFar))
                        
                        'If Not arrParamSets(iParamSetNum) = "No stimulation, No stimulation" Then
                            currPooledHRChN = currPooledHRChN + 1
                            currPooledPretrialHRMean = currPooledPretrialHRMean + ((arrTrial(3) - currPooledPretrialHRMean) / CDbl(currPooledHRChN))
                            currPooledPretrialHRStdDev = currPooledPretrialHRStdDev + (arrTrial(3) ^ 2)

                            currPooledHRChMean = currPooledHRChMean + ((diff - currPooledHRChMean) / CDbl(currPooledHRChN))
                            currPooledHRChCum = currPooledHRChCum + (diff ^ 2)
                        'Else
                        '    noStimPooledHRChN = noStimPooledHRChN + 1
                        '    noStimPooledHRChMean = noStimPooledHRChMean + ((diff - noStimPooledHRChMean) / CDbl(noStimPooledHRChN))
                        '    noStimPooledHRChCum = noStimPooledHRChCum + (diff ^ 2)
                        'End If
                        
                        'check if HR rose or fell
                        If diff < 0 Then
                            HRDecTrials = HRDecTrials + 1
                        Else
                            HRIncTrials = HRIncTrials + 1
                        End If
                    Else
                        nExcluded = nExcluded + 1
                    End If
                    
                    If arrTrial(2) = "" And arrTrial(7) = "" Then
                        nInMeanStdDevSoFar(0) = nInMeanStdDevSoFar(0) + 1
                        meanStdDev(0) = meanStdDev(0) + ((arrTrial(8) - meanStdDev(0)) / CDbl(nInMeanStdDevSoFar(0)))
                        
                        'If Not arrParamSets(iParamSetNum) = "No stimulation, No stimulation" Then
                            currPooledVarN(0) = currPooledVarN(0) + 1
                            currPooledVarMean(0) = currPooledVarMean(0) + CDbl(((arrTrial(8) - currPooledVarMean(0)) / CDbl(currPooledVarN(0))))
                            currPooledVarCum(0) = currPooledVarCum(0) + CDbl((arrTrial(8)) ^ 2)
                        'Else
                        '    noStimPooledVarN(0) = noStimPooledVarN(0) + 1
                        '    noStimPooledVarMean(0) = noStimPooledVarMean(0) + CDbl(((arrTrial(8) - noStimPooledVarMean(0)) / CDbl(noStimPooledVarN(0))))
                        '    noStimPooledVarCum(0) = noStimPooledVarCum(0) + CDbl((arrTrial(8)) ^ 2)
                        'End If
                    End If
                    
                    If arrTrial(4) = "" And arrTrial(7) = "" Then
                        nInMeanStdDevSoFar(1) = nInMeanStdDevSoFar(1) + 1
                        meanStdDev(1) = meanStdDev(1) + ((arrTrial(9) - meanStdDev(1)) / CDbl(nInMeanStdDevSoFar(1)))
                        
                        'If Not arrParamSets(iParamSetNum) = "No stimulation, No stimulation" Then
                            currPooledVarN(1) = currPooledVarN(1) + 1
                            currPooledVarMean(1) = currPooledVarMean(1) + CDbl(((arrTrial(9) - currPooledVarMean(1)) / CDbl(currPooledVarN(1))))
                            currPooledVarCum(1) = currPooledVarCum(1) + CDbl((arrTrial(9)) ^ 2)
                        'Else
                        '    noStimPooledVarN(1) = noStimPooledVarN(1) + 1
                        '    noStimPooledVarMean(1) = noStimPooledVarMean(1) + CDbl(((arrTrial(9) - noStimPooledVarMean(1)) / CDbl(noStimPooledVarN(1))))
                        '    noStimPooledVarCum(1) = noStimPooledVarCum(1) + CDbl((arrTrial(9)) ^ 2)
                        'End If
                    End If
                    
                    If arrTrial(6) = "" And arrTrial(7) = "" Then
                        nInMeanStdDevSoFar(2) = nInMeanStdDevSoFar(2) + 1
                        meanStdDev(2) = meanStdDev(2) + ((arrTrial(9) - meanStdDev(2)) / CDbl(nInMeanStdDevSoFar(2)))
                        
                        'If Not arrParamSets(iParamSetNum) = "No stimulation, No stimulation" Then
                            currPooledVarN(2) = currPooledVarN(2) + 1
                            currPooledVarMean(2) = currPooledVarMean(2) + CDbl(((arrTrial(10) - currPooledVarMean(2)) / CDbl(currPooledVarN(2))))
                            currPooledVarCum(2) = currPooledVarCum(2) + CDbl((arrTrial(10)) ^ 2)
                        'Else
                        '    noStimPooledVarN(2) = noStimPooledVarN(2) + 1
                        '    noStimPooledVarMean(2) = noStimPooledVarMean(2) + CDbl(((arrTrial(10) - noStimPooledVarMean(2)) / CDbl(noStimPooledVarN(2))))
                        '    noStimPooledVarCum(2) = noStimPooledVarCum(2) + CDbl((arrTrial(10)) ^ 2)
                        'End If
                    End If
                Next
                
                'calculate variance
                For iTrialNum = 0 To UBound(arrTrials)
                    arrTrial = allTrials(arrTrials(iTrialNum))
                    If arrTrial(4) = "" And arrTrial(6) = "" And arrTrial(7) = "" Then
                        diff = arrTrial(5) - arrTrial(3)
                        HRChangeVar = HRChangeVar + (meanHRChange - diff) ^ 2
'                        StdDevVar = StdDevVar + (meanStdDev - arrTrial(8)) ^ 2
                    End If
                    
                    If arrTrial(2) = "" And arrTrial(7) = "" Then
                        StdDevVar(0) = StdDevVar(0) + (meanStdDev(0) - arrTrial(8)) ^ 2
                    End If
                    
                    If arrTrial(4) = "" And arrTrial(7) = "" Then
                        StdDevVar(1) = StdDevVar(1) + (meanStdDev(1) - arrTrial(9)) ^ 2
                    End If
                    
                    If arrTrial(6) = "" And arrTrial(7) = "" Then
                        StdDevVar(2) = StdDevVar(2) + (meanStdDev(2) - arrTrial(10)) ^ 2
                    End If
                Next
                
                For iVarCycling = 0 To 2
                    If nInMeanStdDevSoFar(iVarCycling) > 1 Then
                        StdDevVar(iVarCycling) = StdDevVar(iVarCycling) / (nInMeanStdDevSoFar(iVarCycling) - 1)
                    End If
                Next
                
                If nInMeanSoFar > 1 Then
                    HRChangeVar = HRChangeVar / (nInMeanSoFar - 1)
                    'tStat = meanHRChange / ((HRChangeVar / nInMeanSoFar) ^ 0.5)
                    tStat = meanHRChange / ((HRChangeVar ^ 0.5) / (nInMeanSoFar ^ 0.5))
                End If
                
                iExcelOffset = iExcelOffset + 1
                iPrevExcelOffset = iExcelOffset
                
                For iVarCycling = 0 To 2
                    If nInMeanStdDevSoFar(iVarCycling) > 0 Then
                            Select Case iVarCycling
                                Case 0:
                                    iSummaryCol = 10
                                    thisAnimalWorksheet.Cells(iExcelOffset, 4).Value = "-84s to -4s"
                                    thisAnimalWorksheet.Cells(iExcelOffset, 4).Font.Bold = True
                                Case 1:
                                    iSummaryCol = 13
                                    thisAnimalWorksheet.Cells(iExcelOffset, 4).Value = "-4s to 0s"
                                    thisAnimalWorksheet.Cells(iExcelOffset, 4).Font.Bold = True
                                Case 2:
                                    iSummaryCol = 16
                                    thisAnimalWorksheet.Cells(iExcelOffset, 4).Value = "5s to 9s"
                                    thisAnimalWorksheet.Cells(iExcelOffset, 4).Font.Bold = True
                            End Select
                            iExcelOffset = iExcelOffset + 1
                            thisAnimalWorksheet.Cells(iExcelOffset, 4).Value = "N:"
                            thisAnimalWorksheet.Cells(iExcelOffset, 4).Font.Italic = True
                            thisAnimalWorksheet.Cells(iExcelOffset, 5).Value = nInMeanStdDevSoFar(iVarCycling)
                            thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, iSummaryCol) = nInMeanStdDevSoFar(iVarCycling)
                            iExcelOffset = iExcelOffset + 1
                            thisAnimalWorksheet.Cells(iExcelOffset, 4).Value = "Mean StdDev:"
                            thisAnimalWorksheet.Cells(iExcelOffset, 4).Font.Italic = True
                            thisAnimalWorksheet.Cells(iExcelOffset, 5).Value = meanStdDev(iVarCycling)
                            thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, iSummaryCol + 1) = meanStdDev(iVarCycling)
                        If nInMeanStdDevSoFar(iVarCycling) > 1 Then
                            iExcelOffset = iExcelOffset + 1
                            thisAnimalWorksheet.Cells(iExcelOffset, 4).Value = "StdDev of StdDev:"
                            thisAnimalWorksheet.Cells(iExcelOffset, 4).Font.Italic = True
                            thisAnimalWorksheet.Cells(iExcelOffset, 5).Value = StdDevVar(iVarCycling) ^ 0.5
                            thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, iSummaryCol + 2) = StdDevVar(iVarCycling) ^ 0.5
                        Else
                            iExcelOffset = iExcelOffset + 1
                            thisAnimalWorksheet.Cells(iExcelOffset, 4).Value = "StdDev of StdDev:"
                            thisAnimalWorksheet.Cells(iExcelOffset, 4).Font.Italic = True
                            thisAnimalWorksheet.Cells(iExcelOffset, 5).Value = "Could not be computed; N=1"
                            thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, iSummaryCol + 2) = "=NA()"
                        End If
                        iExcelOffset = iExcelOffset + 2
                    End If
                Next

                iMaxExcelOffset = iExcelOffset
                iExcelOffset = iPrevExcelOffset
                                
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
                    thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "HR decrease trials:"
                    thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Italic = True
                    thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = HRDecTrials
                    iExcelOffset = iExcelOffset + 1
                    thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "HR increase trials:"
                    thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Italic = True
                    thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = HRIncTrials
                    iExcelOffset = iExcelOffset + 1
                    thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "% decrease trials:"
                    thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Italic = True
                    thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = (HRDecTrials / nInMeanSoFar)
                    thisAnimalWorksheet.Cells(iExcelOffset, 2).Style = "Percent"
                    Call thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions.Delete
                    Call thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions.Add(xlCellValue, xlNotBetween, ".15", ".85")
                    thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(1).Font.Color = percOutside1585FC.Font.Color
                    thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(1).Font.ColorIndex = percOutside1585FC.Font.ColorIndex
                    thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(1).Interior.Color = percOutside1585FC.Interior.Color
                    thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(1).Interior.ColorIndex = percOutside1585FC.Interior.ColorIndex
                    Call thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions.Add(xlCellValue, xlNotBetween, ".25", ".75")
                    thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(2).Font.Color = percOutside2575FC.Font.Color
                    thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(2).Font.ColorIndex = percOutside2575FC.Font.ColorIndex
                    thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(2).Interior.Color = percOutside2575FC.Interior.Color
                    thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(2).Interior.ColorIndex = percOutside2575FC.Interior.ColorIndex
                    
                    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 4) = (HRDecTrials / nInMeanSoFar)
                    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 4).Style = "Percent"
                    Call thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 4).FormatConditions.Delete
                    Call thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 4).FormatConditions.Add(xlCellValue, xlNotBetween, ".15", ".85")
                    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 4).FormatConditions(1).Font.Color = percOutside1585FC.Font.Color
                    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 4).FormatConditions(1).Font.ColorIndex = percOutside1585FC.Font.ColorIndex
                    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 4).FormatConditions(1).Interior.Color = percOutside1585FC.Interior.Color
                    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 4).FormatConditions(1).Interior.ColorIndex = percOutside1585FC.Interior.ColorIndex
                    Call thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 4).FormatConditions.Add(xlCellValue, xlNotBetween, ".25", ".75")
                    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 4).FormatConditions(2).Font.Color = percOutside2575FC.Font.Color
                    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 4).FormatConditions(2).Font.ColorIndex = percOutside2575FC.Font.ColorIndex
                    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 4).FormatConditions(2).Interior.Color = percOutside2575FC.Interior.Color
                    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 4).FormatConditions(2).Interior.ColorIndex = percOutside2575FC.Interior.ColorIndex

                    
                    iExcelOffset = iExcelOffset + 2
                    thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "Mean change:"
                    thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Italic = True
                    thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = meanHRChange
                    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 5) = meanHRChange
                    iExcelOffset = iExcelOffset + 1
                    If nInMeanSoFar > 1 Then
                        thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "Variance:"
                        thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = HRChangeVar
                        iExcelOffset = iExcelOffset + 1
                        thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "Standard Deviation:"
                        thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = HRChangeVar ^ 0.5
                        thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 6).Value = HRChangeVar ^ 0.5
                        iExcelOffset = iExcelOffset + 1
                        thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "Std. Error of Mean:"
                        thisAnimalWorksheet.Cells(iExcelOffset, 2).Value = ((HRChangeVar / nInMeanSoFar) ^ 0.5)
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
                        
                        thisAnimalSummarySheet.Range("UE" & thisAnimalSummarySheetRow).Value = currPooledPretrialHRMean
                        thisAnimalSummarySheet.Range("UF" & thisAnimalSummarySheetRow).Value = ((currPooledPretrialHRStdDev - (((currPooledPretrialHRMean * CDbl(currPooledHRChN)) ^ 2) / CDbl(currPooledHRChN))) / CDbl(currPooledHRChN - 1)) ^ 0.5
                    Else
                        thisAnimalWorksheet.Cells(iExcelOffset, 1).Value = "Additional stats could not be calculated (N=1)"
                        thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 6) = "=NA()"
                        thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 7) = "=NA()"
                        thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 8) = "=NA()"
                    End If
                End If
                
                If nInHrSoFar > 0 Then
                    TotalnInHrSoFar = TotalnInHrSoFar + nInHrSoFar
                    For HRIterator = 0 To 130
                        TotalHRPlot(HRIterator) = TotalHRPlot(HRIterator) + ((HRPlot(HRIterator) - TotalHRPlot(HRIterator)) / (TotalnInHrSoFar / nInHrSoFar))
                        TotalHRSD(HRIterator) = TotalHRSD(HRIterator) + HRSD(HRIterator)

                        thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 21 + HRIterator) = HRPlot(HRIterator)

                        If nInHrSoFar > 1 Then
                            '1 SD
                            thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 153 + HRIterator).Value = (((HRSD(HRIterator) - (((HRPlot(HRIterator) * CDbl(nInHrSoFar)) ^ 2#) / CDbl(nInHrSoFar))) / CDbl(nInHrSoFar)) ^ 0.5)
                            '2 SE
                            'thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 286 + HRIterator).Value = 2 * ((((HRSD(HRIterator) - (((HRPlot(HRIterator) * CDbl(nInHrSoFar)) ^ 2#) / CDbl(nInHrSoFar))) / CDbl(nInHrSoFar)) ^ 0.5) / (CDbl(nInHrSoFar) ^ 0.5))
                            
                            thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 286 + HRIterator).Value = 2 * ((((HRSD(HRIterator) - (((HRPlot(HRIterator) * CDbl(nInHrSoFar)) ^ 2#) / CDbl(nInHrSoFar))) / CDbl(nInHrSoFar)) ^ 0.5) / (CDbl(nInHrSoFar) ^ 0.5))
                            
                            '1.96 SD (95% CI)
                            thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 419 + HRIterator).Value = (((HRSD(HRIterator) - (((HRPlot(HRIterator) * CDbl(nInHrSoFar)) ^ 2#) / CDbl(nInHrSoFar))) / CDbl(nInHrSoFar)) ^ 0.5) * 1.96
                        End If
                    Next
                End If
                
                If arrTrialTypes(iTrialTypeNum) = "Electrical" Then
                    'If Not arrParamSets(iParamSetNum) = "No stimulation, No stimulation" Then
                        currPooledHRChNExcl = currPooledHRChNExcl + nExcluded
                        currPooledHRChNDec = currPooledHRChNDec + HRDecTrials
                    'Else
                    '    noStimPooledHRChNExcl = noStimPooledHRChNExcl + nExcluded
                    '    noStimPooledHRChNDec = noStimPooledHRChNDec + HRDecTrials
                    'End If
                End If
                                
                If iMaxExcelOffset > iExcelOffset Then
                    iExcelOffset = iMaxExcelOffset + 2
                Else
                    iExcelOffset = iExcelOffset + 2
                End If
                
                Dim myChart As ChartObject
                Dim chartOffset As Integer
                Dim chartHeight As Integer
                If iThisAnimalSummarySheetStartingRow > 2 Then
                    'chartOffset = (iThisAnimalSummarySheetStartingRow) * 15.5 + (UBound(arrParamSets) + 2) * 15.5
                    chartOffset = thisAnimalSummarySheet.Range("A" & iThisAnimalSummarySheetStartingRow + UBound(arrParamSets) + 5 & ":A" & "A" & iThisAnimalSummarySheetStartingRow + UBound(arrParamSets) + 5 + 19).Top
                    chartHeight = thisAnimalSummarySheet.Range("A" & iThisAnimalSummarySheetStartingRow + UBound(arrParamSets) + 5 & ":A" & "A" & iThisAnimalSummarySheetStartingRow + UBound(arrParamSets) + 5 + 19).Height
                Else
                    chartOffset = thisAnimalSummarySheet.Range("A" & UBound(arrParamSets) + 7 & ":A" & UBound(arrParamSets) + 7 + 19).Top
                    chartHeight = thisAnimalSummarySheet.Range("A" & UBound(arrParamSets) + 7 & ":A" & UBound(arrParamSets) + 7 + 19).Height
                    'chartOffset = (UBound(arrParamSets) + 5) * 15.5
                End If

                Set myChart = thisAnimalSummarySheet.ChartObjects.Add(((thisAnimalSummarySheetRow - iThisAnimalSummarySheetStartingRow) * 500) + 1, chartOffset, 500, chartHeight)
                myChart.Chart.ChartType = xlLine
                myChart.Chart.SeriesCollection.NewSeries
                myChart.Chart.SeriesCollection(1).Name = thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 1).Value & " (N=" & nInHrSoFar & ")"
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
                '2 SE
                                myChart.Chart.SeriesCollection(1).ErrorBar Direction:=xlY, Include:=xlBoth, _
                    Type:=xlErrorBarTypeCustom, Amount:=thisAnimalSummarySheet.Range("$JZ$" & thisAnimalSummarySheetRow & ":$OZ$" & thisAnimalSummarySheetRow), MinusValues:=thisAnimalSummarySheet.Range("$JZ$" & thisAnimalSummarySheetRow & ":$OZ$" & thisAnimalSummarySheetRow)


                myChart.Chart.ChartTitle.Characters.Font.Size = 12
                'myChart.Chart.Axes(xlValue).MinimumScale = 0.85
                'myChart.Chart.Axes(xlValue).MaximumScale = 1.15
                myChart.Chart.Axes(xlValue).MinimumScale = -50
                myChart.Chart.Axes(xlValue).MaximumScale = 50
                
                thisAnimalSummarySheetRow = thisAnimalSummarySheetRow + 1
            Next
            iExcelOffset = iExcelOffset + 1
        End If
    Next
    
    thisAnimalSummarySheetRow = thisAnimalSummarySheetRow + 2
            
    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 1).Value = trialType
    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 1).Font.Bold = True
    
    For iVarCycling = 0 To 2
        Select Case iVarCycling
            Case 0:
                iSummaryCol = 10
            Case 1:
                iSummaryCol = 13
            Case 2:
                iSummaryCol = 16
        End Select

        thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, iSummaryCol) = currPooledVarN(iVarCycling)
        thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, iSummaryCol + 1) = currPooledVarMean(iVarCycling)
        If currPooledVarN(iVarCycling) > 1 Then
            thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, iSummaryCol + 2) = ((currPooledVarCum(iVarCycling) - ((currPooledVarMean(iVarCycling) * CDbl(currPooledVarN(iVarCycling)) ^ 2) / CDbl(currPooledVarN(iVarCycling)))) / CDbl(currPooledVarN(iVarCycling) - 1)) ^ 0.5
        End If
    Next

    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 2).Value = currPooledHRChN
    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 3).Value = currPooledHRChNExcl
    'thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 4).Value = currPooledHRChNDec
    
    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 4).Value = (currPooledHRChNDec / currPooledHRChN)
    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 4).Style = "Percent"
    Call thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 4).FormatConditions.Delete
    Call thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 4).FormatConditions.Add(xlCellValue, xlNotBetween, ".15", ".85")
    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 4).FormatConditions(1).Font.Color = percOutside1585FC.Font.Color
    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 4).FormatConditions(1).Font.ColorIndex = percOutside1585FC.Font.ColorIndex
    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 4).FormatConditions(1).Interior.Color = percOutside1585FC.Interior.Color
    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 4).FormatConditions(1).Interior.ColorIndex = percOutside1585FC.Interior.ColorIndex
    Call thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 4).FormatConditions.Add(xlCellValue, xlNotBetween, ".25", ".75")
    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 4).FormatConditions(2).Font.Color = percOutside2575FC.Font.Color
    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 4).FormatConditions(2).Font.ColorIndex = percOutside2575FC.Font.ColorIndex
    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 4).FormatConditions(2).Interior.Color = percOutside2575FC.Interior.Color
    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 4).FormatConditions(2).Interior.ColorIndex = percOutside2575FC.Interior.ColorIndex

    
    thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 5).Value = currPooledHRChMean
    If currPooledHRChN > 1 Then
        thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 6).Value = ((currPooledHRChCum - ((currPooledHRChMean * CDbl(currPooledHRChN) ^ 2) / CDbl(currPooledHRChN))) / CDbl(currPooledHRChN - 1)) ^ 0.5
        thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 7).Value = currPooledHRChMean / (thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 6).Value / (currPooledHRChN ^ 0.5))
        thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 8).Value = "=TDIST(ABS(" & thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 7).Address & ")," & CStr(currPooledHRChN - 1) & ",1)"
        thisAnimalSummarySheet.Range("UE" & thisAnimalSummarySheetRow).Value = currPooledPretrialHRMean
        thisAnimalSummarySheet.Range("UF" & thisAnimalSummarySheetRow).Value = ((currPooledPretrialHRStdDev - (((currPooledPretrialHRMean * CDbl(currPooledHRChN)) ^ 2) / CDbl(currPooledHRChN))) / CDbl(currPooledHRChN - 1)) ^ 0.5
    End If
            
    If TotalnInHrSoFar > 0 Then
        For HRIterator = 0 To 130
            thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 21 + HRIterator) = TotalHRPlot(HRIterator)

            If TotalnInHrSoFar > 1 Then
                '1 SD
                thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 153 + HRIterator).Value = (((TotalHRSD(HRIterator) - (((TotalHRPlot(HRIterator) * CDbl(TotalnInHrSoFar)) ^ 2#) / CDbl(TotalnInHrSoFar))) / CDbl(TotalnInHrSoFar)) ^ 0.5)
                '2 SE
                thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 286 + HRIterator).Value = 2 * ((((TotalHRSD(HRIterator) - (((TotalHRPlot(HRIterator) * CDbl(TotalnInHrSoFar)) ^ 2#) / CDbl(TotalnInHrSoFar))) / CDbl(TotalnInHrSoFar)) ^ 0.5) / (CDbl(TotalnInHrSoFar) ^ 0.5))
                '1.96 SD (95% CI)
                thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 419 + HRIterator).Value = (((TotalHRSD(HRIterator) - (((TotalHRPlot(HRIterator) * CDbl(TotalnInHrSoFar)) ^ 2#) / CDbl(TotalnInHrSoFar))) / CDbl(TotalnInHrSoFar)) ^ 0.5) * 1.96
            End If
        Next
    End If


    If iThisAnimalSummarySheetStartingRow > 2 Then
        'chartOffset = (iThisAnimalSummarySheetStartingRow) * 15.5 + (UBound(arrParamSets) + 2) * 15.5
        chartOffset = thisAnimalSummarySheet.Range("A" & (iThisAnimalSummarySheetStartingRow + UBound(arrParamSets) + 5) & ":A" & "A" & (iThisAnimalSummarySheetStartingRow + UBound(arrParamSets) + 5 + 19)).Top
        chartHeight = thisAnimalSummarySheet.Range("A" & (iThisAnimalSummarySheetStartingRow + UBound(arrParamSets) + 5) & ":A" & "A" & (iThisAnimalSummarySheetStartingRow + UBound(arrParamSets) + 5 + 19)).Height
    Else
        chartOffset = thisAnimalSummarySheet.Range("A" & (UBound(arrParamSets) + 7) & ":A" & (UBound(arrParamSets) + 7 + 19)).Top
        chartHeight = thisAnimalSummarySheet.Range("A" & (UBound(arrParamSets) + 7) & ":A" & (UBound(arrParamSets) + 7 + 19)).Height
        'chartOffset = (UBound(arrParamSets) + 5) * 15.5
    End If

    Set myChart = thisAnimalSummarySheet.ChartObjects.Add(((thisAnimalSummarySheetRow - iThisAnimalSummarySheetStartingRow - 2) * 500) + 1, chartOffset, 500, chartHeight)
    myChart.Chart.ChartType = xlLine
    myChart.Chart.SeriesCollection.NewSeries
    myChart.Chart.SeriesCollection(1).Name = thisAnimalSummarySheet.Cells(thisAnimalSummarySheetRow, 1).Value & " (N=" & TotalnInHrSoFar & ")"
    myChart.Chart.SeriesCollection(1).Format.Line.Weight = 1#
    myChart.Chart.SeriesCollection(1).XValues = thisAnimalSummarySheet.Range("=$U$1:$EU$1")
    myChart.Chart.Legend.Delete
    myChart.Chart.SeriesCollection(1).Values = thisAnimalSummarySheet.Range("$U$" & thisAnimalSummarySheetRow & ":$EU$" & thisAnimalSummarySheetRow)
    myChart.Chart.SeriesCollection(1).HasErrorBars = True
    '1.96 Standard deviation
'    myChart.Chart.SeriesCollection(1).ErrorBar Direction:=xlY, Include:=xlBoth, _
 '       Type:=xlErrorBarTypeCustom, Amount:=thisAnimalSummarySheet.Range("$PC$" & thisAnimalSummarySheetRow & ":$UC$" & thisAnimalSummarySheetRow), MinusValues:=thisAnimalSummarySheet.Range("$PC$" & thisAnimalSummarySheetRow & ":$UC$" & thisAnimalSummarySheetRow)

    '1 Standard deviation
'    myChart.Chart.SeriesCollection(1).ErrorBar Direction:=xlY, Include:=xlBoth, _
'        Type:=xlErrorBarTypeCustom, Amount:=thisAnimalSummarySheet.Range("$EW$" & thisAnimalSummarySheetRow & ":$JW$" & thisAnimalSummarySheetRow), MinusValues:=thisAnimalSummarySheet.Range("$EW$" & thisAnimalSummarySheetRow & ":$JW$" & thisAnimalSummarySheetRow)
    '2 SE
   myChart.Chart.SeriesCollection(1).ErrorBar Direction:=xlY, Include:=xlBoth, _
       Type:=xlErrorBarTypeCustom, Amount:=thisAnimalSummarySheet.Range("$JZ$" & thisAnimalSummarySheetRow & ":$OZ$" & thisAnimalSummarySheetRow), MinusValues:=thisAnimalSummarySheet.Range("$JZ$" & thisAnimalSummarySheetRow & ":$OZ$" & thisAnimalSummarySheetRow)

    myChart.Chart.ChartTitle.Characters.Font.Size = 12
    'myChart.Chart.Axes(xlValue).MinimumScale = 0.85
    'myChart.Chart.Axes(xlValue).MaximumScale = 1.15
    myChart.Chart.Axes(xlValue).MinimumScale = -50
    myChart.Chart.Axes(xlValue).MaximumScale = 50

    thisAnimalSummarySheetRow = thisAnimalSummarySheetRow + 23

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
Function readAmpArrays(ByRef acoAmps, ByRef elAmps, param1 As String, param2 As String, sourceWorksheet As Worksheet, iCurrBlockNum As Integer, experimentTag As String) As Boolean
    readAmpArrays = True
'        param1acoLoweramp = workbookToProcess.Worksheets("Output").Cells(i, 7)
'        param1acoUpperamp = workbookToProcess.Worksheets("Output").Cells(i, 8)
'        param2acoLoweramp = workbookToProcess.Worksheets("Output").Cells(i, 11)
'        param2acoUpperamp = workbookToProcess.Worksheets("Output").Cells(i, 12)
    Dim iRow As Integer
    Dim iIsFirstAcoEntry As Boolean
    Dim iIsFirstElEntry As Boolean
    iIsFirstAcoEntry = True
    iIsFirstElEntry = True
    iRow = 2
    
    Dim param1LowerAmp As Double
    Dim param1UpperAmp As Double
    Dim param2LowerAmp As Double
    Dim param2UpperAmp As Double
    While sourceWorksheet.Range("E" & iRow).Value <> experimentTag And sourceWorksheet.Range("E" & iRow).Value <> ""
        iRow = iRow + 1
    Wend
    While sourceWorksheet.Range("I" & iRow).Value <> iCurrBlockNum And sourceWorksheet.Range("I" & iRow).Value <> ""
        iRow = iRow + 1
    Wend
    If sourceWorksheet.Range("I" & iRow).Value = "" Then 'check we found the row for the block
        readAmpArrays = False
        Exit Function
    End If
    While sourceWorksheet.Range("E" & iRow).Value = experimentTag And sourceWorksheet.Range("I" & iRow).Value = iCurrBlockNum And sourceWorksheet.Range("I" & iRow).Value <> ""
        If sourceWorksheet.Range("O" & iRow).Value = "" Or sourceWorksheet.Range("P" & iRow).Value = "" Or sourceWorksheet.Range("S" & iRow).Value = "" Or sourceWorksheet.Range("T" & iRow).Value = "" Then
            param1LowerAmp = 0
            param1UpperAmp = 0
            param2LowerAmp = 0
            param1UpperAmp = 0
        ElseIf sourceWorksheet.Range("N" & iRow).Value = param1 Then
            param1LowerAmp = CDbl(trimAmpTrailingChars(sourceWorksheet.Range("O" & iRow).Value))
            param1UpperAmp = CDbl(trimAmpTrailingChars(sourceWorksheet.Range("P" & iRow).Value))
            param2LowerAmp = CDbl(trimAmpTrailingChars(sourceWorksheet.Range("S" & iRow).Value))
            param2UpperAmp = CDbl(trimAmpTrailingChars(sourceWorksheet.Range("T" & iRow).Value))
        Else
            param1LowerAmp = CDbl(trimAmpTrailingChars(sourceWorksheet.Range("S" & iRow).Value))
            param1UpperAmp = CDbl(trimAmpTrailingChars(sourceWorksheet.Range("T" & iRow).Value))
            param2LowerAmp = CDbl(trimAmpTrailingChars(sourceWorksheet.Range("O" & iRow).Value))
            param2UpperAmp = CDbl(trimAmpTrailingChars(sourceWorksheet.Range("P" & iRow).Value))
        End If

        If sourceWorksheet.Range("M" & iRow).Value = "Acoustic" Then
            If iIsFirstAcoEntry Then
                acoAmps(0) = param1LowerAmp
                acoAmps(1) = param1UpperAmp
                acoAmps(2) = param2LowerAmp
                acoAmps(3) = param2UpperAmp
                iIsFirstAcoEntry = False
            Else
                If acoAmps(0) > param1LowerAmp Then
                    acoAmps(0) = param1LowerAmp
                End If
                If acoAmps(1) < param1UpperAmp Then
                    acoAmps(1) = param1UpperAmp
                End If
                If acoAmps(2) > param2LowerAmp Then
                    acoAmps(2) = param2LowerAmp
                End If
                If acoAmps(3) < param2UpperAmp Then
                    acoAmps(3) = param2UpperAmp
                End If
            End If
        Else
           If iIsFirstElEntry Then
                elAmps(0) = param1LowerAmp
                elAmps(1) = param1UpperAmp
                elAmps(2) = param2LowerAmp
                elAmps(3) = param2UpperAmp
                iIsFirstElEntry = False
            Else
                If elAmps(0) > param1LowerAmp Then
                    elAmps(0) = param1LowerAmp
                End If
                If elAmps(1) < param1UpperAmp Then
                    elAmps(1) = param1UpperAmp
                End If
                If elAmps(2) > param2LowerAmp Then
                    elAmps(2) = param2LowerAmp
                End If
                If elAmps(3) < param2UpperAmp Then
                    elAmps(3) = param2UpperAmp
                End If
            End If
        End If
        iRow = iRow + 1
    Wend
End Function
Function trimAmpTrailingChars(strToTrim As String) As String
    If LCase(Right(strToTrim, 2)) = "db" Or LCase(Right(strToTrim, 2)) = "ua" Then
        trimAmpTrailingChars = Left(strToTrim, Len(strToTrim) - 2)
    Else
        trimAmpTrailingChars = strToTrim
    End If
End Function
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
    iLenOfPrefix = Len("exclude from results aggregration - ")

    For Each objFile In Files
        If LCase(objFile.Name) = "exclude from results aggregration.txt" Then
            exclusionInfo(0) = "folder"
            Exit For
        ElseIf LCase(Left(objFile.Name, iLenOfPrefix)) = "exclude from results aggregration - " Then
            'exclude from results aggregration - all.txt
            tmpStr1 = Right(LCase(objFile.Name), Len(objFile.Name) - iLenOfPrefix)
            tmpStr2 = Left(tmpStr1, Len(tmpStr1) - 4)
            
            If LCase(Left(tmpStr2, Len("partial"))) = "partial" Then
               exclusionInfo(2) = readPartialFromFile(objFile)
               tmpStr2 = Right(tmpStr2, Len(tmpStr2) - Len("partial") - 1)
            End If
            
            Select Case tmpStr2
                Case "all":
                    exclusionInfo(0) = "all"
                Case "all with message":
                    exclusionInfo(0) = "all"
                    exclusionInfo(1) = readCommentFromFile(objFile)
                Case "acoustic":
                    exclusionInfo(0) = "Acoustic"
                Case "acoustic with message":
                    exclusionInfo(0) = "Acoustic"
                    exclusionInfo(1) = readCommentFromFile(objFile)
                Case "electical":
                    exclusionInfo(0) = "Electrical"
                Case "electrical with message":
                    exclusionInfo(0) = "Electrical"
                    exclusionInfo(1) = readCommentFromFile(objFile)
                'exclude from results aggregration - partial electrical with message.txt
            End Select
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

Function addToDict(ByRef objDict As Dictionary, trialInfo As String, trialType As String, iRow As Integer)
    Dim paramArr As Variant
    Dim iParamOffset As Integer
    If Not objDict(trialType).Exists(trialInfo) Then
        Call objDict(trialType).Add(trialInfo, Array())
    End If
                     
    paramArr = objDict(trialType)(trialInfo)
                     
    ReDim Preserve paramArr(UBound(paramArr) + 1)
    iParamOffset = UBound(paramArr)
    paramArr(iParamOffset) = iRow
                     
    objDict(trialType)(trialInfo) = paramArr
End Function
