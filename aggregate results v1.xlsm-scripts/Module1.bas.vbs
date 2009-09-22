Attribute VB_Name = "Module1"
Option Explicit

Global exIntCountGT As Integer
Global exIntBeatsGT As Integer
Global exLongestIntDurGT As Integer
Global exLongestIntBeatsGT As Integer

Global pLess05FC As FormatCondition
Global pLess10FC As FormatCondition
Global percOutside1585FC As FormatCondition
Global percOutside2575FC As FormatCondition


Sub aggregrate_results()
    Dim oneAnimalOneSheet As Boolean

    Dim trialTypes As Dictionary
    Dim validTrialCount As Integer
    
    Dim animalID As String
    Dim experimentDate As String
    Dim experimentTag As String

    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Dim objFS As FileSystemObject
    Set objFS = CreateObject("Scripting.FileSystemObject")
    
    'get the root folder under which all data is housed
    Dim rootFolder As Folder
    Set rootFolder = objFS.GetFolder(objFS.GetFolder(objFS.GetParentFolderName(ActiveWorkbook.FullName)))
        
    Dim AnimalFolders As Folders
    Dim objAnimalFolder As Folder
    
    Dim experimentFolders As Folders
    Dim objExpFolder As Folder
    
    Dim Files As Files
    Dim objFile As File
    
    Dim strExcelFilename As String
    Dim strExcelPathname As String
    
    Dim blnCurrFolderIsTrial As Boolean
    blnCurrFolderIsTrial = False
    
    Dim thisWorkbook As Workbook
    Set thisWorkbook = ActiveWorkbook
    Dim thisAnimalWorksheet As Worksheet
    
    'Call setUpStyles(thisWorkbook)
    
    oneAnimalOneSheet = thisWorkbook.Worksheets("Controller").Cells(9, 2).Value

    exIntCountGT = CInt(thisWorkbook.Worksheets("Controller").Cells(3, 2).Value)
    exIntBeatsGT = CInt(thisWorkbook.Worksheets("Controller").Cells(4, 2).Value)
    exLongestIntDurGT = CInt(thisWorkbook.Worksheets("Controller").Cells(5, 2).Value)
    exLongestIntBeatsGT = CInt(thisWorkbook.Worksheets("Controller").Cells(6, 2).Value)
    
    Set pLess05FC = thisWorkbook.Worksheets("Controller").Range("B11").FormatConditions(1)
    Set pLess10FC = thisWorkbook.Worksheets("Controller").Range("B12").FormatConditions(1)
    
    Set percOutside1585FC = thisWorkbook.Worksheets("Controller").Range("B14").FormatConditions(1)
    Set percOutside2575FC = thisWorkbook.Worksheets("Controller").Range("B15").FormatConditions(1)
    
    Dim workbookToProcess As Workbook
    
    Call deleteOldWorksheets(thisWorkbook)
    
    Set AnimalFolders = rootFolder.Subfolders
    For Each objAnimalFolder In AnimalFolders 'cycle through the folder for each animal
        If Not checkForExclusion(objAnimalFolder) Then
            Set trialTypes = New Dictionary
            Call trialTypes.Add("Acoustic", New Dictionary)
            Call trialTypes.Add("Electrical", New Dictionary)
            validTrialCount = 0
            Set thisAnimalWorksheet = Nothing
            animalID = objAnimalFolder.Name
                        
            Set experimentFolders = objAnimalFolder.Subfolders
            For Each objExpFolder In experimentFolders 'go through the experiments within an animal folder
                If Not checkForExclusion(objExpFolder) Then
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
                            validTrialCount = validTrialCount + 1
                            'open the workbook to read data from
                            Set workbookToProcess = Workbooks.Open(strExcelPathname)
                            Call parseTrials(trialTypes, workbookToProcess, experimentDate, experimentTag)
                            'workbookToProcess.Activate
                            'workbookToProcess.Worksheets("Variables (do not edit)").Range("B2").Value = tankFilename
                            'workbookToProcess.Worksheets("Variables (do not edit)").Range("B3").Value = blockName
                            'Application.Run ("'" & strExcelFilename & "'!importTrialsFromLabchart")
                            'Call workbookToProcess.Save
                            Call workbookToProcess.Close
                        End If
                    End If
                End If
            Next
            If validTrialCount > 0 Then
                If oneAnimalOneSheet Then
                    Call thisWorkbook.Worksheets("Output template").Copy(, thisWorkbook.Worksheets("Output template"))
                    Set thisAnimalWorksheet = thisWorkbook.Worksheets("Output template (2)")
                    thisAnimalWorksheet.Name = animalID
                    Call outputTrials(trialTypes, "", thisAnimalWorksheet)
                Else
                    If trialTypes("Acoustic").Count > 0 Then
                        Call thisWorkbook.Worksheets("Output template").Copy(, thisWorkbook.Worksheets("Output template"))
                        Set thisAnimalWorksheet = thisWorkbook.Worksheets("Output template (2)")
                        thisAnimalWorksheet.Name = animalID & " Acoustic"
                        Call outputTrials(trialTypes, "Acoustic", thisAnimalWorksheet)
                    End If
                    If trialTypes("Electrical").Count > 0 Then
                        Call thisWorkbook.Worksheets("Output template").Copy(, thisWorkbook.Worksheets("Output template"))
                        Set thisAnimalWorksheet = thisWorkbook.Worksheets("Output template (2)")
                        thisAnimalWorksheet.Name = animalID & " Electrical"
                        Call outputTrials(trialTypes, "Electrical", thisAnimalWorksheet)
                    End If
                End If
            End If
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
    
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Function checkForExclusion(objFolder As Folder) As Boolean
    checkForExclusion = False
    Dim Files As Files
    Dim objFile As File

    Set Files = objFolder.Files

    For Each objFile In Files
        If LCase(objFile.Name) = "exclude from results aggregration.txt" Then
            checkForExclusion = True
            Exit For
        End If
    Next

End Function

Function parseTrials(outputDict As Dictionary, workbookToProcess As Workbook, experimentDate As String, experimentTag As String)
    Dim i As Integer
    i = 2
    
    Dim iParamOffset As Integer
    
    Dim trialInfo As String
    Dim param1 As String
    Dim param1composite As String
    Dim param2 As String
    Dim param2composite As String
    Dim acoAmps(3) As String 'param 1 lower, param 1 upper, param 2 lower, param 2 upper
    Dim elAmps(3) As String 'param 1 lower, param 1 upper, param 2 lower, param 2 upper
    
    Dim param1arr As Variant
    Dim param2arr As Variant
    
    Dim trialArr
    Dim paramArr
    
    Dim iCurrBlockNum As Integer
    
    Dim exclusionReason As String
    
    While workbookToProcess.Worksheets("Settings").Cells(i, 6) <> ""
        param1composite = ""
        param2composite = ""
    
        param1 = workbookToProcess.Worksheets("Settings").Cells(i, 6)
        param2 = workbookToProcess.Worksheets("Settings").Cells(i, 10)

        If workbookToProcess.Worksheets("Settings").Cells(i, 1) <> iCurrBlockNum Then
            iCurrBlockNum = workbookToProcess.Worksheets("Settings").Cells(i, 1)
            Call readAmpArrays(acoAmps, elAmps, param1, param2, workbookToProcess, iCurrBlockNum)
        End If
       
        trialArr = Array()
        ReDim trialArr(7) 'result array contains seven elements - date, HR 10-30s from start, reason for 10-30s exclusion (if excluded), HR at -4s, reason for -4s exclusion (if excluded), HR at 5-9s, reason for 5-9s exclusion (if excluded)
        trialArr(0) = experimentTag & " Trial" & workbookToProcess.Worksheets("Settings").Cells(i, 2)
        If i = 2 Then
            trialArr(1) = "=NA()"
            trialArr(2) = "First trial"
        Else
            exclusionReason = checkForHRExclusions(workbookToProcess, i, 1)
            If exclusionReason <> "" Then
                trialArr(1) = "=NA()"
                trialArr(2) = exclusionReason
            Else
                trialArr(1) = workbookToProcess.Worksheets("HR detection").Cells(i + 1, 1)
            End If
        End If
        exclusionReason = checkForHRExclusions(workbookToProcess, i, 7)
        If exclusionReason <> "" Then
            trialArr(3) = "=NA()"
            trialArr(4) = exclusionReason
        Else
            trialArr(3) = workbookToProcess.Worksheets("HR detection").Cells(i + 1, 7)
        End If
        exclusionReason = checkForHRExclusions(workbookToProcess, i, 13)
        If exclusionReason <> "" Then
            trialArr(5) = "=NA()"
            trialArr(6) = exclusionReason
        Else
            trialArr(5) = workbookToProcess.Worksheets("HR detection").Cells(i + 1, 13)
        End If
        
        If workbookToProcess.Worksheets("Settings").Cells(i, 5) = "Acoustic" Then 'acoustic trial - drop the last 2 letters to remove the Hz
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
            
            If CDbl(param1composite) > CDbl(param2composite) Then
                trialInfo = CStr(param1) & " (" & acoAmps(0) & "dB to " & acoAmps(1) & "dB), " & CStr(param2) & " (" & acoAmps(2) & "dB to " & acoAmps(3) & "dB)"
            Else
                trialInfo = CStr(param2) & " (" & acoAmps(2) & "dB to " & acoAmps(3) & "dB), " & CStr(param1) & " (" & acoAmps(0) & "dB to " & acoAmps(1) & "dB)"
            End If
            
            If Not outputDict("Acoustic").Exists(trialInfo) Then
                Call outputDict("Acoustic").Add(trialInfo, Array())
            End If
            paramArr = outputDict("Acoustic")(trialInfo)
            
            ReDim Preserve paramArr(UBound(paramArr) + 1)
            iParamOffset = UBound(paramArr)
            paramArr(iParamOffset) = trialArr
            
            outputDict("Acoustic")(trialInfo) = paramArr
        Else 'electrical trial
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
            If Right(param1arr(4), 2) = "Hz" Then
                param1composite = param1composite & Left(param1arr(4), Len(param1arr(4)) - 2)
            Else
                param1composite = param1composite & param1arr(4)
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
            If Right(param2arr(4), 2) = "Hz" Then
                param2composite = param2composite & Left(param2arr(4), Len(param2arr(4)) - 2)
            Else
                param2composite = param2composite & param2arr(4)
            End If
            
            param1composite = param1composite & Replace(elAmps(0), ".", "") & Replace(elAmps(1), ".", "")
            param2composite = param2composite & Replace(elAmps(2), ".", "") & Replace(elAmps(3), ".", "")
            
            If CDbl(param1composite) > CDbl(param2composite) Then
                trialInfo = CStr(param1) & " (" & elAmps(0) & "uA to " & elAmps(1) & "uA), " & CStr(param2) & " (" & elAmps(2) & "uA to " & elAmps(3) & "uA)"
            Else
                trialInfo = CStr(param2) & " (" & elAmps(2) & "uA to " & elAmps(3) & "uA), " & CStr(param1) & " (" & elAmps(0) & "uA to " & elAmps(1) & "uA)"
            End If
            
            If Not outputDict("Electrical").Exists(trialInfo) Then
                Call outputDict("Electrical").Add(trialInfo, Array())
            End If
            paramArr = outputDict("Electrical")(trialInfo)
            
            ReDim Preserve paramArr(UBound(paramArr) + 1)
            iParamOffset = UBound(paramArr)
            paramArr(iParamOffset) = trialArr
            
            outputDict("Electrical")(trialInfo) = paramArr
        End If
        i = i + 1
    Wend
End Function

Function checkForHRExclusions(workbookToProcess As Workbook, i As Integer, horizOffset As Integer) As String
            checkForHRExclusions = ""
            If workbookToProcess.Worksheets("HR detection").Cells(i + 1, horizOffset) = -1 Then
                checkForHRExclusions = "HR not detectable (" & workbookToProcess.Worksheets("HR detection").Cells(i + 1, horizOffset) & ")"
            ElseIf workbookToProcess.Worksheets("HR detection").Cells(i + 1, horizOffset + 1) > exIntCountGT And exIntCountGT <> -1 Then
                checkForHRExclusions = "Too many interpolations (" & workbookToProcess.Worksheets("HR detection").Cells(i + 1, horizOffset + 1) & ">" & exIntCountGT & ")"
            ElseIf workbookToProcess.Worksheets("HR detection").Cells(i + 1, horizOffset + 2) > exIntBeatsGT And exIntBeatsGT <> -1 Then
                checkForHRExclusions = "Too many interpolated beats (" & workbookToProcess.Worksheets("HR detection").Cells(i + 1, horizOffset + 2) & ">" & exIntBeatsGT & ")"
            ElseIf workbookToProcess.Worksheets("HR detection").Cells(i + 1, horizOffset + 3) > exLongestIntDurGT And exLongestIntDurGT <> -1 Then
                checkForHRExclusions = "Longest interpolation too long (" & workbookToProcess.Worksheets("HR detection").Cells(i + 1, horizOffset + 3) & ">" & exLongestIntDurGT & ")"
            ElseIf workbookToProcess.Worksheets("HR detection").Cells(i + 1, horizOffset + 4) > exLongestIntBeatsGT And exLongestIntBeatsGT <> -1 Then
                checkForHRExclusions = "Longest interpolation too many beats (" & workbookToProcess.Worksheets("HR detection").Cells(i + 1, horizOffset + 4) & ">" & exLongestIntBeatsGT & ")"
            End If
End Function

Sub outputTrials(trialTypes As Dictionary, trialType As String, thisAnimalWorksheet As Worksheet)
    Dim arrTrialTypes
    arrTrialTypes = trialTypes.Keys
    
    Dim formatCond As FormatCondition
    
    Dim dictParamSets As Dictionary
    
    Dim arrParamSets
    Dim arrTrials
    Dim arrTrial
    
    Dim iTrialTypeNum As Integer
    Dim iParamSetNum As Integer
    Dim iTrialNum As Integer
    
    Dim iExcelOffset As Long
    iExcelOffset = 1
    
    Dim meanHRChange As Double
    Dim HRChangeVar As Double
    Dim nInMeanSoFar As Integer
    Dim diff As Double
    Dim tStat As Double
    
    Dim HRIncTrials As Integer
    Dim HRDecTrials As Integer
    
    For iTrialTypeNum = 0 To UBound(arrTrialTypes)
        If arrTrialTypes(iTrialTypeNum) = trialType Or trialType = "" Then
            thisAnimalWorksheet.Cells(iExcelOffset, 1) = arrTrialTypes(iTrialTypeNum) & " Trials"
            'thisAnimalWorksheet.Cells(iExcelOffset, 1).Style = "Heading"
            thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Bold = True
            iExcelOffset = iExcelOffset + 1
            Set dictParamSets = trialTypes(arrTrialTypes(iTrialTypeNum))
            arrParamSets = dictParamSets.Keys
            For iParamSetNum = 0 To UBound(arrParamSets)
                thisAnimalWorksheet.Cells(iExcelOffset, 1) = arrParamSets(iParamSetNum)
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Bold = True
                iExcelOffset = iExcelOffset + 1
                thisAnimalWorksheet.Range("A" & iExcelOffset, "G" & iExcelOffset).Font.Italic = True
                thisAnimalWorksheet.Cells(iExcelOffset, 1) = "Date"
                thisAnimalWorksheet.Cells(iExcelOffset, 2) = "HR 10-30s"
                thisAnimalWorksheet.Cells(iExcelOffset, 3) = "HR -4s-0s"
                thisAnimalWorksheet.Cells(iExcelOffset, 4) = "HR 5s-9s"
                thisAnimalWorksheet.Cells(iExcelOffset, 5) = "HR 10-30s exclusion reason"
                thisAnimalWorksheet.Cells(iExcelOffset, 6) = "HR -4s-0s exclusion reason"
                thisAnimalWorksheet.Cells(iExcelOffset, 7) = "HR 5s-9s exclusion reason"
                iExcelOffset = iExcelOffset + 1
                arrTrials = dictParamSets(arrParamSets(iParamSetNum))
                nInMeanSoFar = 0
                meanHRChange = 0
                HRChangeVar = 0
                HRIncTrials = 0
                HRDecTrials = 0
                For iTrialNum = 0 To UBound(arrTrials)
                    arrTrial = arrTrials(iTrialNum)
                    thisAnimalWorksheet.Cells(iExcelOffset, 1) = arrTrial(0)
                    thisAnimalWorksheet.Cells(iExcelOffset, 2) = arrTrial(1)
                    thisAnimalWorksheet.Cells(iExcelOffset, 3) = arrTrial(3)
                    thisAnimalWorksheet.Cells(iExcelOffset, 4) = arrTrial(5)
                    thisAnimalWorksheet.Cells(iExcelOffset, 5) = arrTrial(2)
                    thisAnimalWorksheet.Cells(iExcelOffset, 6) = arrTrial(4)
                    thisAnimalWorksheet.Cells(iExcelOffset, 7) = arrTrial(6)
                    iExcelOffset = iExcelOffset + 1
                    
                    If arrTrial(4) = "" And arrTrial(6) = "" Then
                        'contribute to the mean
                        nInMeanSoFar = nInMeanSoFar + 1
                        diff = arrTrial(3) - arrTrial(5)
                        meanHRChange = meanHRChange + ((diff - meanHRChange) / CDbl(nInMeanSoFar))
                        
                        'check if HR rose or fell
                        If diff > 0 Then
                            HRDecTrials = HRDecTrials + 1
                        Else
                            HRIncTrials = HRIncTrials + 1
                        End If
                    End If
                    
                Next
                
                'calculate variance
                For iTrialNum = 0 To UBound(arrTrials)
                    arrTrial = arrTrials(iTrialNum)
                    If arrTrial(4) = "" And arrTrial(6) = "" Then
                        diff = arrTrial(3) - arrTrial(5)
                        HRChangeVar = HRChangeVar + (meanHRChange - diff) ^ 2
                    End If
                Next
                If nInMeanSoFar > 1 Then
                    HRChangeVar = HRChangeVar / (nInMeanSoFar - 1)
                    tStat = meanHRChange / ((HRChangeVar / nInMeanSoFar) ^ 0.5)
                End If
                
                iExcelOffset = iExcelOffset + 1
                thisAnimalWorksheet.Cells(iExcelOffset, 1) = "N included:"
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Italic = True
                thisAnimalWorksheet.Cells(iExcelOffset, 2) = nInMeanSoFar
                iExcelOffset = iExcelOffset + 1
                thisAnimalWorksheet.Cells(iExcelOffset, 1) = "HR decrease trials:"
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Italic = True
                thisAnimalWorksheet.Cells(iExcelOffset, 2) = HRDecTrials
                iExcelOffset = iExcelOffset + 1
                thisAnimalWorksheet.Cells(iExcelOffset, 1) = "HR increase trials:"
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Italic = True
                thisAnimalWorksheet.Cells(iExcelOffset, 2) = HRIncTrials
                iExcelOffset = iExcelOffset + 1
                thisAnimalWorksheet.Cells(iExcelOffset, 1) = "% decrease trials:"
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Italic = True
                thisAnimalWorksheet.Cells(iExcelOffset, 2) = (HRDecTrials / nInMeanSoFar) * 100
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Style = "Percent"
                Call thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions.Delete
                Call thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions.Add(xlCellValue, xlNotBetween, "15", "85")
                thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(1).Font.Color = percOutside1585FC.Font.Color
                thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(1).Font.ColorIndex = percOutside1585FC.Font.ColorIndex
                thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(1).Interior.Color = percOutside1585FC.Interior.Color
                thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(1).Interior.ColorIndex = percOutside1585FC.Interior.ColorIndex
                Call thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions.Add(xlCellValue, xlNotBetween, "25", "75")
                thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(2).Font.Color = percOutside2575FC.Font.Color
                thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(2).Font.ColorIndex = percOutside2575FC.Font.ColorIndex
                thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(2).Interior.Color = percOutside2575FC.Interior.Color
                thisAnimalWorksheet.Cells(iExcelOffset, 2).FormatConditions(2).Interior.ColorIndex = percOutside2575FC.Interior.ColorIndex
                iExcelOffset = iExcelOffset + 2
                thisAnimalWorksheet.Cells(iExcelOffset, 1) = "Mean change:"
                thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Italic = True
                thisAnimalWorksheet.Cells(iExcelOffset, 2) = meanHRChange
                iExcelOffset = iExcelOffset + 1
                If nInMeanSoFar > 1 Then
                    thisAnimalWorksheet.Cells(iExcelOffset, 1) = "Variance:"
                    thisAnimalWorksheet.Cells(iExcelOffset, 2) = HRChangeVar
                    iExcelOffset = iExcelOffset + 1
                    thisAnimalWorksheet.Cells(iExcelOffset, 1) = "Standard Deviation:"
                    thisAnimalWorksheet.Cells(iExcelOffset, 2) = HRChangeVar ^ 0.5
                    iExcelOffset = iExcelOffset + 1
                    thisAnimalWorksheet.Cells(iExcelOffset, 1) = "Std. Error of Mean:"
                    thisAnimalWorksheet.Cells(iExcelOffset, 2) = ((HRChangeVar / nInMeanSoFar) ^ 0.5)
                    iExcelOffset = iExcelOffset + 1
                    thisAnimalWorksheet.Cells(iExcelOffset, 1) = "T-statistic:"
                    thisAnimalWorksheet.Cells(iExcelOffset, 2) = tStat
                    iExcelOffset = iExcelOffset + 1
                    thisAnimalWorksheet.Cells(iExcelOffset, 1) = "P-value:"
                    thisAnimalWorksheet.Cells(iExcelOffset, 1).Font.Italic = True
                    thisAnimalWorksheet.Cells(iExcelOffset, 2) = "=TDIST(ABS(B" & CStr(iExcelOffset - 1) & ")," & CStr(nInMeanSoFar - 1) & ",1)"
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
                Else
                    thisAnimalWorksheet.Cells(iExcelOffset, 1) = "Additional stats could not be calculated (N=1)"
                End If
                
                
                iExcelOffset = iExcelOffset + 2
            Next
            iExcelOffset = iExcelOffset + 1
        End If
    Next
End Sub


Sub deleteOldWorksheets(thisWorkbook As Workbook)
    Dim i As Integer
    i = 1
    While thisWorkbook.Worksheets.Count > 2
        If thisWorkbook.Worksheets(i).Name <> "Controller" And thisWorkbook.Worksheets(i).Name <> "Output template" Then
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
Function readAmpArrays(ByRef acoAmps, ByRef elAmps, param1 As String, param2 As String, workbookToProcess As Workbook, iCurrBlockNum As Integer) As Boolean
    readAmpArrays = True
'        param1acoLoweramp = workbookToProcess.Worksheets("Settings").Cells(i, 7)
'        param1acoUpperamp = workbookToProcess.Worksheets("Settings").Cells(i, 8)
'        param2acoLoweramp = workbookToProcess.Worksheets("Settings").Cells(i, 11)
'        param2acoUpperamp = workbookToProcess.Worksheets("Settings").Cells(i, 12)
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
    While workbookToProcess.Worksheets("Settings").Cells(iRow, 1).Value <> iCurrBlockNum And workbookToProcess.Worksheets("Settings").Cells(iRow, 1).Value <> ""
        iRow = iRow + 1
    Wend
    If workbookToProcess.Worksheets("Settings").Cells(iRow, 1).Value = "" Then 'check we found the row for the block
        readAmpArrays = False
        Exit Function
    End If
    While workbookToProcess.Worksheets("Settings").Cells(iRow, 1).Value = iCurrBlockNum And workbookToProcess.Worksheets("Settings").Cells(iRow, 1).Value <> ""

        If workbookToProcess.Worksheets("Settings").Cells(iRow, 6).Value = param1 Then
            param1LowerAmp = CDbl(trimAmpTrailingChars(workbookToProcess.Worksheets("Settings").Cells(iRow, 7).Value))
            param1UpperAmp = CDbl(trimAmpTrailingChars(workbookToProcess.Worksheets("Settings").Cells(iRow, 8).Value))
            param2LowerAmp = CDbl(trimAmpTrailingChars(workbookToProcess.Worksheets("Settings").Cells(iRow, 11).Value))
            param2UpperAmp = CDbl(trimAmpTrailingChars(workbookToProcess.Worksheets("Settings").Cells(iRow, 12).Value))
        Else
            param1LowerAmp = CDbl(trimAmpTrailingChars(workbookToProcess.Worksheets("Settings").Cells(iRow, 11).Value))
            param1UpperAmp = CDbl(trimAmpTrailingChars(workbookToProcess.Worksheets("Settings").Cells(iRow, 12).Value))
            param2LowerAmp = CDbl(trimAmpTrailingChars(workbookToProcess.Worksheets("Settings").Cells(iRow, 7).Value))
            param2UpperAmp = CDbl(trimAmpTrailingChars(workbookToProcess.Worksheets("Settings").Cells(iRow, 8).Value))
        End If

        If workbookToProcess.Worksheets("Settings").Cells(iRow, 5).Value = "Acoustic" Then
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
