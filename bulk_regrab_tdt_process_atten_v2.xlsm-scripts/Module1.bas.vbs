Attribute VB_Name = "Module1"
Option Explicit


Sub reprocess(isTestRun As Boolean, onlyOne As Boolean)
    
    Application.Calculation = xlCalculationManual
    
    Dim regenerateExcelFiles As Boolean
    Dim regenerateTDTdata As Boolean
    Dim regenerateDropoutData As Boolean
    Dim regenerateHRcalculations As Boolean
    Dim regenerateNeuralData As Boolean
    Dim doNeuralPlots As Boolean
    Dim updateAttenData As Boolean
    Dim DriveDetect_ActivityDifferenceThreshold As Double '= CDbl(Worksheets("Settings").Range("B37").Value)
    Dim DriveDetect_AbsoluteMinimumSpikesInFirstBin As Long '= CLng(Worksheets("Settings").Range("B38").Value)

    Dim dblTotalWidthSecs As Double
    Dim dblBinWidthSecs As Double
    Dim dblStartOffsetSecs As Double
    Dim iNumOfChans As Integer
    Dim sOnlyIncludeChannels As String
    Dim blnExcludeUndrivenChannels As Boolean

    Dim maxAllowVariation As Double
    Dim minAcceptableHR As Integer
    Dim maxAcceptableHR As Integer
    
    Dim maxPercOfBeatsInt As Double
    Dim maxSingleIntSamples As Double
    Dim maxSingleIntBeats As Double
    
    Dim thisWorkbook As Workbook
    Set thisWorkbook = ActiveWorkbook

    maxAllowVariation = Worksheets("Configuration").Cells(7, 2).Value
    minAcceptableHR = Worksheets("Configuration").Cells(4, 2).Value
    maxAcceptableHR = Worksheets("Configuration").Cells(5, 2).Value

    maxPercOfBeatsInt = Worksheets("Configuration").Cells(16, 2).Value
    maxSingleIntSamples = Worksheets("Configuration").Cells(17, 2).Value
    maxSingleIntBeats = Worksheets("Configuration").Cells(18, 2).Value

    dblTotalWidthSecs = CDbl(Worksheets("Configuration").Range("B25").Value)
    dblBinWidthSecs = CDbl(Worksheets("Configuration").Range("B26").Value)
    dblStartOffsetSecs = CDbl(Worksheets("Configuration").Range("B27").Value)
    iNumOfChans = CInt(Worksheets("Configuration").Range("B28").Value)
    sOnlyIncludeChannels = Worksheets("Configuration").Range("B30").Value
    blnExcludeUndrivenChannels = CBool(Worksheets("Configuration").Range("B33").Value)
    DriveDetect_ActivityDifferenceThreshold = CDbl(Worksheets("Settings").Range("B36").Value)
    DriveDetect_AbsoluteMinimumSpikesInFirstBin = CLng(Worksheets("Settings").Range("B37").Value)


    Dim templateFilename As String
    templateFilename = "\Code current\Excel tools\Tank trial importer.xltm"
    Dim newFilename As String
    
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Dim objFS As FileSystemObject
    Set objFS = CreateObject("Scripting.FileSystemObject")
    
    Dim pathToData As String
    pathToData = objFS.GetDriveName(ActiveWorkbook.FullName) & thisWorkbook.Worksheets("Configuration").Cells(1, 2).Value
    
    'get the root folder under which all data is housed
    Dim rootFolder As Folder
    Set rootFolder = objFS.GetFolder(pathToData)
           
    templateFilename = objFS.GetDriveName(ActiveWorkbook.FullName) & templateFilename 'get the drive letter for the template

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
        
    regenerateExcelFiles = thisWorkbook.Worksheets("Configuration").Cells(2, 2).Value
    regenerateTDTdata = thisWorkbook.Worksheets("Configuration").Cells(9, 2).Value
    regenerateDropoutData = thisWorkbook.Worksheets("Configuration").Cells(10, 2).Value
    regenerateHRcalculations = thisWorkbook.Worksheets("Configuration").Cells(11, 2).Value
    updateAttenData = thisWorkbook.Worksheets("Configuration").Cells(12, 2).Value
    regenerateNeuralData = thisWorkbook.Worksheets("Configuration").Cells(21, 2).Value
    doNeuralPlots = thisWorkbook.Worksheets("Configuration").Cells(22, 2).Value
    
    Dim tankFilename As String
    Dim blockName As String
    Dim strUsedRange As String
    
    Dim workbookToProcess As Workbook
    Dim newWorkbook As Workbook
    Set AnimalFolders = rootFolder.Subfolders
    For Each objAnimalFolder In AnimalFolders 'cycle through the folder for each animal
        If Not checkForExclusion(objAnimalFolder) Then
            Set experimentFolders = objAnimalFolder.Subfolders
            For Each objExpFolder In experimentFolders 'go through the experiments within an animal folder
                If Not checkForExclusion(objExpFolder) Then
                    blnCurrFolderIsTrial = False 'assume until file detected this is not an experiment
                    strExcelFilename = ""
                    tankFilename = ""
                    blockName = ""
                    Set Files = objExpFolder.Files
                    For Each objFile In Files 'inside the experiment file, check for a file ending with ".adicht" to check it is a test
                        If UCase(Right(objFile.Name, 7)) = ".ADICHT" Then
                            blnCurrFolderIsTrial = True
                        End If
                        If UCase(Right(objFile.Name, 5)) = ".XLSM" And Left(objFile.Name, 1) <> "~" And UCase(Left(Right(objFile.Name, 13), 8)) <> "_UPDATED" Then 'locate the spreadsheet file with trial data
                            strExcelFilename = objFile.Name
                            strExcelPathname = objFile.Path
                        End If
                    Next
                    
                    If blnCurrFolderIsTrial Then
                        If strExcelFilename <> "" Then
                            'there is already an excel file - update it
                            If regenerateExcelFiles Then 'do we need to update from the template?
                                Set workbookToProcess = Workbooks.Open(strExcelPathname)
                                Set newWorkbook = Workbooks.Open(templateFilename)
                                'newWorkbook.Worksheets("Output").Range("O2:Q173").Value = workbookToProcess.Worksheets("Output").Range("O2:Q173").Value
                                'newWorkbook.Worksheets("HR detection").Range("A3:Q83").Value = workbookToProcess.Worksheets("HR detection").Range("A3:Q83").Value
                                
                                strUsedRange = workbookToProcess.Worksheets("Beat points from LabChart").UsedRange.Address
                                newWorkbook.Worksheets("Beat points from LabChart").Range(strUsedRange).Value = workbookToProcess.Worksheets("Beat points from LabChart").Range(strUsedRange).Value
                                strUsedRange = workbookToProcess.Worksheets("Trial points from LabChart").UsedRange.Address
                                newWorkbook.Worksheets("Trial points from LabChart").Range(strUsedRange).Value = workbookToProcess.Worksheets("Trial points from LabChart").Range(strUsedRange).Value
                                                                
                                If Not regenerateTDTdata Then
                                    newWorkbook.Worksheets("Output").Range("A2:N173").Value = workbookToProcess.Worksheets("Output").Range("A2:N173").Value
                                End If
                                
                                If Not regenerateDropoutData Then
                                    strUsedRange = workbookToProcess.Worksheets("Deadzones").UsedRange.Address
                                    newWorkbook.Worksheets("Deadzones").Range(strUsedRange).Value = workbookToProcess.Worksheets("Deadzones").Range(strUsedRange).Value
                                End If
                                                                
                                If Not regenerateHRcalculations Then
                                    strUsedRange = workbookToProcess.Worksheets("Interpolations").UsedRange.Address
                                    newWorkbook.Worksheets("Interpolations").Range(strUsedRange).Value = workbookToProcess.Worksheets("Interpolations").Range(strUsedRange).Value
                                    
                                    strUsedRange = workbookToProcess.Worksheets("Overbeats").UsedRange.Address
                                    newWorkbook.Worksheets("Overbeats").Range(strUsedRange).Value = workbookToProcess.Worksheets("Overbeats").Range(strUsedRange).Value
                                    
                                    strUsedRange = workbookToProcess.Worksheets("Abberant beats").UsedRange.Address
                                    newWorkbook.Worksheets("Abberant beats").Range(strUsedRange).Value = workbookToProcess.Worksheets("Abberant beats").Range(strUsedRange).Value
                                    
                                    strUsedRange = workbookToProcess.Worksheets("HR detection").UsedRange.Address
                                    newWorkbook.Worksheets("HR detection").Range(strUsedRange).Value = workbookToProcess.Worksheets("HR detection").Range(strUsedRange).Value
                                    
                                    strUsedRange = workbookToProcess.Worksheets("-84-4s HRs").UsedRange.Address
                                    newWorkbook.Worksheets("-84-4s HRs").Range(strUsedRange).Value = workbookToProcess.Worksheets("-84-4s HRs").Range(strUsedRange).Value
                                    newWorkbook.Worksheets("-84-4s HRs").Range(strUsedRange).Formula = workbookToProcess.Worksheets("-84-4s HRs").Range(strUsedRange).Formula
                                    
                                    strUsedRange = workbookToProcess.Worksheets("-4-0s HRs").UsedRange.Address
                                    newWorkbook.Worksheets("-4-0s HRs").Range(strUsedRange).Value = workbookToProcess.Worksheets("-4-0s HRs").Range(strUsedRange).Value
                                    newWorkbook.Worksheets("-4-0s HRs").Range(strUsedRange).Formula = workbookToProcess.Worksheets("-4-0s HRs").Range(strUsedRange).Formula
                                    
                                    strUsedRange = workbookToProcess.Worksheets("5-9s HRs").UsedRange.Address
                                    newWorkbook.Worksheets("5-9s HRs").Range(strUsedRange).Value = workbookToProcess.Worksheets("5-9s HRs").Range(strUsedRange).Value
                                    newWorkbook.Worksheets("5-9s HRs").Range(strUsedRange).Formula = workbookToProcess.Worksheets("5-9s HRs").Range(strUsedRange).Formula
                                    
                                    strUsedRange = workbookToProcess.Worksheets("-4.5-9.5s HRs").UsedRange.Address
                                    newWorkbook.Worksheets("-4.5-9.5s HRs").Range(strUsedRange).Value = workbookToProcess.Worksheets("-4.5-9.5s HRs").Range(strUsedRange).Value
                                    newWorkbook.Worksheets("-4.5-9.5s HRs").Range(strUsedRange).Formula = workbookToProcess.Worksheets("-4.5-9.5s HRs").Range(strUsedRange).Formula
                                    
                                    strUsedRange = workbookToProcess.Worksheets("HRLine").UsedRange.Address
                                    newWorkbook.Worksheets("HRLine").Range(strUsedRange).Value = workbookToProcess.Worksheets("HRLine").Range(strUsedRange).Value
                                    newWorkbook.Worksheets("HRLine").Range(strUsedRange).Formula = workbookToProcess.Worksheets("HRLine").Range(strUsedRange).Formula
                                    
                                    newWorkbook.Worksheets("Output").Range("O2:Q173").Value = workbookToProcess.Worksheets("Output").Range("O2:Q173").Value
                                End If
                                                                
                                                                
                                Call workbookToProcess.Close
                                If Not objFS.FolderExists(objExpFolder.Path & "\xls backups") Then
                                    Call objFS.CreateFolder(objExpFolder.Path & "\xls backups")
                                End If
                                
                                newFilename = objExpFolder.Path & "\xls backups\" & Left(strExcelFilename, Len(strExcelFilename) - 5) & Year(Now()) & "-" & Month(Now()) & "-" & Day(Now()) & "_" & Hour(Now()) & "-" & Minute(Now()) & "-" & Second(Now()) & ".XLSM"
                                If Not isTestRun Then
                                    Call objFS.CopyFile(strExcelPathname, newFilename)
                                                                
                                    Call newWorkbook.SaveAs(objExpFolder.Path & "\" & objExpFolder.Name & ".xlsm", 52)
                                    strExcelFilename = objExpFolder.Name & ".xlsm"
                                Else
                                    strExcelFilename = newWorkbook.Name
                                End If

                                Set workbookToProcess = newWorkbook
                            Else
                                Set workbookToProcess = Workbooks.Open(strExcelPathname)
                            End If
                            
                            workbookToProcess.Activate
                            workbookToProcess.Worksheets("Settings").Cells(5, 2).Value = maxAllowVariation
                            workbookToProcess.Worksheets("Settings").Cells(2, 2).Value = minAcceptableHR
                            workbookToProcess.Worksheets("Settings").Cells(3, 2).Value = maxAcceptableHR
                            
                            workbookToProcess.Worksheets("Settings").Cells(9, 2).Value = maxPercOfBeatsInt
                            workbookToProcess.Worksheets("Settings").Cells(10, 2).Value = maxSingleIntSamples
                            workbookToProcess.Worksheets("Settings").Cells(11, 2).Value = maxSingleIntBeats
                            
                            workbookToProcess.Worksheets("Settings").Range("B20").Value = dblTotalWidthSecs
                            workbookToProcess.Worksheets("Settings").Range("B21").Value = dblBinWidthSecs
                            workbookToProcess.Worksheets("Settings").Range("B22").Value = dblStartOffsetSecs
                            workbookToProcess.Worksheets("Settings").Range("B23").Value = iNumOfChans
                            workbookToProcess.Worksheets("Settings").Range("B25").Value = sOnlyIncludeChannels
                            workbookToProcess.Worksheets("Settings").Range("B34").Value = blnExcludeUndrivenChannels
                            
                            workbookToProcess.Worksheets("Settings").Range("B37").Value = DriveDetect_ActivityDifferenceThreshold
                            workbookToProcess.Worksheets("Settings").Range("B38").Value = DriveDetect_AbsoluteMinimumSpikesInFirstBin

                            
                            If updateAttenData Then
                                workbookToProcess.Worksheets("Attenuations").Range("B2:B44301").Value = thisWorkbook.Worksheets("Attenuation").Range("B1:B44300").Value
                            End If
                            
                            If getTDTTank(objExpFolder, tankFilename, blockName) Then
                                workbookToProcess.Worksheets("Variables (do not edit)").Range("B2").Value = tankFilename
                                workbookToProcess.Worksheets("Variables (do not edit)").Range("B3").Value = blockName
                                If regenerateTDTdata Then
                                    Application.Run ("'" & strExcelFilename & "'!importTrialsFromLabchart")
                                End If
                            End If
                            
                            If regenerateHRcalculations Then
                                Application.Run ("'" & strExcelFilename & "'!processHeartRate")
                                Application.Run ("'" & strExcelFilename & "'!generateHrAtTimePoints")
                            End If
                            If regenerateDropoutData Then
                                Application.Run ("'" & strExcelFilename & "'!buildDeadzoneLists")
                            End If
                            
                            If regenerateNeuralData Then
                                If doNeuralPlots Then
                                    Application.Run ("'" & strExcelFilename & "'!ExtractNeuralDataWithCharts")
                                Else
                                    Application.Run ("'" & strExcelFilename & "'!ExtractNeuralDataWithoutCharts")
                                End If
                            End If
                            
                            If Not isTestRun Then
                                Call workbookToProcess.Save
                            End If
                            If onlyOne Then
                                GoTo outsideTheLoop
                            Else
                                Call workbookToProcess.Close
                            End If
                        End If
                    End If
                End If
            Next
        End If
    Next
    
outsideTheLoop:

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

Function getTDTTank(objCurrFolder As Folder, ByRef tankFilename, ByRef blockName)
    tankFilename = ""
    blockName = ""

    Dim objFS As FileSystemObject
    Set objFS = New FileSystemObject
    
    Dim objFolder As Folder
    
    Set objFolder = objCurrFolder

    Dim Subfolders As Folders
    Dim objSubfolder As Folder
    
    Dim BlockFolders As Folders
    Dim objBlockFolder As Folder
    
    Dim Files As Files
    Dim objFile As File
    
    Set Subfolders = objFolder.Subfolders
    For Each objSubfolder In Subfolders
        Set Files = objSubfolder.Files
        For Each objFile In Files
            If objFile.Name = "desktop.ini" Then
                Set BlockFolders = objSubfolder.Subfolders
                For Each objBlockFolder In BlockFolders
                    If objBlockFolder.Name <> "TempBlk" Then
                        blockName = objBlockFolder.Name
                    End If
                Next
                tankFilename = objSubfolder.Path
                Exit For
            End If
        Next
        If tankFilename <> "" Then
            Exit For
        End If
    Next
    If tankFilename <> "" Then
        getTDTTank = True
    Else
        getTDTTank = False
        MsgBox ("Could not find a TDT tank for " & objCurrFolder.Path)
    End If


    Set objFile = Nothing
    Set objFolder = Nothing
    Set objFS = Nothing

End Function


Function checkForExclusion(objFolder As Folder) As Boolean
    checkForExclusion = False
    Dim Files As Files
    Dim objFile As File

    Set Files = objFolder.Files

    For Each objFile In Files
        If LCase(objFile.Name) = "exclude from bulk reprocessing.txt" Then
            checkForExclusion = True
            Exit For
        End If
    Next

End Function
