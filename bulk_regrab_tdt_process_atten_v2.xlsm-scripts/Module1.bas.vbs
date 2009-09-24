Attribute VB_Name = "Module1"
Option Explicit

Sub reprocess()
    Dim regenerateExcelFiles As Boolean

    Dim thisWorkbook As Workbook
    Set thisWorkbook = ActiveWorkbook

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
                                'newWorkbook.Worksheets("Settings").Range("O2:Q173").Value = workbookToProcess.Worksheets("Settings").Range("O2:Q173").Value
                                'newWorkbook.Worksheets("HR detection").Range("A3:Q83").Value = workbookToProcess.Worksheets("HR detection").Range("A3:Q83").Value
                                
                                strUsedRange = workbookToProcess.Worksheets("Beat points from LabChart").UsedRange.Address
                                newWorkbook.Worksheets("Beat points from LabChart").Range(strUsedRange).Value = workbookToProcess.Worksheets("Beat points from LabChart").Range(strUsedRange).Value
                                strUsedRange = workbookToProcess.Worksheets("Trial points from LabChart").UsedRange.Address
                                newWorkbook.Worksheets("Trial points from LabChart").Range(strUsedRange).Value = workbookToProcess.Worksheets("Trial points from LabChart").Range(strUsedRange).Value
                                                                
                                Call workbookToProcess.Close
                                If Not objFS.FolderExists(objExpFolder.Path & "\xls backups") Then
                                    Call objFS.CreateFolder(objExpFolder.Path & "\xls backups")
                                End If
                                
                                newFilename = objExpFolder.Path & "\xls backups\" & Left(strExcelFilename, Len(strExcelFilename) - 5) & Year(Now()) & "-" & Month(Now()) & "-" & Day(Now()) & "_" & Hour(Now()) & "-" & Minute(Now()) & "-" & Second(Now()) & ".XLSM"
                                Call objFS.CopyFile(strExcelPathname, newFilename)
                    
                                                                
                                Call newWorkbook.SaveAs(objExpFolder.Path & "\" & objExpFolder.Name & ".xlsm", 52)
                                strExcelFilename = objExpFolder.Name & ".xlsm"

                                Set workbookToProcess = newWorkbook
                            Else
                                Set workbookToProcess = Workbooks.Open(strExcelPathname)
                            End If
                            
                            workbookToProcess.Activate
                            workbookToProcess.Worksheets("Attenuations").Range("B2:B44301").Value = thisWorkbook.Worksheets("Attenuation").Range("B1:B44300").Value
                            If getTDTTank(objExpFolder, tankFilename, blockName) Then
                                workbookToProcess.Worksheets("Variables (do not edit)").Range("B2").Value = tankFilename
                                workbookToProcess.Worksheets("Variables (do not edit)").Range("B3").Value = blockName
                                Application.Run ("'" & strExcelFilename & "'!importTrialsFromLabchart")
                                Application.Run ("'" & strExcelFilename & "'!processHeartRate")
                                Application.Run ("'" & strExcelFilename & "'!buildDeadzoneLists")
                                Call workbookToProcess.Save
                                Call workbookToProcess.Close
                            End If
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
