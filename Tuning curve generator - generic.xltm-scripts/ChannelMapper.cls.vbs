Attribute VB_Name = "ChannelMapper"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public fwdLookupDict As Dictionary
Public revLookupDict As Dictionary
Public hasValidMap As Boolean
Const mapFilenamePrefix = "Channel map"

Public Function fwdLookup(srchFor As Long) As Variant
    If Not hasValidMap Then
        fwdLookup = srchFor
    ElseIf fwdLookupDict.Exists(srchFor) Then
        fwdLookup = fwdLookupDict(srchFor)
    Else
        fwdLookup = vbEmpty
    End If
End Function

Public Function revLookup(srchFor As Long) As Variant
    If Not hasValidMap Then
        revLookup = srchFor
    ElseIf revLookupDict.Exists(srchFor) Then
        revLookup = revLookupDict(srchFor)
    Else
        revLookup = vbEmpty
    End If
End Function

Public Function Add(fwdVal As Variant, revVal As Variant) As Boolean
    If Not hasValidMap Then
        Add = False
    ElseIf fwdLookupDict.Exists(fwdVal) Then
        Add = False
    ElseIf revLookupDict.Exists(revVal) Then
        Add = False
    Else
        Call fwdLookupDict.Add(fwdVal, revVal)
        Call revLookupDict.Add(revVal, fwdVal)
        Add = True
    End If

End Function

Public Function readMappingLists(rFirstTDTEntry As Range, rFirstMappedEntry As Range, lNumOfChans As Long) As Boolean
    Dim lIter As Long
    readMappingLists = True
    Set fwdLookupDict = Nothing
    Set fwdLookupDict = New Dictionary
    Set revLookupDict = Nothing
    Set revLookupDict = New Dictionary
    lIter = 0
    Do
        If Not (rFirstTDTEntry.Offset(lIter, 0).Value <> "" And rFirstMappedEntry.Offset(lIter, 0).Value <> "") Then
            If lNumOfChans > (lIter + 1) Then
                readMappingLists = False
            End If
            Exit Do
        End If
        
        If Not IsNumeric(rFirstTDTEntry.Offset(lIter, 0).Value) And _
            IsNumeric(rFirstMappedEntry.Offset(lIter, 0).Value) And _
            Int(rFirstTDTEntry.Offset(lIter, 0).Value) = rFirstTDTEntry.Offset(lIter, 0).Value And _
            Int(rFirstMappedEntry.Offset(lIter, 0).Value) = rFirstMappedEntry.Offset(lIter, 0).Value Then
                readMappingLists = False
                Exit Do
        End If
            
        Call fwdLookupDict.Add(rFirstTDTEntry.Offset(lIter, 0).Value, rFirstMappedEntry.Offset(lIter, 0).Value)
        Call revLookupDict.Add(rFirstMappedEntry.Offset(lIter, 0).Value, rFirstTDTEntry.Offset(lIter, 0).Value)
        lIter = lIter + 1
    Loop
        
    If readMappingLists = True Then
        hasValidMap = True
    Else
        Set fwdLookupDict = Nothing
        Set revLookupDict = Nothing
        hasValidMap = False
'        readMappingLists = True
    End If
        
End Function


Public Function readMappingListsFromFilename(sFilename As String, lNumOfChans As Long, Optional rFirstTDTEntry As Range, Optional rFirstMappedEntry As Range) As Boolean
    Dim objFS As FileSystemObject
    Set objFS = New FileSystemObject
    
    If Not objFS.FileExists(sFilename) Then
        readMappingListsFromFilename = False
        hasValidMap = False
    Else
        Dim objFile As File
        Set objFile = objFS.GetFile(sFilename)
        readMappingListsFromFile = readMappingListsFromFile(objFile, lNumOfChans, rFirstTDTEntry, rFirstMappedEntry)
        Set objFile = Nothing
    End If
    
    Set objFS = Nothing

End Function

Public Function readMappingListsFromFile(objFile As File, lNumOfChans As Long, Optional rFirstTDTEntry As Range, Optional rFirstMappedEntry As Range) As Boolean
    
    Dim lIter As Long
    readMappingListsFromFile = True
    Set fwdLookupDict = Nothing
    Set fwdLookupDict = New Dictionary
    Set revLookupDict = Nothing
    Set revLookupDict = New Dictionary
    lIter = 0
    
    Dim objTxt As TextStream
    Set objTxt = objFile.OpenAsTextStream(ForReading)
    
    Dim strLine As String
    Dim arrComponents As Variant
    
    Dim TDTChanCol As Integer
    TDTChanCol = 1
    
    Dim parsedHeader As Boolean
    
    
    Do
        If objTxt.AtEndOfStream Then
            If lNumOfChans > (lIter + 1) Then
                readMappingListsFromFile = False
            End If
            Exit Do
        End If
        strLine = objTxt.ReadLine
        arrComponents = Split(strLine, Chr(9), -1, vbTextCompare)
        If Not UBound(arrComponents) = 1 Then
            If lNumOfChans > (lIter + 1) Then
                readMappingListsFromFile = False
            End If
            Exit Do
        End If
        
        If Not IsNumeric(arrComponents(0)) Or Not IsNumeric(arrComponents(1)) Then 'check they are numeric
            If lIter = 0 Then
                If UCase(arrComponents(0)) = "TDT" Then
                    TDTChanCol = 1
                ElseIf UCase(arrComponents(1)) = "TDT" Then
                    TDTChanCol = 2
                Else
                    readMappingListsFromFile = False
                    Exit Do
                End If
            Else
                If lNumOfChans > (lIter + 1) Then
                    readMappingListsFromFile = False
                End If
                Exit Do
            End If
        Else
            If Not Int(arrComponents(0)) = arrComponents(0) And Int(arrComponents(1)) = arrComponents(0) Then 'check they are integers
                If lNumOfChans > (lIter + 1) Then
                    readMappingListsFromFile = False
                End If
                Exit Do
            End If
               
            If TDTChanCol = 1 Then
                Call fwdLookupDict.Add(Int(arrComponents(0)), Int(arrComponents(1)))
                Call revLookupDict.Add(Int(arrComponents(1)), Int(arrComponents(0)))
            Else
                Call fwdLookupDict.Add(Int(arrComponents(1)), Int(arrComponents(0)))
                Call revLookupDict.Add(Int(arrComponents(0)), Int(arrComponents(1)))

            End If
            lIter = lIter + 1
        End If
    Loop
        
    If readMappingListsFromFile = True Then
        hasValidMap = True
        If Not IsMissing(rFirstTDTEntry) And Not IsMissing(rFirstMappedEntry) Then
            Dim iIter As Integer
            Dim vKeys As Variant
            vKeys = fwdLookupDict.Keys
            For iIter = LBound(vKeys) To UBound(vKeys)
                rFirstTDTEntry.Offset(iIter, 0).Value = vKeys(iIter)
                rFirstMappedEntry.Offset(iIter, 0).Value = fwdLookupDict(vKeys(iIter))
            Next
        End If
    Else
        Set fwdLookupDict = Nothing
        Set revLookupDict = Nothing
        hasValidMap = False
'        readMappingListsFromFile = True
    End If
        
End Function

Public Function readMappingListsFromDirName(sDirName As String, lNumOfChans As Long, Optional rFirstTDTEntry As Range, Optional rFirstMappedEntry As Range) As Boolean
    Dim objFS As FileSystemObject
    Set objFS = New FileSystemObject
    
    Dim blnFoundFile As Boolean
    
    If Not objFS.FolderExists(sDirName) Then
        blnFoundFile = False
    Else
        blnFoundFile = False
        Dim objFolder As Folder
        Set objFolder = objFS.GetFolder(sDirName)
        Dim objFiles As Files
        Dim objFile As File
        
        Set objFiles = objFolder.Files
        
        For Each objFile In objFiles
            If LCase(Left(objFile.Name, Len(mapFilenamePrefix))) = LCase(mapFilenamePrefix) Then
                readMappingListsFromDirName = readMappingListsFromFile(objFile, lNumOfChans, rFirstTDTEntry, rFirstMappedEntry)
                blnFoundFile = True
                Exit For
            End If
        Next
        
        Set objFile = Nothing
        Set objFiles = Nothing
        Set objFolder = Nothing
    End If
    
    Set objFS = Nothing

    If Not blnFoundFile Then
        readMappingListsFromDirName = False
        hasValidMap = False
    End If

End Function


