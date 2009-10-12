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
        readMappingLists = True
    End If
        
End Function


