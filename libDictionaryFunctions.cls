VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "libDictionaryFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Library")

Option Explicit

Public Function GetDuplicateItemsInTwoDicts(dict1 As Object, dict2 As Object) As Object
    Dim key As Variant
    Dim dictRet As Object
    Set dictRet = CreateObject("Scripting.Dictionary")
    For Each key In dict1.keys
        If dict2.exists(key) Then
            dictRet.Add key, dict1(key)
        End If
    Next key
    Set GetDuplicateItemsInTwoDicts = dictRet
End Function

Public Function GetDuplicateItemsInDicts(ParamArray dict() As Variant) As Object
    Dim i As Long, count As Long
    Dim lLimit As Long, uLimit As Long
    Dim dictRet As Object
    Dim dict1 As Object, dict2 As Object
    lLimit = LBound(dict)
    uLimit = UBound(dict)
    count = uLimit - lLimit + 1
    If count = 0 Then
        Exit Function
    ElseIf count = 1 Then
        Set dictRet = dict(lLimit)
    ElseIf count = 2 Then
        Set dict1 = dict(lLimit)
        Set dict2 = dict(lLimit + 1)
        Set dictRet = GetDuplicateItemsInTwoDicts(dict1, dict2)
    Else
        Set dict1 = dict(lLimit)
        Set dict2 = dict(lLimit + 1)
        Set dictRet = GetDuplicateItemsInTwoDicts(dict1, dict2)
        For i = LBound(dict) + 2 To UBound(dict)
            Set dict2 = dict(i)
            Set dictRet = GetDuplicateItemsInTwoDicts(dictRet, dict2)
        Next i
    End If
    
    Set GetDuplicateItemsInDicts = dictRet
End Function
