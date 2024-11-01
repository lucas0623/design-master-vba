VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "libDataString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Library"
Option Explicit
Public Function SortStr(str As String, andOperator As String) As String
    'example: SC1,SB2,SB3,SB3 --> SB2,SB3,SB3,SC1
    Dim arr() As String
    Dim i As Long, arrL As Object

    arr = SplitAndTrimString(str, andOperator)
    Set arrL = CreateObject("System.Collections.ArrayList")
    For i = LBound(arr) To UBound(arr)
        arrL.Add arr(i)
    Next i
    arrL.Sort
    SortStr = JoinTextInArrList(arrL)
End Function

Public Function SortAndRemoveDuplicateStr(str As String, andOperator As String)
    'example: SC1,SB2,SB3,SB3 --> SB2,SB3,SC1
    Dim arr() As String
    Dim i As Long, arrL As Object

    arr = SplitAndTrimString(str, andOperator)
    Set arrL = CreateObject("System.Collections.ArrayList")
    For i = LBound(arr) To UBound(arr)
        If Not arrL.contains(arr(i)) Then arrL.Add arr(i)
    Next i
    arrL.Sort
    SortAndRemoveDuplicateStr = JoinTextInArrList(arrL)
    
End Function
Public Function SplitAndTrimString(str As String, andOperator As String) As String()
    Dim result() As String

    Dim i As Long

    ' Use the Split function to split the input string using the andOperator delimiter
    result = Split(str, andOperator)
    
    ' Trim and store the values in the result array
    For i = LBound(result) To UBound(result)
        result(i) = Trim(result(i))
    Next i
    result(0) = Replace(result(0), "'", "", 1, 1)
    ' Set the output as the result array
    SplitAndTrimString = result
End Function

Public Function JoinText(strArr() As String, Optional delimiter As String = ",")
    Dim i As Long, retStr As String

    For i = LBound(strArr) To UBound(strArr)
        retStr = retStr & delimiter & strArr(i)
    Next i
    retStr = Replace(retStr, delimiter, "", 1, 1)
    JoinText = retStr
End Function

Public Function JoinTextInArrList(strArrList As Object, Optional delimiter As String = ",")
    Dim i As Long, retStr As String

    For i = 0 To strArrList.count - 1
        retStr = retStr & delimiter & strArrList.item(i)
    Next i
    retStr = Replace(retStr, delimiter, "", 1, 1)
    JoinTextInArrList = retStr
End Function
