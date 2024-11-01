VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LibCollectionMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Library"
Public Sub AddColltoColl(coll As Collection, coll2 As Collection)
    Dim i As Long
    For i = 1 To coll2.count
        coll.Add coll2(i)
    Next i
End Sub
Public Function FlattenCollOfColl(collOfColl As Collection) As Collection
    Dim result As Collection
    Dim coll As Variant
    Dim item As Variant
    
    Set result = New Collection
    
    ' Iterate through each collection in the input collection of collections
    For Each coll In collOfColl
        ' Iterate through each item in the current collection
        For Each item In coll
            ' Add the item to the result collection
            result.Add item
        Next item
    Next coll
    
    Set FlattenCollOfColl = result
End Function

Public Function isKeyExist(coll As Collection, sKey As String) As Boolean
    'Input shall be colleciton of OBJECT
    Dim obj As Object
    On Error GoTo NotExist:
    
    If VarType(coll(sKey)) > -1 Then isKeyExist = True
    
    'isKeyExist = True
    Exit Function
NotExist:
    isKeyExist = False
    On Error GoTo -1
    'Debug.Print "Err Num: =" & Err.Number
End Function


Public Function CollToArr1D(coll As Collection) As Variant
    Dim i As Long, arr As Variant
    ReDim arr(coll.count - 1)
    
    For i = 1 To coll.count
        If TypeOf coll(i) Is Object  Then
            Set arr(i - 1) = coll(i)
        Else
            arr(i - 1) = coll(i)
        End If
    Next i
    CollToArr1D = arr
End Function
