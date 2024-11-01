Attribute VB_Name = "Module1"
'@Folder("Tests.TestModule")
Function AreArraysIdentical(arr1 As Variant, arr2 As Variant) As Boolean
    Dim dict As Object
    Dim i As Long
    Dim key As Variant
    
    ' Check if both arrays have the same number of elements
    If UBound(arr1) <> UBound(arr2) Then
        AreArraysIdentical = False
        Exit Function
    End If
    
    ' Create a dictionary to store the count of each element in arr1
    Set dict = CreateObject("Scripting.Dictionary")
    
    For i = LBound(arr1) To UBound(arr1)
        key = arr1(i)
        If dict.Exists(key) Then
            dict(key) = dict(key) + 1
        Else
            dict.Add key, 1
        End If
    Next i
    
    ' Compare the count of each element in arr2 to the count in the dictionary
    For i = LBound(arr2) To UBound(arr2)
        key = arr2(i)
        If dict.Exists(key) Then
            dict(key) = dict(key) - 1
            If dict(key) < 0 Then
                AreArraysIdentical = False
                Exit Function
            End If
        Else
            AreArraysIdentical = False
            Exit Function
        End If
    Next i
    
    ' If all elements match, the arrays are identical
    AreArraysIdentical = True
End Function
Sub TestArrays()
    Dim arr1() As Variant
    Dim arr2() As Variant
    
    arr1 = Array("A", "B", "C", "D", "A")
    arr2 = Array("D", "C", "A", "B", "A")
    
    If AreArraysIdentical(arr1, arr2) Then
        MsgBox "The arrays are identical."
    Else
        MsgBox "The arrays are not identical."
    End If
End Sub

Sub TestGetArr()
    Dim genFunc As New clsGeneralFunctions
    Dim wsInteract As New clsWorksheetsInteraction
    Dim arr As Variant
    arr = wsInteract.SetArrayOneDim(4, 5)
    
End Sub

