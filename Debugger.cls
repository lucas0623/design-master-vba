VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Debugger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Library"

Public Sub Log(ByVal inputVariant As Variant, ByVal variableName As String)
    Dim i As Long, j As Long
    
    'Debug.Print "Logging variable: " & variableName
    
    ' Check if the input is an array
    If IsArray(inputVariant) Then
        ' Determine the number of dimensions
        Dim dimensions As Integer
        dimensions = GetArrayDimensions(inputVariant)
        
        If dimensions = 1 Then
            ' Handle a 1D array
            For i = LBound(inputVariant) To UBound(inputVariant)
                Debug.Print variableName & "(" & i & "): " & FormatValue(inputVariant(i))
            Next i
        ElseIf dimensions = 2 Then
            ' Handle a 2D array
            For i = LBound(inputVariant, 1) To UBound(inputVariant, 1)
                For j = LBound(inputVariant, 2) To UBound(inputVariant, 2)
                    Debug.Print variableName & "(" & i & "," & j & "): " & FormatValue(inputVariant(i, j))
                Next j
            Next i
        Else
            Debug.Print "Input array has more than 2 dimensions, which is not supported."
        End If
        
    ' Check if the input is a collection
    ElseIf TypeOf inputVariant Is Collection Then
        Dim item As Variant
        Dim index As Long
        index = 1
        For Each item In inputVariant
            Debug.Print variableName & "(" & index & "): " & FormatValue(item)
            index = index + 1
        Next item
        
    ' Check if the input is a dictionary
    ElseIf TypeName(inputVariant) = "Dictionary" Then
        Dim key As Variant
        For Each key In inputVariant.keys
            Debug.Print variableName & "(" & key & "): " & FormatValue(inputVariant(key))
        Next key

    ' Handle other types if necessary
    
    ElseIf IsObject(inputVariant) Then
        Debug.Print "Input is neither a simple variable, array, a collection, nor a dictionary."
    Else
        Debug.Print variableName & ": " & inputVariant
    End If
End Sub

' Helper function to determine the number of dimensions in an array
Private Function GetArrayDimensions(ByVal arr As Variant) As Integer
    Dim dimCount As Integer
    On Error Resume Next
    Do
        dimCount = dimCount + 1
    Loop Until IsEmpty(LBound(arr, dimCount))
    GetArrayDimensions = dimCount - 1
    On Error GoTo 0
End Function

' Helper function to format values for logging
Private Function FormatValue(value As Variant) As String
    If IsObject(value) Then
        On Error Resume Next
        FormatValue = TypeName(value) & " object"
        On Error GoTo 0
    ElseIf IsArray(value) Then
        FormatValue = "Array"
    Else
        FormatValue = CStr(value)
    End If
End Function
