VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGeneralFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Library"

'********************************************************
'The module contains some useful functions in macro development
'Arthor: Lucas LEUNG
'Update Log
'30 May 2023 - Minor Updates
'15 Feb 2023 - Add 'Interpolate' Function
'1 Feb 2023 - Minor Update
'17 Jan 2023 - Updated Sorting Function (Quick Sort + Bubble Sort)
'16 Jan 2023 - Add 'sorting' functions (1D and 2D)
'7 Dec 2022 - Minor Update
'28 Nov 2022 - Add Function "Max" and "Min" and "ChangeArrBase", "CreatDictOfObj"
'25 Nov 2022 - Add Function "Condense 2D arr to 1D arr", "expandarrTo2DArr", "Combine2DArr", Modify Function "AddItemToArr"
'21 Nov 2022 - Update join str function
'4 Nov 2022 - Add Functions to convert Variant Array to Dobule/String/Long array
'26 Oct 2022 - Add Functions "getArrDimension" and "Get1DArrFrom2DArr"
'31 Jul 2022 - Add Functions to change varaint array to double/string array

'*******************************************************
'if isRow = true, will search the first row
Public Function IsString(ByVal v As Variant) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when v is neither an object nor numeric.
' ----------------------------------------------------------------------------
    Dim s As String
    On Error Resume Next
    s = v
    If Err.Number = 0 Then
        If Not IsNumeric(v) Then IsString = True
    End If
End Function

Public Function Interpolate_val(x As Double, x1 As Double, y1 As Double, _
                        x2 As Double, y2 As Double) As Double
                        
    Dim ratio As Double
    Dim coor(2) As Double
    s = ((x2 - x1) ^ 2 + (y2 - y1) ^ 2) ^ (1 / 2)
    ratio = (x - x1) / (x2 - x1)
    'coor(0) = x1 + (x2 - x1) * ratio
    Interpolate_val = y1 + (y2 - y1) * ratio
    'coor(2) = z1 + (z2 - z1) * ratio
End Function

Public Function ChangeObjArrBase(arr As Variant, base As Integer) As Variant

    Dim i As Long, tempArr() As Object
    Dim n As Long, count As Long
    n = UBound(arr) - LBound(arr) + 1
    ReDim tempArr(base To base + n - 1)
    For i = LBound(arr) To UBound(arr)
        Set tempArr(base + count) = arr(i)
        count = count + 1
    Next i
    ChangeObjArrBase = tempArr
End Function

Public Function ArrListTo2DArr(arrList As Object) As Variant
    Dim arr_row As Variant
    Dim arr_2D As Variant
    Dim arr_list As Variant
    Dim i As Long, j As Long
    Dim is_ini As Boolean
    
    arr_list = arrList.ToArray
    is_ini = False
    
    For i = LBound(arr_list) To UBound(arr_list)
        arr_row = arr_list(i)
        If Not is_ini Then
    
            ReDim arr_2D(LBound(arr_list) To UBound(arr_list), LBound(arr_row) To UBound(arr_row))
            is_ini = True
        End If
        For j = LBound(arr_row) To UBound(arr_row)
            arr_2D(i, j) = arr_row(j)
        Next j
    Next i
    
    ArrListTo2DArr = arr_2D
End Function

Public Function Compare(ByVal value1 As Double, ByVal value2 As Double, is_ABS As Boolean) As Boolean 'return true if value2 is larger than value1

    If is_ABS = True Then 'turns all the number to positive if is_ABS is true
        
        If value1 < 0 Then
            value1 = -value1
        End If
        
        If value2 < 0 Then
           value2 = -value2
        End If
    End If
        
    If value1 < value2 Then
        Compare = True
        Exit Function
    Else
        Compare = False
        Exit Function
    End If
    
End Function
Public Function isArraysIdentical(arr1 As Variant, arr2 As Variant, Optional isCaseSentitive = True) As Boolean
    Dim dict As Object
    Dim i As Long
    Dim key As Variant
    
    ' Check if both arrays have the same number of elements
    If UBound(arr1) <> UBound(arr2) Then
        isArraysIdentical = False
        Exit Function
    End If
    
    If Not isCaseSentitive Then
        For i = LBound(arr1) To UBound(arr1)
            arr1(i) = UCase(arr1(i))
        Next i
        For i = LBound(arr2) To UBound(arr2)
            arr2(i) = UCase(arr2(i))
        Next i
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
                isArraysIdentical = False
                Exit Function
            End If
        Else
            isArraysIdentical = False
            Exit Function
        End If
    Next i
    
    ' If all elements match, the arrays are identical
    isArraysIdentical = True
End Function

Public Function ListBoxToArray(myListBox As msForms.listBox) As String()
    Dim i As Long, myArr As Variant
    If Not myListBox.ListCount = 0 Then
        ReDim myArr(myListBox.ListCount - 1)
        For i = 0 To myListBox.ListCount - 1
            myArr(i) = myListBox.column(0, i)
        Next i
        ListBoxToArray = CStr_arr(myArr)
    
    End If
End Function

Public Sub AddArrToArrList(arrList As Object, arr As Variant)
    
    Dim i As Long
    If IsArray(arr) Then
        For i = LBound(arr) To UBound(arr)
            arrList.Add CVar(arr(i))
        Next i
    Else
        arrList.Add CVar(arr)
    End If
End Sub

Public Function JoinStr(divider As String, ParamArray str() As Variant) As String
    Dim i As Long
    Dim isInitialised As Boolean
    
    isInitialised = False

    For i = LBound(str) To UBound(str)
        If isInitialised Then
            If Not CStr(str(i)) = vbNullString Then JoinStr = JoinStr & divider & str(i)
        Else
            If Not CStr(str(i)) = vbNullString Then
                JoinStr = str(i)
                isInitialised = True
            End If
        End If
    Next i
    
End Function

'************************************************************************************************************
'*************************************FUNCTIONS FOR ARRAY************************************************
'************************************************************************************************************
Public Function CreateArray(ParamArray arr() As Variant) As Variant
    Dim i As Long, j As Long, arrL As Object
    Dim str2 As String
    Set arrL = CreateObject("System.Collections.ArrayList")
    For i = LBound(arr) To UBound(arr)
        If isInitialised(arr(i)) Then
            For j = LBound(arr(i)) To UBound(arr(i))
                arrL.Add CVar(arr(i)(j))
            Next j
        Else
            arrL.Add arr(i)
        End If
    Next i
    CreateArray = arrL.ToArray
End Function
Public Sub BubbleSortArr2D1Key(ByRef vA As Variant, _
                       Optional ByVal bIsAscending As Boolean = True, _
                       Optional ByVal bIsRowSort As Boolean = True, _
                       Optional ByVal SortIndex As Long = -1, _
                       Optional ByRef vRet As Variant)
'--------------------------------------------------------------------------------
' Procedure : SortArr2D1Key
' Purpose   : Bubblesorts a 2D array on 1 key, up or down, on any column or row.
'             Options include in-place, with the source changed, or
'             returned in vRet, with the source array intact.
'             Optional parameters default to: ROW SORT in place, ASCENDING,
'             using COLUMN ONE as the key.
'--------------------------------------------------------------------------------
    
    Dim condition1 As Boolean, vR As Variant
    Dim i As Long, j As Long, y As Long, t As Variant
    Dim loR As Long, hiR As Long, Loc As Long, hiC As Long
    Dim r As Long, c As Long, bWasMissing As Boolean
    
    'find bounds of vA data input array
    loR = LBound(vA, 1): hiR = UBound(vA, 1)
    Loc = LBound(vA, 2): hiC = UBound(vA, 2)
    
    'find whether optional vR was initially missing
    bWasMissing = IsMissing(vRet)
    'If Not bWasMissing Then Set vRet = Nothing
    
    'check input range of SortIndex
    If bIsRowSort And (SortIndex < Loc Or SortIndex > hiC) Then
        MsgBox "SortIndex out of bounds in Sort2DArr; closing now"
        Exit Sub
    Else:
    End If
    
    If Not bIsRowSort And (SortIndex < loR Or SortIndex > hiR) Then
        MsgBox "SortIndex out of bounds in Sort2DArr; closing now"
        Exit Sub
    Else:
    End If
    
    'pass to a work variable
    vR = vA
    
    'steer input options
    If bIsRowSort Then GoTo ROWSORT Else GoTo COLSORT
    
ROWSORT:
    For i = loR To hiR - 1
        For j = loR To hiR - 1
            If bIsAscending Then
                condition1 = vR(j, SortIndex) > vR(j + 1, SortIndex)
            Else
                condition1 = vR(j, SortIndex) < vR(j + 1, SortIndex)
            End If
            If condition1 Then
                For y = Loc To hiC
                    t = vR(j, y)
                    vR(j, y) = vR(j + 1, y)
                    vR(j + 1, y) = t
                Next y
            End If
        Next
    Next
    GoTo TRANSFERS
    
COLSORT:
    For i = Loc To hiC - 1
        For j = Loc To hiC - 1
            If bIsAscending Then
                condition1 = vR(SortIndex, j) > vR(SortIndex, j + 1)
            Else
                condition1 = vR(SortIndex, j) < vR(SortIndex, j + 1)
            End If
            If condition1 Then
                For y = loR To hiR
                    t = vR(y, j)
                    vR(y, j) = vR(y, j + 1)
                    vR(y, j + 1) = t
                Next y
            End If
        Next
    Next
    GoTo TRANSFERS
    
TRANSFERS:
    'decide whether to return in vA or vRet
    If Not bWasMissing Then
        'vRet was the intended return array
        'so return vRet leaving vA intact
        vRet = vR
    Else:
        'vRet is not intended return array
        'so reload vA with vR
        vA = vR
    End If
    
'    'set return function value
'    SortArr2D1Key = True
    
End Sub

Public Sub QuickSortArray_2D(ByRef SortArray As Variant, Optional lngColumn As Long = 0, Optional lngMin As Long = -1, Optional lngMax As Long = -1)
    On Error Resume Next

    'Sort a 2-Dimensional array

    ' SampleUsage: sort arrData by the contents of column 3
    '
    '   QuickSortArray arrData, , , 3

    '
    'Posted by Jim Rech 10/20/98 Excel.Programming

    'Modifications, Nigel Heffernan:

    '       ' Escape failed comparison with empty variant
    '       ' Defensive coding: check inputs

    Dim i As Long
    Dim j As Long
    Dim varMid As Variant
    Dim arrRowTemp As Variant
    Dim lngColTemp As Long

    If IsEmpty(SortArray) Then
        Exit Sub
    End If
    If InStr(TypeName(SortArray), "()") < 1 Then  'IsArray() is somewhat broken: Look for brackets in the type name
        Exit Sub
    End If
    If lngMin = -1 Then
        lngMin = LBound(SortArray, 1)
    End If
    If lngMax = -1 Then
        lngMax = UBound(SortArray, 1)
    End If
    If lngMin >= lngMax Then    ' no sorting required
        Exit Sub
    End If

    i = lngMin
    j = lngMax

    varMid = Empty
    varMid = SortArray((lngMin + lngMax) \ 2, lngColumn)

    ' We  send 'Empty' and invalid data items to the end of the list:
    If IsObject(varMid) Then  ' note that we don't check isObject(SortArray(n)) - varMid *might* pick up a valid default member or property
        i = lngMax
        j = lngMin
    ElseIf IsEmpty(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf IsNull(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf varMid = vbNullString Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) = vbError Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) > 17 Then
        i = lngMax
        j = lngMin
    End If

    While i <= j
        While SortArray(i, lngColumn) < varMid And i < lngMax
            i = i + 1
        Wend
        While varMid < SortArray(j, lngColumn) And j > lngMin
            j = j - 1
        Wend

        If i <= j Then
            ' Swap the rows
            ReDim arrRowTemp(LBound(SortArray, 2) To UBound(SortArray, 2))
            For lngColTemp = LBound(SortArray, 2) To UBound(SortArray, 2)
                arrRowTemp(lngColTemp) = SortArray(i, lngColTemp)
                SortArray(i, lngColTemp) = SortArray(j, lngColTemp)
                SortArray(j, lngColTemp) = arrRowTemp(lngColTemp)
            Next lngColTemp
            Erase arrRowTemp

            i = i + 1
            j = j - 1
        End If
    Wend

    If (lngMin < j) Then QuickSortArray_2D SortArray, lngColumn, lngMin, j
    If (i < lngMax) Then QuickSortArray_2D SortArray, lngColumn, i, lngMax
    
End Sub

Public Function SortArray(arr As Variant, Optional isA2Z As Boolean = True) As Variant
    If Not isInitialised(arr) Then Exit Function
    Dim arrL As Object, i As Long
    Set arrL = CreateObject("System.Collections.ArrayList")
    For i = LBound(arr) To UBound(arr)
        arrL.Add CStr(arr(i))
    Next i
    arrL.Sort
    If Not isA2Z Then arrL.Reverse
    SortArray = arrL.ToArray
End Function

Public Function TransposeArray(MyArray As Variant) As Variant
    Dim x As Long, y As Long
    Dim maxX As Long, minX As Long
    Dim maxY As Long, minY As Long
    
    Dim tempArr As Variant
    
    'Get Upper and Lower Bounds
    maxX = UBound(MyArray, 1)
    minX = LBound(MyArray, 1)
    maxY = UBound(MyArray, 2)
    minY = LBound(MyArray, 2)
    
    'Create New Temp Array
    ReDim tempArr(minY To maxY, minX To maxX)
    
    'Transpose the Array
    For x = minX To maxX
        For y = minY To maxY
            tempArr(y, x) = MyArray(x, y)
        Next y
    Next x
    
    'Output Array
    TransposeArray = tempArr
    
End Function

Public Function isInitialised(ByRef a As Variant) As Boolean
'This sub check if an ARRAY is initialized.
    isInitialised = False
    On Error GoTo ErrHandler
    If IsArray(a) Then
        If Not UBound(a) = -1 Then
            isInitialised = True
        End If
    ElseIf Not a = vbNullString Then
        isInitialised = True
    End If

    Exit Function
ErrHandler:
    isInitialised = False
End Function

Public Function IsInArr(str As Variant, arr As Variant, Optional isCaseSensitive As Boolean = True) As Boolean
    Dim i As Long
    If Not isCaseSensitive Then str = UCase(str)
    
    If IsArray(arr) Then
        For i = LBound(arr) To UBound(arr)
            If Not isCaseSensitive Then arr(i) = UCase(arr(i))
            If str = arr(i) Then
                IsInArr = True
                Exit Function
            End If
        Next i
    Else
        If Not isCaseSensitive Then arr = UCase(arr)
        If str = arr Then
            IsInArr = True
            Exit Function
        End If
    End If
    IsInArr = False
End Function

Public Function InArr(str As Variant, arr As Variant) As Long
'This Function return the index of certain item in an array
    Dim i As Long
    If IsArray(arr) Then
        For i = LBound(arr) To UBound(arr)
            If str = arr(i) Then
                InArr = i
                Exit Function
            End If
        Next i
    End If
    InArr = -1
End Function

Public Function Max(arr As Variant) As Double
    Dim i As Long
    Max = arr(LBound(arr))
    For i = LBound(arr) To UBound(arr)
        If IsNumeric(arr(i)) And arr(i) > Max Then Max = arr(i)
    Next i
End Function

Public Function Min(arr As Variant) As Double
    Dim i As Long
    Min = arr(LBound(arr))
    For i = LBound(arr) To UBound(arr)
        If IsNumeric(arr(i)) And arr(i) < Min Then Min = arr(i)
    Next i
End Function

Public Function ChangeArrBase(arr As Variant, base As Integer) As Variant
    Dim i As Long, tempArr As Variant
    Dim n As Long, count As Long
    n = UBound(arr) - LBound(arr) + 1
    ReDim tempArr(base To base + n - 1)
    For i = LBound(arr) To UBound(arr)
        tempArr(base + count) = arr(i)
        count = count + 1
    Next i
    ChangeArrBase = tempArr
End Function
'nDim shall be 1 or 2, representing the dimension of the array
'n shall be the rowNum/ Column Num
Public Function Get1DArrFrom2DArr(var As Variant, nDim As Long, n As Long) As Variant
    Dim tempArr As Variant, i As Long
    Dim loopDim As Long
    If nDim = 1 Then
        loopDim = 2
    ElseIf nDim = 2 Then
        loopDim = 1
    Else
        'Debug.Print "Function Get1DArrFrom2DArr: nDim shall be either 1 or 2"
    End If
    
    If getArrDimension(var) > 1 Then
        ReDim tempArr(LBound(var, loopDim) To UBound(var, loopDim))
        
        
            For j = LBound(tempArr) To UBound(tempArr)
                If nDim = 1 Then
                    tempArr(j) = var(n, j)
                Else
                    tempArr(j) = var(j, n)
                End If
            Next j
       
        
        Get1DArrFrom2DArr = tempArr
        
    ElseIf getArrDimension(var) = 1 Then
        Get1DArrFrom2DArr = var
        Exit Function
    End If
    
End Function

Public Function getArrDimension(var As Variant) As Long
    On Error GoTo Err
    Dim i As Long
    Dim tmp As Long
    i = 0
    Do While True
        i = i + 1
        tmp = UBound(var, i)
    Loop
Err:
    getArrDimension = i - 1
End Function

Public Function Condense2DArrTo1D(var As Variant) As Variant
    Dim tempArr As Variant, n1 As Long, n2 As Long, n As Long
    Dim count As Long, i As Long, j As Long
    
    n1 = UBound(var, 1) - LBound(var, 1) + 1
    n2 = UBound(var, 2) - LBound(var, 2) + 1
    n = n1 * n2
    ReDim tempArr(n1 * n2 - 1)
    
    For i = LBound(var, 2) To UBound(var, 2)
        For j = LBound(var, 1) To UBound(var, 1)
            tempArr(count) = var(j, i)
            count = count + 1
        Next j
    Next i
    Condense2DArrTo1D = tempArr
End Function

Public Function Expand1DArrTo2D(var As Variant, Optional base As Long = 0) As Variant
    Dim tempArr As Variant, i As Long
    ReDim tempArr(LBound(var) To UBound(var), base To base)
    For i = LBound(var) To UBound(var)
        tempArr(i, base) = var(i)
    Next i
    Expand1DArrTo2D = tempArr
End Function

Public Function AddItemsToArr(arr As Variant, val As Variant) As Variant
    Dim newArr As Variant
    Dim i As Long, n As Long
    Dim count As Long
    If IsArray(val) Then
        n = UBound(val) - LBound(val) + 1
    Else
        n = 1
    End If
    
    If IsArray(arr) Then
        ReDim newArr(LBound(arr) To UBound(arr) + n)
        For i = LBound(arr) To UBound(arr)
            newArr(i) = arr(i)
        Next i
    Else
        ReDim newArr(n)
        newArr(0) = arr
    End If
    
    'add new data to array
    If IsArray(val) Then
        count = LBound(val)
        For i = UBound(arr) + 1 To UBound(newArr)
            newArr(i) = val(count)
            count = count + 1
        Next i
    Else
        newArr(UBound(newArr)) = val
    End If
    
    AddItemsToArr = newArr
End Function

Public Function Combine2DArr(a As Variant, b As Variant, Optional stacked As Boolean = True) As Variant
    'assumes that A and B are 2-dimensional variant arrays
    'if stacked is true then A is placed on top of B
    'in this case the number of rows must be the same,
    'otherwise they are placed side by side A|B
    'in which case the number of columns are the same
    'LBound can be anything but is assumed to be
    'the same for A and B (in both dimensions)
    'False is returned if a clash

    Dim lb As Long, m_A As Long, n_A As Long
    Dim m_B As Long, n_B As Long
    Dim m As Long, n As Long
    Dim i As Long, j As Long, k As Long
    Dim c As Variant

    If TypeName(a) = "Range" Then a = a.value
    If TypeName(b) = "Range" Then b = b.value

    lb = LBound(a, 1)
    m_A = UBound(a, 1)
    n_A = UBound(a, 2)
    m_B = UBound(b, 1)
    n_B = UBound(b, 2)

    If stacked Then
        m = m_A + m_B + 1 - lb
        n = n_A
        If n_B <> n Then
            Combine = False
            Exit Function
        End If
    Else
        m = m_A
        If m_B <> m Then
            Combine = False
            Exit Function
        End If
        n = n_A + n_B + 1 - lb
    End If
    ReDim c(lb To m, lb To n)
    For i = lb To m
        For j = lb To n
            If stacked Then
                If i <= m_A Then
                    c(i, j) = a(i, j)
                Else
                    c(i, j) = b(lb + i - m_A - 1, j)
                End If
            Else
                If j <= n_A Then
                    c(i, j) = a(i, j)
                Else
                    c(i, j) = b(i, lb + j - n_A - 1)
                End If
            End If
        Next j
    Next i
    Combine2DArr = c
End Function


Public Function CheckDuplicate(arr As Variant) As Boolean
    Dim i As Long, j As Long
    Dim count As Long
    
    For i = LBound(arr) To UBound(arr)
        count = 0
        For j = LBound(arr) To i
            If arr(i) = arr(j) Then
                count = count + 1
            End If
        Next j
        
        If count > 1 Then
            CheckDuplicate = True
            Exit Function
        End If
    Next i
    CheckDuplicate = False
End Function

Public Function DeDupeOneDimArray(vArray As Variant, Optional isExcludeEmpty As Boolean = True) As Variant
  Dim oDict As Object, i As Long
  Set oDict = CreateObject("Scripting.Dictionary")
  For i = LBound(vArray) To UBound(vArray)
    oDict(vArray(i)) = True
  Next i
  
  On Error Resume Next
  If isExcludeEmpty Then oDict.Remove (vbNullString)
  On Error GoTo 0
  
  DeDupeOneDimArray = oDict.keys()
End Function

Public Function DeDupeTwoDimArray(vArray As Variant, Optional isExcludeEmpty As Boolean = True, Optional excludeArr As Variant) As Variant
  Dim oDict As Object, i As Long, j As Long
  Set oDict = CreateObject("Scripting.Dictionary")
  
  For i = LBound(vArray, 1) To UBound(vArray, 1)
    For j = LBound(vArray, 2) To UBound(vArray, 2)
        oDict(vArray(i, j)) = True
    Next j
  Next i
  
  If isExcludeEmpty Then oDict.Remove (vbNullString)
  
  If isInitialised(excludeArr) Then
    For i = LBound(excludeArr) To UBound(excludeArr)
        oDict.Remove (excludeArr(i))
    Next i
  End If
  
  DeDupeTwoDimArray = oDict.keys()
End Function

'************************************************************************************************************
'*************************************FUNCTIONS FOR DICTIONARY************************************************
'************************************************************************************************************
Public Function DictToArrObj(dict As Object) As Object()
    Dim arr() As Object
    Dim key As Variant, count As Long
    ReDim arr(dict.count - 1)
    For Each key In dict.keys
        Set arr(count) = dict(key)
        count = count + 1
    Next key
End Function

Public Sub AddItemsToDict(dict As Object, key As Variant, val As Variant)
    Dim newArr As Variant
    Dim i As Long
    
    If dict.Exists(key) Then
        dict(key) = AddItemsToArr(dict(key), val)
    Else
        If Not IsObject(val) Then
            dict(key) = val
        Else
            Set dict(key) = val
        End If
    End If
End Sub

Public Function CreateDictFromObjs(objs As Variant, prop As String) As Object
    Dim i As Long, dict As Object
    Dim key As Variant
    Set dict = CreateObject("Scripting.Dictionary")
    
    For i = LBound(objs) To UBound(objs)
        key = CallByName(objs(i), prop, VbGet)
        AddItemsToDict dict, key, objs(i)
    Next i
    
    Set CreateDictFromObjs = dict
End Function

'************************************************************************************************************
'*************************************CONVERSION OF DATA TYPE************************************************
'************************************************************************************************************
Public Function ArrToStr(var As Variant) As String
    Dim i As Long, myStr As String
    myStr = CStr(LBound(var))
    For i = LBound(var) + 1 To UBound(var)
        myStr = myStr & "," & var(i)
    Next i
End Function

Public Function StrToArrDbl(myStr As String) As Double()
    Dim str_splitted() As String
    str_splitted = Split(myStr, ",")
    CDbl_arr (str_splitted)
End Function

Public Function StrToArrLng(myStr As String) As Double()
    Dim str_splitted() As String
    str_splitted = Split(myStr, ",")
    CLng_arr (str_splitted)
End Function

Public Function CObj_arr(var As Variant) As Variant
    Dim arr() As Object, i As Long
    ReDim arr(LBound(var) To UBound(var))
    For i = LBound(var) To UBound(var)
        Set arr(i) = var(i)
    Next i
    CObj_arr = arr
End Function

Public Function CDbl_arr(var As Variant) As Double()
    Dim arr() As Double, i As Long
    ReDim arr(LBound(var) To UBound(var))
    For i = LBound(var) To UBound(var)
        arr(i) = CDbl(var(i))
    Next i
    CDbl_arr = arr
End Function

Public Function CStr_arr(var As Variant) As String()
    Dim arr() As String, i As Long
    If IsArray(var) Then
        ReDim arr(LBound(var) To UBound(var))
        For i = LBound(var) To UBound(var)
            arr(i) = CStr(var(i))
        Next i
    Else
        ReDim arr(0)
        arr(0) = CStr(var)
    End If
    CStr_arr = arr
End Function

Public Function CLng_arr(var As Variant) As Long()
    Dim arr() As Long, i As Long
    ReDim arr(LBound(var) To UBound(var))
    For i = LBound(var) To UBound(var)
        arr(i) = CLng(var(i))
    Next i
    CLng_arr = arr
End Function
Public Function CBoolArr(var As Variant) As Boolean()
    Dim arr() As Boolean, i As Long
    ReDim arr(LBound(var) To UBound(var))
    For i = LBound(var) To UBound(var)
        arr(i) = CBool(var(i))
    Next i
    CBoolArr = arr
End Function
'************************************************************************************************************
'*************************************Boolean Array to String************************************************
'***************************************************************************************************
Public Function BooleanArrToString(bArr() As Boolean) As String
'Convert Boolean Array to a string, which 0 for false, 1 for true
    Dim i As Long, sTemp As String
    Dim sRet As String
    
    For i = LBound(bArr) To UBound(bArr)
        If bArr(i) Then
            sTemp = "1"
        Else
            sTemp = "0"
        End If
        sRet = sRet & sTemp
    Next i
    BooleanArrToString = sRet
End Function

Public Function StringToBooleanArr(str As String) As Boolean()
    Dim lenStr As Integer
    lenStr = Len(str)
    
    Dim boolArr() As Boolean
    ReDim boolArr(1 To lenStr)
    
    Dim i As Integer
    For i = 1 To lenStr
        If Mid(str, i, 1) = "0" Then
            boolArr(i) = False
        ElseIf Mid(str, i, 1) = "1" Then
            boolArr(i) = True
        Else
            'Invalid character, return empty array
            ReDim boolArr(0)
            Exit Function
        End If
    Next i
    
    StringToBooleanArr = ChangeArrBase(boolArr, 0)
    
End Function

Public Function CombineArrays(arr1 As Variant, arr2 As Variant) As Variant
    Dim resultArray() As Variant
    Dim i As Long, j As Long
    Dim size1 As Long, size2 As Long

    ' Check if arr1 and arr2 are initialized arrays
    If isInitialised(arr1) And Not isInitialised(arr2) Then
        CombineArrays = arr1
        Exit Function
    ElseIf isInitialised(arr1) And Not isInitialised(arr2) Then
        CombineArrays = arr2
        Exit Function
    ElseIf Not isInitialised(arr1) And Not isInitialised(arr2) Then
        Exit Function
    End If

    ' Get the size of each array
    size1 = UBound(arr1) - LBound(arr1) + 1
    size2 = UBound(arr2) - LBound(arr2) + 1

    ' Resize the result array to fit elements of both arrays
    ReDim resultArray(0 To size1 + size2 - 1)

    ' Copy elements from arr1 to resultArray
    For i = LBound(arr1) To UBound(arr1)
        resultArray(i - LBound(arr1)) = arr1(i)
    Next i

    ' Copy elements from arr2 to resultArray
    For j = LBound(arr2) To UBound(arr2)
        resultArray(size1 + j - LBound(arr2)) = arr2(j)
    Next j

    CombineArrays = resultArray
End Function

