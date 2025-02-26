VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWorksheetsInteraction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Library"

'********************************************************
'This class contains some useful functions in macro development.
'@Folder("General")
'********************************************************
'Arthor: Lucas LEUNG
'Update Log
'30 May 2023 - Minor Updates
'06 Apr 2023 - Updated 'setworksheet' and 'Finds' function with better Err Handling Method. Add Check Input Type Function.
'12 Mar 2023 - Finds become excat match
'1 Feb 2023 - minor update
'21 Dec 2022 - update 'getLocVar' Function
'11 Dec 2022 - Update 'setWorksheet' function
'09 Dec 2022 - new function "Finds"
'25 Oct 2022 - Add Functions (FindFirstNonEmptyCell, CountNonEmptyCell, GetNonEmptyCell)
'31 Jul 2022 - Add Functions to change varaint array to double/string array

'*******************************************************

Private genFunc As clsGeneralFunctions

Public Sub Class_Initialize()
    Set genFunc = New clsGeneralFunctions
End Sub

Public Function RowNumInputToLong(val As Variant, Optional errText As String = vbNullString) As Long
    Const sProcName As String = "ValidateRowNumInput"
    On Error GoTo Err:
    Dim rng As Range
    Set rng = rows(val)
    'Debug.Print rng.row
    'x = CLng(val)
    RowNumInputToLong = CLng(val)
    Exit Function
Err:
    If errText = vbNullString Then errText = "Row Number Input Format is Wrong!! Please Input Integer"
    Err.Raise 5002, sProcName, errText
    ValidateRowNumInput = False
End Function

Public Function ColNumInputToLong(val As Variant, Optional errText As String = vbNullString) As Long
    'Check if the input is a valid column input, either letter or integer
    On Error GoTo TryLetter:
    Dim rng As Range
    Set rng = Columns(val)
    Debug.Print rng.column
    ColNumInputToLong = rng.column
    Exit Function
    
TryLetter:
    On Error GoTo -1
    On Error GoTo Err:
    ColNumInputToLong = Range(val & 1).column
    Exit Function
Err:
    If errText = vbNullString Then errText = "Column Number Input Format is Wrong!! Please Input Integer or Column Letter"
    Err.Raise 5002, sProcName, errText
End Function


Public Function Letter2Number(letter As String) As Long
'PURPOSE: Convert a given letter into it's corresponding Numeric Reference
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

    Dim columnNumber As Long
    columnNumber = Range(letter & 1).column
    Letter2Number = columnNumber
End Function
Public Function Number2Letter(num As Long) As String
'PURPOSE: Convert a given number into it's corresponding Letter Reference
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

    Dim ColumnLetter As String

    'Convert To Column Letter
    ColumnLetter = Split(Cells(1, num).address, "$")(1)
    Number2Letter = ColumnLetter
End Function


Public Function Finds(searchRng As Range, searchArr As Variant, Optional isMustFind As Boolean = True, Optional errText As String = vbNullString) As Range()
    'find an array of text in a specified range
    Dim rng() As Range
    Dim i As Long
    
    ReDim rng(LBound(searchArr) To UBound(searchArr))
    For i = LBound(searchArr) To UBound(searchArr)
        Set rng(i) = searchRng.Find(searchArr(i), LookAt:=xlWhole, MatchCase:=True)
        If rng(i) Is Nothing And isMustFind Then
            If errText = vbNullString Then errText = "The tag '" & searchArr(i) & "' cannot be found in worksheet '" & searchRng.Worksheet.Name & "'"
            Err.Raise 5001, "Finds", errText
            Exit Function
        End If
    Next i
    
    Finds = rng
    
End Function

Function FindAll(searchRng As Range, searchText As String, Optional ws As Worksheet) As Collection
    'to find all cells that contain the specified text in specified range
    Dim c As Range
    Dim coll As New Collection
    Dim firstAddress As String
    If ws Is Nothing Then Set ws = ActiveSheet
    
    With searchRng
        Set c = .Find(searchText, LookIn:=xlValues)
        If Not c Is Nothing Then
            firstAddress = c.address
            Do
                Debug.Print c.address
                coll.Add c
                Set c = .FindNext(c)
            Loop While Not c.address = firstAddress
        End If
    End With
    Set FindAll = coll
End Function

Public Function getLocVar(ws As Worksheet, label As String, isRow As Boolean, Optional errText As String = vbNullString, _
                            Optional isMustMatch = True, Optional rowOrColNum As Long = 1) As Long
    'When isMustMatch = false, return -1 if no match
    'When isMustMatch = true, pop up err message and terminate the program if no match
    getLocVar = -1
    If isMustMatch Then
        On Error GoTo Err_noMatch
    Else
        On Error Resume Next
    End If
    
    With ws
        If isRow Then
            getLocVar = Application.match(label, .rows(rowOrColNum), 0)
            
        Else
            getLocVar = Application.match(label, .Columns(rowOrColNum), 0)
        End If
    End With
    
    On Error GoTo 0
    
    Exit Function
    
     
Err_noMatch:
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Dim errMsg As String

    If isRow Then
        If errText = vbNullString Then
            errMsg = "Cannot locate the column containing the correct data!!! :(" & vbNewLine & _
                    "Please input '" & label & "' at the column containing the correct data in Row 1 of the '" & ws.Name & "' worksheet!!!!"
        Else
            errMsg = errText
        End If
        Err.Raise 10001, "clsWorksheetInteractuib.getLocVar", errMsg

    Else:
        If errText = vbNullString Then
            errMsg = "Cannot locate the row containing the correct data!!! :(" & vbNewLine & _
                    "Please input '" & label & "' at the row containing the correct data in Column 1 of the '" & ws.Name & "' worksheet!!!!"
        Else
            errMsg = errText
        End If
        Err.Raise 10001, Description:=errMsg
        'MsgBox "Err: Cannot locate the row containing the " & errText & "!!! :(" & vbNewLine & _
                "Please input '" & label & "' at the row containing the" & errText & " in Column 1 of the '" & ws.name & "' worksheet!!!!"
    End If

    End
End Function

Public Function getLocVar_containText(ws As Worksheet, label As String, isSearchRow As Boolean) As Long()
    Dim lCol As Long, searchRng As Range, lRow As Long
    Dim rng As Range, i As Long
    Dim arrL_matched As Object, arr_matched() As Long
    '1. set the range by finding the first and last row
    Set arrL_matched = CreateObject("System.Collections.ArrayList")
    If isSearchRow Then
        lCol = FindLastColOfWS(1, ws)
        Set searchRng = Range(ws.Cells(1, 1), ws.Cells(1, lCol))
    Else
        lRow = FindLastRowOfWS(1, ws)
        Set searchRng = Range(ws.Cells(1, 1), ws.Cells(lRow, 1))
    End If
    
    For Each rng In searchRng
        If InStr(1, rng.Text, label) > 0 Then
            If isSearchRow Then
                arrL_matched.Add rng.column
            Else
                arrL_matched.Add rng.row
            End If
        End If
    Next

    ReDim arr_matched(arrL_matched.count - 1)
    For i = 0 To arrL_matched.count - 1
        arr_matched(i) = arrL_matched(i)
    Next i
    getLocVar_containText = arr_matched
End Function


Public Function setWorksheet(wsName As String, Optional myWB As Workbook, Optional errText As String = vbNullString, Optional isSkipErr As Boolean = False) As Worksheet
    
    On Error GoTo Err_cantFindWS
    If myWB Is Nothing Then Set myWB = ActiveWorkbook
    Set setWorksheet = myWB.Worksheets(wsName)
    
    On Error GoTo 0
    Exit Function
    
Err_cantFindWS:
    If Not isSkipErr Then
        Application.ScreenUpdating = True
        Application.Calculation = xlAutomatic
        If errText = vbNullString Then errText = "Worksheet '" & wsName & "' cannot be found."
        Err.Raise 5000, "SetWorksheet", "Worksheet '" & wsName & "' cannot be found."
        'MsgBox "Err: Cannot find the Worksheet '" & wsName & "' " & errText & " :(" '& vbNewLine & _
               '"The worksheet shall be name as '" & wsName & "' and follow the required format. "
        Exit Function
    Else
        Set setWorksheet = Nothing
    End If
End Function
Public Function FindFirstNonEmptyCell(rng As Range) As Range
    On Error GoTo ErrHandler
    
    ' Check if the first cell in the range is non-empty
    If Not IsEmpty(rng.Cells(1, 1).value) Then
        Set FindFirstNonEmptyCell = rng.Cells(1, 1)
        Exit Function
    End If
    
    ' Use Find method to locate the first non-empty cell
    Dim foundCell As Range
    Set foundCell = rng.Find(What:="*", LookIn:=xlValues, LookAt:=xlPart, _
                             SearchOrder:=xlByRows, SearchDirection:=xlNext)
    
    If Not foundCell Is Nothing Then
        Set FindFirstNonEmptyCell = foundCell
    Else
        Set FindFirstNonEmptyCell = Nothing
    End If

    Exit Function

ErrHandler:
    ' Handle errors (e.g., if no non-empty cell is found)
    Set FindFirstNonEmptyCell = Nothing
End Function


Public Function FindLastNonEmptyCell(rng As Range) As Range
    Set FindLastNonEmptyCell = rng.Find(What:="*", LookIn:=xlValues, SearchDirection:=xlPrevious)
End Function

Public Function FindNextNonEmptyCell(startRng As Range, Optional lRow As Long = 1048576) As Range
    Dim searchRng As Range
    Dim fRow As Long, rCol As Long, ws As Worksheet
    Set ws = startRng.Worksheet
    fRow = startRng.row
    rCol = startRng.column
    If lRow < fRow Then lRow = fRow
    With ws
        Set searchRng = .Range(.Cells(fRow, rCol), .Cells(lRow, rCol))
        Set FindNextNonEmptyCell = searchRng.Find(What:="*", LookIn:=xlValues)
    End With
End Function

Public Function FindNextEmptyCell(startRng As Range, Optional lRow As Long = 1048576) As Range
   
    Dim searchRng As Range
    Dim fRow As Long, rCol As Long, ws As Worksheet
    Set ws = startRng.Worksheet
    fRow = startRng.row
    rCol = startRng.column
    If lRow < fRow Then lRow = fRow
    With ws
        Set searchRng = .Range(.Cells(fRow, rCol), .Cells(lRow, rCol))
        Set FindNextEmptyCell = searchRng.Find(What:=vbNullString, LookIn:=xlValues)
    End With
End Function

Public Function CountNonEmptyCell(rng As Range) As Long
    Dim cell As Range
    Dim count As Long
    Dim uniqueCells As Collection
    Set uniqueCells = New Collection
    
    On Error Resume Next
    
    For Each cell In rng.Cells
        If cell.MergeArea.Cells(1, 1).address = cell.address Then
            ' Only consider the first cell in the merge area
            If cell.MergeArea.Cells(1, 1).HasFormula Or cell.MergeArea.Cells(1, 1).value <> "" Then
                uniqueCells.Add cell, cell.address
            End If
        End If
    Next cell
    
    On Error GoTo 0
    CountNonEmptyCell = uniqueCells.count
End Function
Public Function GetNonEmptyCellVal(rng As Range) As Variant
    Dim arr As Variant, count As Long
    Dim rng_temp() As Range
    rng_temp = FindNonEmptyCell(rng)
    ReDim arr(UBound(rng_temp))
    count = 0
    For Each cell In rng_temp
        arr(count) = cell.value
        count = count + 1
    Next
    GetNonEmptyCellVal = arr
End Function
Public Function FindNonEmptyCell(rng As Range) As Range()
    Dim rng_nonEmpty() As Range, cell As Range
    Dim num_nonEmpty As Long, count As Long, i As Long
    Dim fRng As Range, lRng As Range, searchRng As Range
    
    Set fRng = FindFirstNonEmptyCell(rng)
    Set lRng = FindLastNonEmptyCell(rng)
    Set searchRng = Range(fRng, lRng)
    num_nonEmpty = CountNonEmptyCell(rng)
    ReDim rng_nonEmpty(num_nonEmpty - 1)
    
    For Each cell In searchRng
        If Not cell.Text = vbNullString Then
            Set rng_nonEmpty(count) = cell
            count = count + 1
        End If
        If count = num_nonEmpty Then Exit For
    Next
    FindNonEmptyCell = rng_nonEmpty
End Function

Public Function FindLastRowOfWS(Optional rCol As Long = 1, Optional ws As Worksheet) As Long
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    FindLastRowOfWS = ws.Cells(ws.rows.count, rCol).End(xlUp).row
    'FindLastRowOfWS = .Range("A" & .Rows.count).End(xlUp).row
    '
End Function
Public Function FindLastRow(ByVal ref_Row As Long, ByVal refCol As Long, Optional ws As Worksheet) As Long 'return the last row of a tables
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    With ws
        If .Cells(ref_Row, refCol).Text = vbNullString Then
            FindLastRow = ref_Row - 1
        'MsgBox "NO DATA IN THE WS"
        ElseIf .Cells(ref_Row + 1, refCol).Text <> vbNullString Then
            FindLastRow = .Cells(ref_Row, refCol).End(xlDown).row
        Else
            FindLastRow = ref_Row
        End If
    End With

End Function

Public Function FindLastRowOfSameGrp(ByVal rRow As Long, ByVal rCol As Long, Optional ws As Worksheet) As Long
    Dim cRow As Long
    Dim lRow As Long
    
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    cRow = rRow
    lRow = rRow + 1
'find lRow_SchSameGrp
    With ws
        Do Until .Cells(cRow, rCol).Text <> .Cells(lRow, rCol).Text
            cRow = lRow
            lRow = lRow + 1
        Loop
        FindLastRowOfSameGrp = cRow
    End With
End Function

Public Function FindLastCol(ByVal ref_Row As Long, ByVal refCol As Long, Optional ws As Worksheet) As Long
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    With ws
        If .Cells(ref_Row, refCol).Text = vbNullString Then
            FindLastCol = refCol
'            MsgBox "NO DATA IN THE WS"
        ElseIf .Cells(ref_Row, refCol + 1) <> vbNullString Then
            FindLastCol = .Cells(ref_Row, refCol).End(xlToRight).column

        Else
            FindLastCol = refCol
        End If
    End With
End Function

Public Function FindLastColOfWS(Optional rRow As Long = 1, Optional ws As Worksheet) As Long
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    FindLastColOfWS = ws.Cells(rRow, ws.Columns.count).End(xlToLeft).column
    'FindLastRowOfWS = .Range("A" & .Rows.count).End(xlUp).row
    '
End Function

Public Function SetArrayOneDim(fRow As Long, fCol As Long, Optional isByRow As Boolean = True, Optional arrSize As Long = 0, Optional ws As Worksheet) As Variant
    
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    Dim i As Long
    Dim lRow As Long, lCol As Long
    Dim var As Variant
    
    With ws
        If arrSize = 0 Then
            If isByRow Then
                lRow = FindLastRow(fRow, fCol, ws)
                If lRow < fRow Then Exit Function
                arrSize = lRow - fRow + 1
            Else
                lCol = FindLastCol(fRow, fCol, ws)
                arrSize = lCol - fCol + 1
            End If
        End If
        
'        If arrSize > 1 Then
'            ReDim var(arrSize - 1)
'
'        Else
'
'        End If
        
        If isByRow Then
            var = .Range(.Cells(fRow, fCol), .Cells(fRow + arrSize - 1, fCol))
        Else
            var = .Range(.Cells(fRow, fCol), .Cells(fRow, fCol + arrSize - 1))
        End If
        
        If arrSize > 1 Then var = genFunc.Condense2DArrTo1D(var)
    End With
    SetArrayOneDim = var
End Function




Public Function SetArrayTwoDim(fRow As Long, fCol As Long, Optional isGrpByRow As Boolean = True, Optional arrSizeY As Long = 0, Optional arrSizeX As Long = 0, Optional ws As Worksheet) As Variant
    
    Dim rowNum As Long, colNum As Long
    Dim lRow As Long, lCol As Long
    Dim size1 As Long, size2 As Long
    
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    lRow = FindLastRow(fRow, fCol, ws)
    If lRow < fRow Then Exit Function
    
    lCol = FindLastCol(fRow, fCol, ws)
    If lCol < fCol Then Exit Function
    
    If arrSizeY = 0 Then
        rowNum = lRow - fRow + 1
    Else
        rowNum = arrSizeY
    End If
    
    If arrSizeX = 0 Then
        colNum = lCol - fCol + 1
    Else
        colNum = arrSizeX
    End If
    
    SetArrayTwoDim = ws.Range(ws.Cells(fRow, fCol), ws.Cells(fRow + rowNum - 1, fCol + colNum - 1)).value
    
    If Not isGrpByRow Then
        SetArrayTwoDim = TransposeArray(SetArrayTwoDim)
    End If
End Function

Public Function TransposeArray(MyArray As Variant) As Variant
Dim x As Long
Dim y As Long
Dim Xupper As Long
Dim Yupper As Long
Dim Xlower As Long
Dim Ylower As Long

Dim tempArray As Variant
    Xupper = UBound(MyArray, 2)
    Yupper = UBound(MyArray, 1)
    Xlower = LBound(MyArray, 2)
    Ylower = LBound(MyArray, 1)
    
    ReDim tempArray(Xlower To Xupper, Ylower To Yupper)
    For x = Xlower To Xupper
        For y = Ylower To Yupper
            tempArray(x, y) = MyArray(y, x)
        Next y
    Next x
    TransposeArray = tempArray
End Function

Public Sub WriteArrToRngTwoDim(myArr As Variant, fRow As Long, fCol As Long, Optional ws As Worksheet)
    Dim arrSizeX As Long, arrSizeY As Long
    Dim myRng As Range
    
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    arrSizeY = UBound(myArr, 1) - LBound(myArr, 1) + 1
    arrSizeX = UBound(myArr, 2) - LBound(myArr, 2) + 1
    Set myRng = ws.Range(ws.Cells(fRow, fCol), ws.Cells(fRow + arrSizeY - 1, fCol + arrSizeX - 1))
    myRng.value = myArr
End Sub

Public Sub WriteArrToColumn(myArr As Variant, fRow As Long, fCol As Long, Optional ws As Worksheet, Optional size As Long)
    Dim arrSizeY As Long
    Dim myRng As Range
    Dim arr_2D As Variant
    Dim i As Long
    
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    If size > 0 Then
        arrSizeY = size
    Else
        arrSizeY = UBound(myArr) - LBound(myArr) + 1
    End If
    
    ReDim arr_2D(LBound(myArr) To UBound(myArr), 0)
    For i = LBound(myArr) To UBound(myArr)
        arr_2D(i, 0) = myArr(i)
    Next i
    
    Set myRng = ws.Range(ws.Cells(fRow, fCol), ws.Cells(fRow + arrSizeY - 1, fCol))
    If genFunc.IsString(myArr(LBound(myArr))) Then myRng.NumberFormat = "@"
    myRng.value = arr_2D
End Sub

Public Sub WriteArrToRow(myArr As Variant, fRow As Long, fCol As Long, Optional ws As Worksheet, Optional size As Long)
    Dim arrSizeX As Long
    Dim myRng As Range
    Dim arr_2D As Variant
    Dim i As Long
    
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    If genFunc.isInitialised(myArr) Then
        If size > 0 Then
            arrSizeX = size
        Else
            arrSizeX = UBound(myArr) - LBound(myArr) + 1
        End If
        
        ReDim arr_2D(0, LBound(myArr) To UBound(myArr))
        For i = LBound(myArr) To UBound(myArr)
            arr_2D(0, i) = myArr(i)
        Next i
        
        Set myRng = ws.Range(ws.Cells(fRow, fCol), ws.Cells(fRow, fCol + arrSizeX - 1))
        myRng.value = arr_2D
    End If
End Sub

Public Function SetWorkbook(filePath As String) As Workbook
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    wbName = fso.GetFileName(filePath)
    
    On Error GoTo WBNotOpenYet
    Set SetWorkbook = Workbooks(wbName)
    Exit Function
    
WBNotOpenYet:
    Set SetWorkbook = Workbooks.Open(filePath, ReadOnly:=True)
    
End Function

Public Function LastNonEmptyColumn(ws As Worksheet, rowNum As Long) As Long
    Dim lastCol As Long

    ' Check if the specified row is valid
    If rowNum < 1 Then
        LastNonEmptyColumn = -1 ' Return -1 for invalid row
        Exit Function
    End If

    ' Find the last non-empty column in the specified row
    lastCol = ws.Cells(rowNum, ws.Columns.count).End(xlToLeft).column

    ' Check if the last column found is empty
    If IsEmpty(ws.Cells(rowNum, lastCol)) Then
        LastNonEmptyColumn = -1 ' Return -1 if the last column is empty
    Else
        LastNonEmptyColumn = lastCol
    End If

End Function

'Function IsCellAddress(str As String) As Boolean
'    Dim regex As Object
'    Dim match As Object
'    Dim columnPart As String
'    Dim rowPart As String
'    Dim columnNumber As Long
'    Dim rowNumber As Long
'
'    Dim var As Variant, CellAddress As String
'    var = Split(str, "!")
'    CellAddress = var(UBound(var))
'    ' Create the regex object
'    Set regex = CreateObject("VBScript.RegExp")
'
'    ' Regular expression to match full cell addresses
'    ' Example: '[Workbook.xlsx]Sheet'!$A$1 or just A1
'    regex.Pattern = "(\$?[A-Z]+\$?\d+)"
'    regex.IgnoreCase = True
'    regex.Global = False
'
'    ' Check if the cell address matches the pattern
'    If Not regex.TEST(CellAddress) Then
'        IsCellAddress = False
'        Exit Function
'    End If
'
'    ' Extract the actual cell reference (e.g., A1 from $A$1)
'    regex.Pattern = "\$?([A-Z]+)\$?(\d+)"
'    Set match = regex.Execute(CellAddress)
'
'    If match.count = 0 Then
'        IsCellAddress = False
'        Exit Function
'    End If
'
'    columnPart = match(0).SubMatches(0)
'    rowPart = match(0).SubMatches(1)
'
'    ' Convert column part to a column number
'    columnNumber = 0
'    Dim i As Long
'    For i = 1 To Len(columnPart)
'        columnNumber = columnNumber * 26 + (Asc(UCase(Mid(columnPart, i, 1))) - Asc("A") + 1)
'    Next i
'
'    ' Convert row part to a row number
'    rowNumber = CLng(rowPart)
'
'    ' Validate the column and row numbers against Excel's limits
'    If columnNumber >= 1 And columnNumber <= 16384 And rowNumber >= 1 And rowNumber <= 1048576 Then
'        IsCellAddress = True
'    Else
'        IsCellAddress = False
'    End If
'End Function
Function IsCellAddress(str As String) As Boolean
    On Error Resume Next
    IsCellAddress = Not IsError(Range(str).address)
    On Error GoTo 0
End Function

Function GetAddressWorksheetName(str As String) As String
    Dim regex As Object
    Dim match As Object
    Dim lastExclaimation As Integer, lastBracket As Integer
    Dim wsName As String
    
    If Not IsCellAddress(str) Then Exit Function

    lastExclaimation = FindLastCharInStr(str, "!")
    If lastExclaimation = 0 Then Exit Function 'no worksheet name
    
    lastBracket = FindLastCharInStr(str, "]")
    
    wsName = Mid(str, lastBracket + 1, lastExclaimation - lastBracket - 1)
    
    If left(wsName, 1) = "'" Then
        wsName = Mid(wsName, 2, Len(wsName) - 1)
    End If
    
    If Right(wsName, 1) = "'" Then
        wsName = Mid(wsName, 1, Len(wsName) - 1)
    End If
    
    GetAddressWorksheetName = wsName
End Function

Function FindLastCharInStr(str As String, findChar As String) As Long
    Dim i As Long

    ' Initialize the function to return 0 if no '!' is found
    FindLastExclamation = 0

    ' Loop through the string backwards
    For i = Len(str) To 1 Step -1
        If Mid(str, i, 1) = findChar Then
            FindLastCharInStr = i
            Exit Function
        End If
    Next i
End Function
