VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "oDataSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("DataSheet.ModelData")

'********************************************************
'This is the class object representing Data Sheet Containing Model Data
'Arthor: Lucas LEUNG
'Update Log
'28 Aug 2023 - Initial Setup
'*******************************************************
Option Explicit
Private wsInteract As clsWorksheetsInteraction

Private mWsName As String, mWS As Worksheet

'These are the default properties/tags that will be used in the system
Private iRow As Long, sRow As String 'the reference row representing the heading of the table
Private fCol As Long, lCol As Long, width As Long
Private oLastRow As Long 'the last row at the beginning of certain procedure
Private oNumOfDataBlocks As Long
Private mCollProps As New Collection
Private mTagsSelector As Object

Private Sub Class_Initialize()
    Set wsInteract = New clsWorksheetsInteraction
    sRow = "rRow"
    
End Sub

Public Sub Initialize(wsName As String, tagSelector As Object, ParamArray tags() As Variant)
    Dim ret As Integer
    
    mWsName = wsName
    
    
    Set mTagsSelector = tagSelector
    Set mWS = wsInteract.setWorksheet(wsName, isSkipErr:=True) 'if fail to set the worksheet, error msg will be popped up.
    If mWS Is Nothing Then
        MsgBox "Cannot Find System Worksheet ' " & wsName & "'! The Macro will be terminated."
    End If
    iRow = wsInteract.getLocVar(mWS, sRow, False)
    
    
    Dim i As Long
    For i = LBound(tags) To UBound(tags)
        CreateProp CStr(tags(i))
        If ret = -1 Then GoTo ExitSub
    Next i

'    Dim rCol As Long
'    rCol = wsInteract.getLocVar(mWS, CStr(tags(0)), True, isMustMatch:=False)

'    lRow = wsInteract.FindLastRowOfWS(rCol, mWS)
'    If lRow <= iRow Then lRow = iRow
    fCol = wsInteract.FindFirstNonEmptyCell(mWS.rows(1)).column
    lCol = wsInteract.FindLastNonEmptyCell(mWS.rows(1)).column
    width = lCol - fCol + 2
    oLastRow = Me.lastRow(True) + 1
    oNumOfDataBlocks = Me.NumOfDataBlocks
    Exit Sub

ExitSub:
    'Err occur if the worksheet/ any tag cannot be found
    g_log.HandleError Err.Source, CStr(Err.Number), Err.Description
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    End
End Sub

Public Function GetDataframe(ParamArray propName() As Variant) As clsDataFrame
    Dim df As clsDataFrame, heads() As String, i As Long
    Set df = New clsDataFrame
    ReDim heads(0 To UBound(propName))
    For i = 0 To UBound(propName)
        heads(i) = CStr(propName(i))
    Next i
    df.Init_ReadWorksheet2 mWS, rRow:=iRow, headTags:=heads
    Set GetDataframe = df
End Function

Public Sub WriteDataframe(df As clsDataFrame, isFirstRow As Boolean, _
                            isShiftColWhenRowLimitExceed As Boolean, ParamArray propName() As Variant)
    'the df shall be in the same order of the 'true' column
    Dim count As Long, i As Long, col As Long
    count = 1
    
    For i = LBound(propName) To UBound(propName)
        
        WriteColumn df.iColumn(count), isFirstRow, CStr(propName(i)), isShiftColWhenRowLimitExceed
'        col = mCollProps(propName(i)).Loc
'        If Not col = -1 Then wsInteract.WriteArrToColumn df.iColumn(count), rowNum, col, mWS
        count = count + 1
    Next i
End Sub

Public Sub WriteColumn(arr As Variant, isFirstRow As Boolean, propName As String, _
                        Optional isShiftColWhenRowLimitExceed As Boolean = False)
    'input shall be 1D array, the array will be written on the first field that set as 'true'
    Dim col As Long, rowNum As Long
    
    If isFirstRow And Not isShiftColWhenRowLimitExceed Then
        rowNum = iRow + 1
        col = mCollProps(propName).Loc
        If Not col = -1 Then wsInteract.WriteArrToColumn arr, rowNum, col, mWS
    ElseIf Not isShiftColWhenRowLimitExceed Then
        rowNum = oLastRow
        col = mCollProps(propName).Loc
        If Not col = -1 Then wsInteract.WriteArrToColumn arr, rowNum, col, mWS
    ElseIf isShiftColWhenRowLimitExceed Then
        If isFirstRow Then
            rowNum = iRow + 1
            col = mCollProps(propName).Loc
        Else
            rowNum = oLastRow
            col = mCollProps(propName).Loc + width * (oNumOfDataBlocks - 1)
        End If
        
        Dim arrSize As Long
        Const rowLimit As Long = 1048576
        arrSize = UBound(arr) - LBound(arr) + 1
        If (rowNum + arrSize) < rowLimit Then
            If Not col = -1 Then wsInteract.WriteArrToColumn arr, rowNum, col, mWS
            GoTo ExitSub
        End If
        
        Dim tempArr As Variant, tempArrSize As Long
        Dim startIndex As Long
        tempArr = SplitArr(arr, LBound(arr), LBound(arr) + rowLimit - rowNum)
        wsInteract.WriteArrToColumn tempArr, rowNum, col, mWS
        startIndex = LBound(arr)
        Do Until arrSize <= 0
            tempArrSize = rowLimit - rowNum + 1
            If arrSize < tempArrSize Then tempArrSize = arrSize
            arrSize = arrSize - tempArrSize
            tempArr = SplitArr(arr, startIndex, tempArrSize - 1 + startIndex)
            wsInteract.WriteArrToColumn tempArr, rowNum, col, mWS
            startIndex = tempArrSize + startIndex
            rowNum = iRow + 1
            col = col + width
        Loop
        
    End If
    
ExitSub:
End Sub

Private Function SplitArr(arr As Variant, startIndex As Long, endIndex As Long) As Variant
    Dim var As Variant, i As Long
    ReDim var(0 To endIndex - startIndex)
    For i = 0 To endIndex - startIndex
        var(i) = arr(i + startIndex)
    Next i
    SplitArr = var
End Function

Public Sub FillColumns(ParamArray propNames() As Variant)
'for some of the ModelDataSheets Only
    Dim rng1 As Range, rng2 As Range
    Dim rCol As Long, i As Long
    
    For i = LBound(propNames) To UBound(propNames)
        FillColumn CStr(propNames(i))
    Next i
End Sub

Private Sub FillColumn(propName As String)
'for some of the ModelDataSheets Only
    Dim rng1 As Range, rng2 As Range
    Dim startRow As Long, rCol As Long, i As Long
    Dim lRow As Long
    startRow = iRow + 1
    For i = 1 To Me.NumOfDataBlocks
        rCol = mCollProps(propName).Loc + width * (i - 1)
        If i > 1 Then
            mWS.Cells(iRow + 1, mCollProps(propName).Loc).Copy mWS.Cells(startRow, rCol)
        End If

        With mWS
            lRow = wsInteract.FindLastRow(startRow, mCollProps(1).Loc + width * (i - 1), mWS)
            Set rng1 = .Cells(startRow, rCol)
            Set rng2 = Range(.Cells(startRow, rCol), .Cells(lRow, rCol))
            rng1.AutoFill Destination:=rng2
            Application.Calculate
            .Range(.Cells(startRow + 1, rCol), .Cells(lRow, rCol)).Copy
            .Cells(startRow + 1, rCol).PasteSpecial Paste:=xlPasteValues
        End With
    Next i
End Sub

Public Function GetTagColumns(ws As Worksheet) As Long()
'Return the columns of all tags
    Dim Loc() As Long, i As Long
    ReDim Loc(0 To mCollProps.count - 1)
    For i = 1 To mCollProps.count
        mCollProps(i).Loc = wsInteract.getLocVar(ws, mCollProps(i).Text, True, isMustMatch:=False)
        Loc(i - 1) = mCollProps(i).Loc
    Next i
    GetTagColumns = Loc
End Function

Public Function GetTagsText(ParamArray propName() As Variant) As String()
'Return the string of all tags
    Dim str() As String, i As Long
    ReDim str(0 To mCollProps.count - 1)
    For i = 1 To mCollProps.count
        str(i - 1) = mCollProps(i).Text
    Next i
    GetTagsText = str
End Function

Public Function GetTagsDesc(ParamArray propName() As Variant) As String()
'Return the description of all tags
    Dim str() As String, i As Long
    ReDim str(0 To mCollProps.count - 1)
    For i = 1 To mCollProps.count
        str(i - 1) = mCollProps(i).Desc
    Next i
    GetTagsDesc = str
End Function

Private Function CreateProp(Text As String, Optional Desc As String) As Integer
    Dim prop As oProp, ret As Integer
    
    Set prop = New oProp
    prop.Text = Text
    prop.Desc = Desc
    prop.Loc = wsInteract.getLocVar(mWS, Text, True, isMustMatch:=False)
    If prop.Loc = -1 Then
        MsgBox "cannot find tag '" & Text & "' in worksheet '" & mWS.Name & "'!. The macro will be terminate"
        ret = -1
    End If
    mCollProps.Add prop, Text
    CreateProp = ret
End Function

Public Sub ClearData(Optional fRow As Long)
    Dim lRow As Long
    If fRow = 0 Then
        fRow = iRow + 1
    End If
    lRow = wsInteract.FindLastRowOfWS(mCollProps(1).Loc, mWS)
    If lRow <= fRow Then lRow = fRow + 1
    
    mWS.rows(fRow & ":" & lRow).Clear
    If Me.NumOfDataBlocks > 1 Then
        Dim lColData As Long
        lColData = wsInteract.LastNonEmptyColumn(mWS, iRow + 1)
        mWS.Range(Cells(iRow + 1, fCol + width), Cells(fRow - 1, lColData)).Clear
    End If
    oLastRow = iRow + 1
    oNumOfDataBlocks = 1
    g_log.WriteLog "Rows " & fRow & " to " & lRow & " of Worksheets '" & mWsName & "'" & "is cleared."
    
    'lRow = iRow
End Sub

Public Sub ClearColumnData(propName As String, Optional fRow As Long = 0)
    Dim lRow As Long
    Dim col As Long, i As Long
    
    If fRow = 0 Then fRow = Me.refRow + 1
    'lRow = Me.lastRow
    lRow = 1048576
    col = mCollProps(propName).Loc
    
    If lRow <= fRow Then lRow = fRow + 1
    
    For i = 1 To oNumOfDataBlocks
        mWS.Range(mWS.Cells(fRow, col + (i - 1) * width), mWS.Cells(lRow, col + (i - 1) * width)).Clear
    Next i
    g_log.WriteLog "Column " & col & " of Worksheets '" & mWsName & "'" & "is cleared."
End Sub

Public Sub ClearColumns(ParamArray propNames() As Variant)
    Dim propName As Variant
    
    For Each propName In propNames
        ClearColumnData CStr(propName)
    Next
    oLastRow = iRow + 1
    oNumOfDataBlocks = 1
End Sub

Public Property Get tagSelector() As Object
    Set tagSelector = mTagsSelector
End Property

Public Property Get refRow() As Long
    refRow = iRow
End Property

Public Property Get lastRow(Optional isShiftColWhenRowLimitExceed As Boolean = False) As Long
    Dim col As Long
    If isShiftColWhenRowLimitExceed Then
        col = mCollProps(1).Loc + width * (Me.NumOfDataBlocks - 1)
    Else
        col = mCollProps(1).Loc
    End If
    lastRow = wsInteract.FindLastRowOfWS(mCollProps(1).Loc, mWS)
End Property

Public Property Get NumOfDataBlocks() As Long
    Dim lColData As Long
    lColData = wsInteract.LastNonEmptyColumn(mWS, iRow + 1)
    If lColData = 0 Then
        NumOfDataBlocks = 1
    Else
        NumOfDataBlocks = Int((lColData - fCol) / width) + 1
    End If
End Property
