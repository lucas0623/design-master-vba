VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SummaryToWS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Operation.SummaryToWS")
'System and Library
Private mCompleteMessageAddition As String
Private mTerminateMessageMain As String, mTerminateMessageAddition As String
Private Const mOperName = "SummaryToWs"

Private genFuncs As New clsGeneralFunctions, wsInteract As New clsWorksheetsInteraction
Private collMethod As New LibCollectionMethod, console As New Debugger
Private DsSys As New DataSheetSystem


'Private colTag As ColumnTag
Private sumWS As Worksheet, columnTags As Object 'a dictionary that contains all column tags data in the summary sheet
Private targetWsTags As Object 'a collection that contains all 'targetWS' Tag in the summary sheet

Private fRow_summary As Long, lRow_summary As Long, rows As Long
Private df_sum As clsDataFrame

Private dataWS As Worksheet
Private designWorksheets As Object

'Userform and Print PDF variables
Private isBoldedOnly As Boolean, isOutputPDF As Boolean, isIncludePrefix As Boolean
Private myFilePath As String
Private strCount As String, countPDF As Long
Private countSucessful As Long, countTotal As Long 'count how many rows of data are sucessfully transferred.

Private isMissingKey As Boolean
Private macroName As String, objName As String, procName As String

Public Property Get CompleteMessageAddition() As String
    CompleteMessageAddition = mCompleteMessageAddition
End Property

Public Property Get TerminateMessageMain() As String
    TerminateMessageMain = mTerminateMessageMain
End Property

Public Property Get TerminateMessageAddition() As String
    TerminateMessageAddition = mTerminateMessageAddition
End Property

Public Function Main(Optional summarySheet As Worksheet) As Integer
    Dim i As Long, ret As Integer
    
    'On Error GoTo ExitFunc:
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    g_log.WriteLog "Operation: 'SummaryToDsWs' Started."
    
    ret = Initialize
    If Not ret = 0 Then GoTo ExitFunc
    
    ret = GetUserInput
    If Not ret = 0 Then
        g_log.WriteLog "Userform canceled. Exit Operation."
        GoTo ExitFunc
    End If
    
    ret = ReadSummaryDataToDF
    If Not ret = 0 Then
        g_log.WriteLog "No Data is Found in the Summary Sheet."
        mTerminateMessageMain = "No Data is Found in the Summary Sheet. Please check if there are column tags in the first row of the Summary Sheet"
        GoTo ExitFunc
    End If
    
    ret = ReadDesignWorksheetIOData
    If Not ret = 0 Then
        g_log.WriteLog "Fail to read WS Input/Output data. Exit Operation."
        GoTo ExitFunc
    End If
    
    ret = TransferDataFromSummary
    If Not ret = 0 Then
    End If
    
    ret = WriteData
    If Not ret = 0 Then
    End If
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    mCompleteMessageAddition = countSucessful & " out of " & countTotal & " number of rows of data are sucessfully transferred."
    Exit Function
ExitFunc:
    Main = -1
    If mTerminateMessageMain = vbNullString Then mTerminateMessageMain = "Unexpected Error in 'SummaryToWS.Main'"
End Function

Private Function Initialize(Optional summarySheet As Worksheet) As Integer
    'On Error GoTo ExitFunc:
    Dim rng As Variant, coll As Collection
    'Dim colTag As ColumnTag
    Dim i As Long
    
    Set DsSys = New DataSheetSystem
    Set columnTags = CreateObject("Scripting.Dictionary")
    Set targetWsTags = CreateObject("Scripting.Dictionary")
    Set designWorksheets = CreateObject("Scripting.Dictionary")
    Set dataWS = wsInteract.setWorksheet("WSData", errText:="that contains the target worksheet input/output data")
    If summarySheet Is Nothing Then Set sumWS = ActiveSheet
    Set coll = wsInteract.FindAll(sumWS.rows(1), "targetWS")
    If coll.count = 0 Then
        mTerminateMessageMain = "Cannot Find any 'targetWS' tag in the first row of the summary sheet. Please check!"
        Initialize = -1
        Exit Function
    End If
    'rng = collMethod.CollToArr1D(coll)
    For Each item In coll
        'colTag = CreateColumnTag(item.column, item.Text)
        'columnTags.Add colTag
        'columnTags.Add CreateColumnTag(tag.column, tag.Text)
    Next
'    ReDim col_tarWSName(LBound(rng) To UBound(rng))
'    For i = LBound(rng) To UBound(rng)
'        col_tarWSName(i) = rng(i).column
'    Next i
'    console.Log col_tarWSName, "col_tarWSName"
    Exit Function
ExitFunc:
    Initialize = -1
    If mTerminateMessageMain = vbNullString Then mTerminateMessageMain = "Unexpected Error in 'SummaryToWS.Initialize'"
End Function

Private Function GetUserInput() As Integer
    Dim ret As Integer
    
    'Create user form for getting user input range
    Dim form As UFBasic
    Dim textBox1 As msForms.Textbox
    Dim CheckBox1 As msForms.checkbox, CheckBox2 As msForms.checkbox, CheckBox3 As msForms.checkbox
    Dim Textbox As msForms.Textbox
    
    'dim rng as Range, dim rng2 as Range
    Set form = New UFBasic
    form.Initialize 300, True
    form.AddRngInputBox textBox1, "Please select the Rows for design", Selection.address
    form.AddCheckBox CheckBox1, "is Output to the Bolded Rows Only?", isCheck:=False
    form.AddCheckBox CheckBox2, "Print PDF?"
    form.AddCheckBox CheckBox3, "Include target worksheet's name as prefix of PDF?", isCheck:=False
    form.AdjustHeight
    LoadUserInput CheckBox1, CheckBox2, CheckBox3
    form.Show
    
    If form.CloseState = 0 Then
        'Get Input
        isBoldedOnly = CheckBox1.value
        isOutputPDF = CheckBox2.value
        isIncludePrefix = CheckBox3.value
        fRow_summary = Range(textBox1.Text).rows(1).row
        lRow_summary = Range(textBox1.Text).rows.count + fRow_summary - 1
        SaveUserInput
    Else
        ret = -1
    End If
    
    GetUserInput = ret
End Function

Private Function SaveUserInput() As Integer
    Dim propGrpName As String
    
    propGrpName = sumWS.name & mOperName
    g_log.WriteLogInDetailMode "Saving User Input..."
    '1.Check if there is existing save. If yes, clear existing data. If no, create a new save
    If DsSys.isPropGrpExist(propGrpName) Then
        'Clear Existing Data
        DsSys.ClearPropGrpVal propGrpName
    Else
        'Create new save
        DsSys.CreateCustomPropGrp propGrpName, "wsName", _
                                                "isBoldedRows", "isPrintPDF", "isIncludePrefix"

    End If

    '2. Write Data
    DsSys.prop(propGrpName, "wsName") = sumWS.name
    'ds_sys.prop(propGrpName, "isBySec") = mbIsBySec
    'ds_sys.prop(propGrpName, "Rows") = msLcSelect
    DsSys.prop(propGrpName, "isBoldedRows") = isBoldedOnly
    DsSys.prop(propGrpName, "isPrintPDF") = isOutputPDF
    DsSys.prop(propGrpName, "isIncludePrefix") = isIncludePrefix
End Function

Private Function LoadUserInput(cbIsBoldedOnly As msForms.checkbox, _
                                cbIsOutputPDF As msForms.checkbox, cbIsIncludePrefix As msForms.checkbox) As Integer
    Dim propGrpName As String
    
    propGrpName = sumWS.name & mOperName
    '1.Check if there is existing save. If yes, try loading the data
    If Not DsSys.isPropGrpExist(propGrpName) Then
        g_log.WriteLogInDetailMode "Cannot find the propGrpName in Data_System Sheet"
        ret = -1
        GoTo ExitFunc
    End If
    
    cbIsBoldedOnly.value = DsSys.prop(propGrpName, "isBoldedRows")
    cbIsOutputPDF.value = DsSys.prop(propGrpName, "isPrintPDF")
    cbIsIncludePrefix.value = DsSys.prop(propGrpName, "isIncludePrefix")
    'tbRngInput.Text = ds_sys.prop(propGrpName, "Rows")
ExitFunc:
    LoadUserInput = ret
End Function

Private Function ReadSummaryDataToDF() As Long
    'Read the data columns by columns.
    Dim heads() As String, rowTags() As String
    Dim data As Variant
    Dim col_heads() As Long, rows() As Long 'an array represent the column number and the row number of the data in the worksheet
    Dim i As Long, ret As Long
    
    'Read worksheet to dataframe
    Set df_sum = New clsDataFrame
    df_sum.Init_ReadWorksheet sumWS, rRow:=fRow_summary - 1, lRow:=lRow_summary
    console.Log df_sum.heads, "df_sum.heads"
    console.Log df_sum.data("1", "1", True), "df_sum.data"
    If df_sum.CountRows = 0 Then GoTo ExitFunc
    'Read Column Tags Data in the summary
    'CreateColumnTag
    'Read targetWS tags data in the summary
    
    'Read WS_Data
    
    'Get the rows number to be considered.
'    For i = fRow_summary To lRow_summary
'        If isBoldedOnly Then
'            'If sumWS.Cells(i, col_tarWSName(0)).Font.Bold = True & sumWS.rows(i).height > 0 Then coll.Add i
'        Else
'            If sumWS.rows(i).height > 0 Then coll.Add i
'        End If
'    Next i
    'If coll.count = 0 Then GoTo ExitFunc:
    'rows = collMethod.CollToArr1D(coll)
    'console.Log (rows)
    
    'get the columns Num of all columns containing tags
    Exit Function
ExitFunc:
    ReadSummaryDataToDF = ret
End Function

Function ReadDesignWorksheetIOData() As Long
    Dim ws As Worksheet
    Dim refRow As Long
    Dim col As Long
    Dim i As Long
    Dim nonEmptyCells() As Range
    Dim designWorksheet As DesignWsIO
    
    
    Set ws = wsInteract.setWorksheet("WSData", isSkipErr:=True)
    If ws Is Nothing Then
        mTerminateMessageMain = "Error: Cannot find the sheet 'WSData' that contains the information of input/output data for all design worksheets."
        GoTo ExitFunc
    End If
    
    ' Find all non-empty cells in the first row
    nonEmptyCells = wsInteract.FindNonEmptyCell(ws.rows(1))
    
    If Not genFuncs.isInitialised(nonEmptyCells) Then
        mTerminateMessageMain = "Error: Cannot find any design worksheet name in the first row of the 'WSData' worksheet."
        GoTo ExitFunc
    End If
    
    ' Find the text 'rRow_tarWSData' in the first row
    refRow = wsInteract.getLocVar(ws, "rRow_tarWSData", False, isMustMatch:=False)
    If refRow = -1 Then
        mTerminateMessageMain = "Error: Cannot find the row tag 'rRow_tarWSData' in the first column of the 'WSData' worksheet."
        GoTo ExitFunc
    End If
    
    
    ' Loop through each non-empty cell found in the first row
    For i = LBound(nonEmptyCells) To UBound(nonEmptyCells)
        If Not collMethod.isKeyExist(designWorksheets, nonEmptyCells(i).value) Then
            Set designWorksheet = ReadSingleDesignWorksheetData(nonEmptyCells(i), refRow)
            designWorksheets.Add designWorksheet, designWorksheet.name
            console.Log designWorksheet.WsInput, designWorksheet.name & " INPUT"
            console.Log designWorksheet.WsInput, designWorksheet.name & " OUTPUT"
        Else
            g_log.RaiseWarning "Duplicate Design Worksheet in 'WSData'. Only the first sheet is considered.", duplicateDesignWorksheet 'Raise Warning
        End If
    Next i
    
    Exit Function
    
ExitFunc:
    ReadDesignWorksheetIOData = -1
End Function

Function ReadSingleDesignWorksheetData(refCell As Range, refRow As Long) As DesignWsIO
    Dim ws As Worksheet
    Dim rCol As Long
    Dim fRow As Long, lRow As Long
    Dim key As String
    Dim value As String
    Dim designObject As DesignWsIO
    Dim ret As String
    
    Set ws = refCell.Parent
    rCol = refCell.column
    
    ' Create a new DesignWsIO object using refCell's value as name
    Set designObject = New DesignWsIO
    designObject.Initialize refCell.value
    
    ' Read and set Input values
    fRow = refRow + 1
    lRow = wsInteract.FindLastRow(refRow, rCol, ws)
    
    For i = fRow To lRow
        key = ws.Cells(i, rCol).value
        value = ws.Cells(i, rCol + 1).value
        ret = designObject.AddInput(key, value)
    Next i
    
    ' Read and set Output values
    For i = refRow To lRow
        key = ws.Cells(i, rCol + 2).value
        value = ws.Cells(i, rCol + 3).value
        ret = designObject.AddOutput(key, value)
    Next i
    
    Set ReadSingleDesignWorksheetData = designObject
End Function
'Private Function ReadDesignWorksheetData() As Long
'    Dim ws As Worksheet 'the worksheet 'WSData' that contains the input/output information
'    designWorksheets
'    If iDataDictofDict.Exists(wsName) Then Exit Function
'
'    'Read the Design Worksheet Setting Data in the 'WS_Data' Sheet
'    Dim col_tarWSData As Long
'    Dim lRow_inputData As Long, lRow_returnData As Long
'    Dim tarWS_iDataLoc As Object 'a dictionary containing the input setting of certain worksheet
'    Dim tarWS_rDataLoc As Object 'a dictionary containing the output setting of certain worksheet
'    Dim key As String
'
'    Set tarWS_iDataLoc = CreateObject("Scripting.Dictionary")
'    Set tarWS_rDataLoc = CreateObject("Scripting.Dictionary")
'
'
'    With dataWS
'        col_tarWSData = getLocVar(dataWS, wsName, True, "target worksheet name")
'        If col_tarWSData = 0 Then
'            g_log.RaiseWarning "Cannot Find the Worksheet " & wsName & " in the 'WSData' Sheet. Skip this row.", CannotFindWSinWSData
'            GoTo ExitFunc
'        End If
'        macroName = .Cells(rRow_tarWSMacro, col_tarWSData + 1)
'
'        lRow_inputData = FindLastRow(rRow_tarWSData + 1, col_tarWSData, dataWS)
'        lRow_returnData = FindLastRow(rRow_tarWSData + 1, col_tarWSData + 2, dataWS)
'
'        'Clear the previous data
'
'        For i = rRow_tarWSData + 1 To lRow_inputData
'            key = .Cells(i, col_tarWSData).Text
'            If tarWS_iDataLoc.Exists(key) Then
'                g_log.RaiseWarning "Duplicate Input Label: '" & key & "' is found. Worksheet: '" & wsName & "'. Record Skipped. Please Check 'WSData'.", duplicateLabelInSummary
'            Else
'                tarWS_iDataLoc.Add .Cells(i, col_tarWSData).Text, .Cells(i, col_tarWSData + 1)
'            End If
'        Next i
'
'        For i = rRow_tarWSData + 1 To lRow_returnData
'            key = .Cells(i, col_tarWSData).Text
'            If tarWS_rDataLoc.Exists(key) Then
'                g_log.RaiseWarning "Duplicate Output Label: '" & key & "' is found. Worksheet: '" & wsName & "'. Record Skipped. Please Check 'WSData'.", duplicateLabelInSummary
'            Else
'                tarWS_rDataLoc.Add .Cells(i, col_tarWSData + 2).Text, .Cells(i, col_tarWSData + 3)
'            End If
'        Next i
'
'        iDataDictofDict.Add wsName, tarWS_iDataLoc
'        rDataDictofDict.Add wsName, tarWS_rDataLoc
'    End With
'    Exit Function
'
'ExitFunc:
'    ReadDSData = -1
'End Function

'Private Function CreateColumnTag(colNum As Long, name As String) As ColumnTag
'    Dim obj As ColumnTag
'    obj.columnNum = colNum
'    obj.name = name
'    CreateColumnTag = obj
'End Function

'Private Function CreateDesignWsData() As DesignWorksheet
'    Dim obj As DesignWorksheet
'    CreateDesignWsData = obj
'End Function
Private Function CreateSequenceArray(startNum As Long, endNum As Long, Optional arrBase As Long = 0) As Variant
    Dim arr() As Long, n As Long
    Dim i As Long, count As Long
    
    n = endNum - startNum + 1
    ReDim arr(arrBase To arrBase + n - 1)
    
    count = 0
    For i = arrBase To UBound(arr)
        arr(i) = startNum + count
        count = count + 1
    Next i
    
    CreateSequenceArray = arr
End Function
