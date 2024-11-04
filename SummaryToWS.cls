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

'Summary
Private sumWS As Worksheet, columnTags As Object 'a dictionary that contains all column tags data in the summary sheet
Private targetWsTags As Object 'a dictionary that contains all 'targetWS' Tag in the summary sheet
Private fRow_summary As Long, lRow_summary As Long, rows As Long
Private activeRows() As Boolean 'The rows that will be updated/ calculated, which = selected rows + (bolded rows), based 1
Private activeOutputCols As Object 'a dictionary that contains all active output columns to be written back to the summary sheet after calculation
Private df_sum As clsDataFrame

Private dataWS As Worksheet
Private designWorksheets As Object 'a dictionary that contains all designWorksheet input/output data (dictionary of 'DesignWsOP' Object)

'Userform and Print PDF variables
Private isBoldedOnly As Boolean, isOutputPDF As Boolean, isIncludePrefix As Boolean
Private myFilePath As String
Private strCount As String, countPDF As Long
Private countSucessful As Long, countTotal As Long 'count how many rows of data are sucessfully transferred.

Private isMissingKey As Boolean
'Private macroName As String, objName As String, procName As String

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
        mTerminateMessageMain = "Userform canceled. Exit Operation."
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
    
    ret = FormColumnTags
    If Not ret = 0 Then
    End If
    
    ret = IdentifyActiveRows
    If Not ret = 0 Then
    End If
    
    ret = TransferDataFromSummary
    If Not ret = 0 Then
    End If
    
    ret = WriteResultToSummary
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
    Set activeOutputCols = CreateObject("Scripting.Dictionary")
    
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
        If Not designWorksheets.Exists(nonEmptyCells(i).value) Then
            Set designWorksheet = ReadSingleDesignWorksheetData(nonEmptyCells(i), refRow)
            designWorksheets.Add designWorksheet.name, designWorksheet
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

Private Function FormColumnTags() As Integer
    Dim ws As Worksheet
    Dim colTag As SummaryWsColTags
    Dim nonEmptyCells() As Range, cell As Variant
        
    Set ws = sumWS
    nonEmptyCells = wsInteract.FindNonEmptyCell(ws.rows(1))
    
    For Each cell In nonEmptyCells
        Set colTag = New SummaryWsColTags
        colNum = cell.column
        
        ' Initialize the object with name and column number
        colTag.Initialize CStr(cell.value), colNum
        
        ' Check if the cell value contains "targetWS"
        If InStr(1, cell.value, "targetWS", vbTextCompare) > 0 Then
            colTag.isTargetWsTag = True
            If Not targetWsTags.Exists(colTag.name) Then targetWsTags.Add colTag.name, colTag
        Else
            colTag.isOutputTag = True
            If Not columnTags.Exists(colTag.name) Then columnTags.Add colTag.name, colTag
        End If
        colTag.colNumInDF = df_sum.columnNum(colTag.name)
    Next
    Exit Function
End Function

Private Function IdentifyActiveRows() As Integer
    ReDim activeRows(1 To df_sum.CountRows)
    Dim i As Long
    Dim rCol As Long
    rCol = targetWsTags.Items()(0).colNumInWs
    If isBoldedOnly Then
        For i = 1 To UBound(activeRows)
            If sumWS.Cells(fRow_summary + i - 1, rCol).Font.Bold = True And sumWS.rows(fRow_summary + i - 1).height > 0 Then
                activeRows(i) = True
            End If
        Next i
    Else
        For i = 1 To UBound(activeRows)
            If sumWS.rows(fRow_summary + i - 1).height > 0 Then
                activeRows(i) = True
            End If
        Next i
    End If
    
End Function

'Private Function GetActiveTargetWS() As Integer
'
'End Function
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

Private Function TransferDataFromSummary() As Integer
    Dim i As Long, j As Long
    Dim targetWsTag As SummaryWsColTags, targetWs As DesignWsIO
    Dim ret As Long, key As Variant
    For i = 1 To df_sum.CountRows
        If Not activeRows(i) Then GoTo NextIteration
        For Each key In targetWsTags
            Set targetWsTag = targetWsTags(key)
            Set targetWs = designWorksheets(df_sum.idata(i, targetWsTag.colNumInDF))
            ret = TransferSingleDesignWorksheetData(targetWs, i)
        Next key
NextIteration:
    Next i
End Function

Private Function TransferSingleDesignWorksheetData(targetWsIO As DesignWsIO, dfRowNum As Long) As Integer
    Dim WsInputs As Object, WsOutputs As Object
    Dim key As Variant
    Dim targetWs As Worksheet
    Dim tag As String
    
    Set WsInputs = targetWsIO.WsInput
    Set WsOutputs = targetWsIO.WsOutput
    
    Set targetWs = wsInteract.setWorksheet(targetWsIO.name, isSkipErr:=True)
    If targetWs Is Nothing Then GoTo ExitFunc:
    
    For Each key In WsInputs.keys
        tag = key
        If columnTags.Exists(tag) Then
            targetWs.Range(WsInputs(tag)) = df_sum.idata(dfRowNum, columnTags(tag).colNumInDF)
        End If
    Next key
    
    targetWs.Calculate
    For Each key In WsOutputs.keys
        tag = key
        If columnTags.Exists(tag) Then
            df_sum.idata(dfRowNum, columnTags(tag).colNumInDF) = targetWs.Range(WsOutputs(tag))
            If Not activeOutputCols.Exists(tag) Then
                activeOutputCols.Add tag, columnTags(tag)
            End If
        End If
    Next key
    
    Dim pdfName As String
    If isOutputPDF Then
        strCount = VBA.Format(countPDF, "00")
        If isIncludePrefix Then
            pdfName = targetWs.name & "-" & strCount
        Else
            pdfName = strCount
        End If
        Create_PDF targetWs, myFilePath, pdfName
        countPDF = countPDF + 1
    End If
    
    Exit Function
ExitFunc:
    ret = -1
    g_log.RaiseWarning "No Worksheet '" & targetWsIO.name & "' is Found in the workbook. Skip this row.", NoWorksheetFoundInWorkbook
End Function

Private Function WriteResultToSummary() As Long
    
    Dim key As Variant, tagObj As SummaryWsColTags
    Dim colWs As Long, colDf As Long
    For Each key In activeOutputCols
        Set tagObj = activeOutputCols(key)
        colDf = tagObj.colNumInDF
        colWs = tagObj.colNumInWs
        wsInteract.WriteArrToColumn df_sum.iColumn(colDf), fRow_summary, colWs, sumWS
    Next key
    
End Function
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
Private Sub Create_PDF(myWS As Worksheet, myFilePath As String, myFileName As String)
    

    Dim strFile As String
    strFile = myFilePath & "\" & myFileName

        myWS.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            fileName:=strFile, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
    
    Exit Sub
    
ErrHandler:
    g_log.RaiseWarning "Could not create PDF file! Record Skipped.", failToCreatePDF
    Resume Next
End Sub

