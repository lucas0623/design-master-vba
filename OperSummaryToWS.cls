VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OperSummaryToWS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Operation"
Private mCompleteMessageAddition As String
Private mTerminateMessageMain As String, mTerminateMessageAddition As String
Private Const mOperName = "SummaryToWs"

Private genFuncs As New clsGeneralFunctions, wsInteract As New clsWorksheetsInteraction
Private DsSys As New DataSheetSystem
Private fRow_summary As Long, lRow_summary As Long
Private dataLabels As Object

Private sumWS As Worksheet
Private dataWS As Worksheet
Private col_tarWSName As Long
Private rRow_tarWSMacro As Long, rRow_tarWSData As Long

Private iDataDictofDict As Object 'a dictionary containing dictionary of 'tarWS_iDataLoc'
'Private tarWS_iDataLoc As Object 'a dictionary containing the input setting of certain worksheet

Private rDataDictofDict As Object 'a dictionary containing dictionary of 'tarWS_rDataLoc'
'Private tarWS_rDataLoc As Object 'a dictionary containing the output setting of certain worksheet

Private isAllowArrInput As Boolean
Private isBoldedOnly As Boolean
Private isOutputPDF As Boolean
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
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    g_log.WriteLog "Operation: 'SummaryToDsWs' Started."
    
    If summarySheet Is Nothing Then
        Set sumWS = ActiveSheet
    Else
        Set sumWS = summarySheet
    End If
    
    ret = GetUserInput
    If Not ret = 0 Then
        g_log.WriteLog "Userform canceled. Exit Operation."
        GoTo ExitFunc
    End If
    
    ret = Initialize
    If Not ret = 0 Then
        g_log.WriteLog mTerminateMessageMain
        GoTo ExitFunc
    End If
    
    g_log.WriteLog "Summary Worksheet: " & sumWS.name
    g_log.WriteLog "Row Range: " & fRow_summary & " to " & lRow_summary
    For i = fRow_summary To lRow_summary
        
        If (sumWS.rows(i).height > 0 And isBoldedOnly And sumWS.Cells(i, col_tarWSName).Font.Bold = True) Or _
            (sumWS.rows(i).height > 0 And Not isBoldedOnly) Then
            
            'Tranfer Data
            
            i = TransferData(i)
            
            'Print pdf if required
            
        End If
        
    Next i
    Set tarWs = Nothing
'    If isMissingKey Then
'        MsgBox "COMPLETED with Warning: There are missing data Labels in the first Row of the summary sheet. Please check and add the missing data labels if necessary"
'    End If
    sumWS.Activate
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    mCompleteMessageAddition = countSucessful & " out of " & countTotal & " number of rows of data are sucessfully transferred."
    Exit Function
ExitFunc:
    Main = -1
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
    form.AddCheckBox CheckBox1, "is Allow for Array Input?", isCheck:=False
    form.AddCheckBox CheckBox2, "is Output to the Bolded Rows Only?", isCheck:=False
    form.AddCheckBox CheckBox3, "Print PDF?"
    form.AdjustHeight
    LoadUserInput CheckBox1, CheckBox2, CheckBox3
    form.Show
    
    If form.CloseState = 0 Then
        'Get Input
        isAllowArrInput = CheckBox1.value
        isBoldedOnly = CheckBox2.value
        isOutputPDF = CheckBox3.value
        fRow_summary = Range(textBox1.Text).rows(1).row
        lRow_summary = Range(textBox1.Text).rows.count + fRow_summary - 1
        SaveUserInput
    Else
        ret = -1
        
    End If
    
    GetUserInput = ret
'    'Get Text File
'    strPath = GetFilePath_txt
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
        DsSys.CreateCustomPropGrp propGrpName, "wsName", "IsArrayInput", _
                                                "IsBoldedRows", "isPrintPDF"

    End If

    '2. Write Data
    DsSys.prop(propGrpName, "wsName") = sumWS.name
    'ds_sys.prop(propGrpName, "isBySec") = mbIsBySec
    'ds_sys.prop(propGrpName, "Rows") = msLcSelect
    DsSys.prop(propGrpName, "IsArrayInput") = isAllowArrInput
    DsSys.prop(propGrpName, "IsBoldedRows") = isBoldedOnly
    DsSys.prop(propGrpName, "isPrintPDF") = isOutputPDF
End Function

Private Function LoadUserInput(cbIsAllowInput As msForms.checkbox, cbIsBoldedOnly As msForms.checkbox, _
                                cbIsOutputPDF As msForms.checkbox) As Integer
    Dim propGrpName As String
    
    propGrpName = sumWS.name & mOperName
    '1.Check if there is existing save. If yes, try loading the data
    If Not DsSys.isPropGrpExist(propGrpName) Then
        g_log.WriteLogInDetailMode "Cannot find the propGrpName in Data_System Sheet"
        ret = -1
        GoTo ExitFunc
    End If
    cbIsAllowInput.value = DsSys.prop(propGrpName, "IsArrayInput")
    cbIsBoldedOnly.value = DsSys.prop(propGrpName, "IsBoldedRows")
    cbIsOutputPDF.value = DsSys.prop(propGrpName, "isPrintPDF")
    'tbRngInput.Text = ds_sys.prop(propGrpName, "Rows")
ExitFunc:
    LoadUserInput = ret
End Function

Private Function Initialize() As Integer
    Set DsSys = New DataSheetSystem
    Set dataLabels = CreateObject("Scripting.Dictionary")
    Set iDataDictofDict = CreateObject("Scripting.Dictionary")
    Set rDataDictofDict = CreateObject("Scripting.Dictionary")
    
    
    Set dataWS = setWorksheet("WSData", "that contains the target worksheet input/output data")
    
    col_tarWSName = getLocVar(sumWS, "targetWS", True, "'targetWS' text (used for locating the target worksheet name's column)")
    If col_tarWSName = 0 Then
        mTerminateMessageMain = "Cannot Find the 'targetWS' tag in the first row of the summary sheet. Please check!"
        Initialize = -1
        Exit Function
    End If
    rRow_tarWSMacro = getLocVar(dataWS, "rRow_tarWSMacro", False, _
                                "'rRow_tarWSMacro' text (used for locating the row for input the 'WS macro Ref Num')")
    rRow_tarWSData = getLocVar(dataWS, "rRow_tarWSData", False, _
                               "'rRow_tarWSData' text (used for locating the reference row of input/output data of the target Worksheets)")
    
    If isOutputPDF Then
        myFilePath = GetFilePath(DsSys.prop("OutputFolderPath", "path"))
        If myFilePath = vbNullString Then
            Initialize = -1
            Exit Function
        End If
        DsSys.prop("OutputFolderPath", "path") = myFilePath & "\"
    End If
    
    
    'Get the dataLabels in the summary sheet
    dataLabels.RemoveAll
    On Error GoTo err_duplicateDataLabel
    With sumWS
        For Each cell In .Range(Cells(1, 1), Cells(1, 500))
            If Not IsEmpty(cell) Then
                'Debug.Print cell.Address
                'do sth
                dataLabels.Add cell.Text, cell.column
            End If
        Next
    End With
    On Error GoTo 0
    
    
    Exit Function
    
err_duplicateDataLabel:
    g_log.RaiseWarning "There are duplicate data Labels in the first Row of the summary sheet >_< (name: " & cell.Text & ")", duplicateLabelInSummary
    Resume Next
End Function

Private Function TransferData(cRow As Long) As Long
    'return the last row the input data
    
    Dim i As Long
    Dim tarWS_iDataLoc As Object 'a dictionary containing the input setting of certain worksheet
    Dim tarWS_rDataLoc As Object
    Dim tarWsName As String, tarWs As Worksheet
    Dim lRow As Long
    Dim ret As Integer
    
    countTotal = countTotal + 1
    tarWsName = sumWS.Cells(cRow, col_tarWSName)
    If tarWsName = vbNullString Then
        g_log.RaiseWarning "No Target Worksheet Name Input is Found in row " & cRow & " of the summary. Skip this row.", EmptyWsNameInSummary
        GoTo ExitFunc
    End If
    
    Set tarWs = wsInteract.setWorksheet(tarWsName, isSkipErr:=True)
    If tarWs Is Nothing Then
        g_log.RaiseWarning "No Worksheet '" & tarWsName & "' is Found in the workbook. Skip this row.", NoWorksheetFoundInWorkbook
        GoTo ExitFunc
    End If
    
    If Not iDataDictofDict.Exists(tarWsName) Then
        ret = ReadDSData(tarWsName)
        If Not ret = 0 Then
            GoTo ExitFunc
        End If
    End If
    
    Set tarWS_iDataLoc = iDataDictofDict(tarWsName)
    Set tarWS_rDataLoc = rDataDictofDict(tarWsName)
    
    'For Array Input Use. Find the range of rows
    If isAllowArrInput Then
        Dim nextNonEmptyCell As Range
        Set nextNonEmptyCell = wsInteract.FindNextNonEmptyCell(sumWS.Cells(cRow, col_tarWSName), lRow_summary)
        If Not nextNonEmptyCell Is Nothing And Not nextNonEmptyCell.row = cRow Then
            lRow = nextNonEmptyCell.row - 1
        Else
            lRow = lRow_summary
        End If
    Else
        lRow = cRow
    End If
    
    If isAllowArrInput And Not lRow = cRow Then
        TransferMultipleRows cRow, lRow, tarWs, tarWS_iDataLoc, tarWS_rDataLoc
    Else
        TransferSingleRow cRow, tarWs, tarWS_iDataLoc, tarWS_rDataLoc
    End If
    
    'run macro if necessary
    If Not macroName = vbNullString Then
        tarWs.Calculate
        Application.Run macroName
    End If
    
    'Return data from the target WS
    tarWs.Calculate

    For Each key In tarWS_rDataLoc.keys
        If dataLabels.Exists(key) Then
            g_log.WriteLog "  Output Tag: '" & key & "';  Design Worksheet: '" & tarWs.name & "'."
            sumWS.Cells(cRow, dataLabels(key)) = tarWs.Range(tarWS_rDataLoc(key))
            g_log.WriteLog "    Trasnferring Data from '" & tarWs.name & "'!" & tarWs.Range(tarWS_rDataLoc(key)).address & " to " & _
                            "'" & sumWS.name & "'!" & sumWS.Cells(cRow, dataLabels(key)).address
        Else
            isMissingKey = True
        End If
    Next
    
    If isOutputPDF Then
        strCount = VBA.Format(countPDF, "00")
        Create_PDF tarWs, myFilePath, strCount
        countPDF = countPDF + 1
    End If
    
    TransferData = lRow
    countSucessful = countSucessful + 1
    Exit Function
    
ExitFunc:
    TransferData = cRow
End Function


'if isRow = true, will search the first row
Private Function ReadDSData(wsName As String) As Integer

    If iDataDictofDict.Exists(wsName) Then Exit Function
    
    'Read the Design Worksheet Setting Data in the 'WS_Data' Sheet
    Dim col_tarWSData As Long
    Dim lRow_inputData As Long, lRow_returnData As Long
    Dim tarWS_iDataLoc As Object 'a dictionary containing the input setting of certain worksheet
    Dim tarWS_rDataLoc As Object 'a dictionary containing the output setting of certain worksheet
    Dim key As String

    Set tarWS_iDataLoc = CreateObject("Scripting.Dictionary")
    Set tarWS_rDataLoc = CreateObject("Scripting.Dictionary")
    

    With dataWS
        col_tarWSData = getLocVar(dataWS, wsName, True, "target worksheet name")
        If col_tarWSData = 0 Then
            g_log.RaiseWarning "Cannot Find the Worksheet " & wsName & " in the 'WSData' Sheet. Skip this row.", CannotFindWSinWSData
            GoTo ExitFunc
        End If
        macroName = .Cells(rRow_tarWSMacro, col_tarWSData + 1)

        lRow_inputData = FindLastRow(rRow_tarWSData + 1, col_tarWSData, dataWS)
        lRow_returnData = FindLastRow(rRow_tarWSData + 1, col_tarWSData + 2, dataWS)

        'Clear the previous data
        
        For i = rRow_tarWSData + 1 To lRow_inputData
            key = .Cells(i, col_tarWSData).Text
            If tarWS_iDataLoc.Exists(key) Then
                g_log.RaiseWarning "Duplicate Input Label: '" & key & "' is found. Worksheet: '" & wsName & "'. Record Skipped. Please Check 'WSData'.", duplicateLabelInSummary
            Else
                tarWS_iDataLoc.Add .Cells(i, col_tarWSData).Text, .Cells(i, col_tarWSData + 1)
            End If
        Next i
        
        For i = rRow_tarWSData + 1 To lRow_returnData
            key = .Cells(i, col_tarWSData).Text
            If tarWS_rDataLoc.Exists(key) Then
                g_log.RaiseWarning "Duplicate Output Label: '" & key & "' is found. Worksheet: '" & wsName & "'. Record Skipped. Please Check 'WSData'.", duplicateLabelInSummary
            Else
                tarWS_rDataLoc.Add .Cells(i, col_tarWSData + 2).Text, .Cells(i, col_tarWSData + 3)
            End If
        Next i
        
        iDataDictofDict.Add wsName, tarWS_iDataLoc
        rDataDictofDict.Add wsName, tarWS_rDataLoc
    End With
    Exit Function
    
ExitFunc:
    ReadDSData = -1
End Function

Private Sub TransferSingleRow(cRow As Long, tarWs As Worksheet, tarWS_iDataLoc As Object, tarWS_rDataLoc As Object)
    g_log.WriteLog "Transferring Data for Summary Row Number " & cRow & "."
    
    Dim key As Variant
    'Read the data from the summary and paste it to the target worksheet
    For Each key In tarWS_iDataLoc.keys
        If dataLabels.Exists(key) Then
            g_log.WriteLog "  Input Tag: '" & key & "';  Design Worksheet: '" & tarWs.name & "'."
            tarWs.Range(tarWS_iDataLoc(key)).value2 = sumWS.Cells(cRow, dataLabels(key)).value2
            g_log.WriteLog "    Trasnferring Data from '" & sumWS.name & "'!" & sumWS.Cells(cRow, dataLabels(key)).address & " to " & _
                            "'" & tarWs.name & "'!" & tarWs.Range(tarWS_iDataLoc(key)).address
        Else
            g_log.RaiseWarning "Data Label: '" & CStr(key) & "' of Design Worksheet '" & tarWs.name & "' is missing in the summary sheet. ", missingLabelInSummary
            isMissingKey = True
        End If
    Next
End Sub

Private Sub TransferMultipleRows(cRow As Long, lRow As Long, tarWs As Worksheet, tarWS_iDataLoc As Object, tarWS_rDataLoc As Object)
    g_log.WriteLog "Transferring Data for Summary Row Number " & cRow & " to " & lRow & "."
    
    Dim key As Variant, numRows As Long
    Dim sumFirstCell As Range, sumNextEmptyCell As Range, sumLastCell As Range
    Dim tarFirstCell As Range, tarLastCell As Range, tarLastCellForClean As Range
    'Read the data from the summary and paste it to the target worksheet
    For Each key In tarWS_iDataLoc.keys
        If dataLabels.Exists(key) Then
            g_log.WriteLog "  Input Tag: '" & key & "';  Design Worksheet: '" & tarWs.name & "'."
            Set sumFirstCell = sumWS.Cells(cRow, dataLabels(key))
            Set sumNextEmptyCell = wsInteract.FindNextEmptyCell(sumFirstCell, lRow)
            If sumNextEmptyCell Is Nothing Then
                Set sumLastCell = sumWS.Cells(lRow, dataLabels(key))
            Else
                Set sumLastCell = sumWS.Cells(sumNextEmptyCell.row - 1, dataLabels(key))
            End If
            numRows = sumLastCell.row - sumFirstCell.row + 1
            Set tarFirstCell = tarWs.Range(tarWS_iDataLoc(key))
            Set tarLastCell = tarWs.Cells(tarFirstCell.row + numRows - 1, tarFirstCell.column)
            If numRows > 1 Then
                Set tarLastCellForClean = tarWs.Cells(wsInteract.FindLastRow(tarFirstCell.row, tarFirstCell.column, tarFirstCell.Worksheet), tarFirstCell.column)
                tarWs.Range(tarFirstCell, tarLastCellForClean).Clear
                g_log.WriteLog "    '" & tarWs.name & "'!" & tarFirstCell.address & ":" & tarLastCellForClean.address & " is cleared."
            End If
            tarWs.Range(tarFirstCell, tarLastCell).value2 = sumWS.Range(sumFirstCell, sumLastCell).value2
            g_log.WriteLog "    Trasnferring Data from '" & sumWS.name & "'!" & sumFirstCell.address & ":" & sumLastCell.address & " to " & _
                            "'" & tarWs.name & "'!" & tarFirstCell.address & ":" & tarLastCell.address

        Else
            g_log.RaiseWarning "Data Label: '" & CStr(key) & "' of Design Worksheet '" & tarWs.name & "' is missing in the summary sheet. ", missingLabelInSummary
            isMissingKey = True
        End If
    Next
    'g_log.WriteLog "Summary Row Number " & cRow & " to " & lRow & " data input to worksheet '" & tarWs.Name & "' completed."
End Sub

Private Function getLocVar(ws As Worksheet, label As String, isRow As Boolean, Optional errText As String = "correct data") As Long
    On Error GoTo Err_noMatch

    With ws
        If isRow Then
            getLocVar = Application.Match(label, .rows(1), 0)
        Else
            getLocVar = Application.Match(label, .Columns(1), 0)
        End If
    End With
    
    On Error GoTo 0

    Exit Function
    
Err_noMatch:
    
End Function

Private Function setWorksheet(wsName As String, Optional errText As String = "") As Worksheet
    
    On Error GoTo Err_cantFindWS
    
    Set setWorksheet = Worksheets(wsName)
    
    On Error GoTo 0
    Exit Function
    
Err_cantFindWS:
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    MsgBox "Err: Cannot find the Worksheet '" & wsName & "' " & errText & " :(" & vbNewLine & _
           "The worksheet shall be name as '" & wsName & "' and follow the required format. "
    End
End Function

Private Function FindLastRow(ByVal ref_Row As Long, ByVal refCol As Long, Optional ws As Worksheet) As Long
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    With ws
        If .Cells(ref_Row, refCol).Text = "" Then
            FindLastRow = ref_Row
'            MsgBox "NO DATA IN THE WS"
        ElseIf .Cells(ref_Row + 1, refCol).Text <> "" Then
            FindLastRow = .Cells(ref_Row, refCol).End(xlDown).row

        Else
            FindLastRow = ref_Row
        End If
    End With
End Function

Private Sub Create_PDF(myWS As Worksheet, myFilePath As String, myFileName As String)
    
    'Dim ws As Worksheet
    'Dim strPath As String
    'Dim myFile As Variant
    Dim strFile As String
    'On Error GoTo ErrHandler
    
    
    strFile = myFilePath & "\" & myFileName

    'If myFile <> "False" Then
        myWS.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            fileName:=strFile, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
    
        'MsgBox "PDF file has been created."
    'End If
    
    'On Error GoTo 0
    Exit Sub
    
ErrHandler:
    g_log.RaiseWarning "Could not create PDF file! Record Skipped.", failToCreatePDF
    Resume Next
End Sub

Private Function GetFilePath(Optional defaultPath As String) As String
    
    Dim path As String
    Dim checkPath As String
    On Error Resume Next
    checkPath = Dir(defaultPath, vbDirectory)
    On Error GoTo 0
    If checkPath = "" Then
        path = ActiveWorkbook.path & "\"
    Else
        path = defaultPath
    End If
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = path
        .title = "Please choose a folder"
        .AllowMultiSelect = False
    
        If .Show = -1 Then
            GetFilePath = .SelectedItems(1) '=-1 if action button is clicked, =0 if cancel button is clicked
        Else:
            GetFilePath = vbNullString
        End If
    End With

'    Dim strFileToOpen As String
'    strFileToOpen = Application.GetOpenFilename _
'    (Title:="Please choose the file path for excel workbooks containing SAP output")
'    'FileFilter:="Excel Workbooks *.xls* (*.xls*),")
'
'    If strFileToOpen = "False" Then
'        'MsgBox "No file selected.", vbExclamation, "Sorry!"
'        Exit Function
'    Else
'        GetFilePath = strFileToOpen
'    End If
    
End Function


