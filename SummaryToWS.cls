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

Private dataWS As Worksheet, designGroupWS As Worksheet
Private designWorksheetsIO As Object 'a dictionary that contains all designWorksheet input/output data (dictionary of 'DesignWsOP' Object)
Private designWsGroups As Object 'a dictionary that contains all desing Ws Group data

'Userform and Print PDF variables
Private isBoldedOnly As Boolean, isOutputPDF As Boolean, isIncludePrefix As Boolean
Private myFilePath As String
Private strCount As String, countPDF As Long
Private countSucessful As Long, countTotal As Long 'count how many rows of data are sucessfully transferred.

Private isMissingKey As Boolean

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
        g_log.WriteLog "Fail to read worksheet Input/Output data. Exit Operation."
        GoTo ExitFunc
    End If
    
    ret = ReadDesignGroupData
    If Not ret = 0 Then
        g_log.WriteLog "Fail to read Design Worksheet Group data. Exit Operation."
        GoTo ExitFunc
    End If
    
    ret = FormColumnTags
    If Not ret = 0 Then
    End If
    
    ret = IdentifyActiveRows
    If Not ret = 0 Then
    End If
    
    g_log.WriteLog "Summary Worksheet: " & sumWS.Name
    g_log.WriteLog "Row Range: " & fRow_summary & " to " & lRow_summary
    
    ret = TransferDataFromSummary
    If Not ret = 0 Then
    End If
    
    ret = WriteResultToSummary
    If Not ret = 0 Then
    End If
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    mCompleteMessageAddition = countSucessful & " out of " & countTotal & " number of worksheets' data are sucessfully transferred."
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
    Set designWorksheetsIO = CreateObject("Scripting.Dictionary")
    Set designWsGroups = CreateObject("Scripting.Dictionary")

    Set activeOutputCols = CreateObject("Scripting.Dictionary")
    
    Set dataWS = wsInteract.setWorksheet("WSData", isSkipErr:=True)
    If dataWS Is Nothing Then
        Dim oper As New InitializeWorkbook
        oper.InitializeWSData ("WSData")
        mTerminateMessageMain = "No WSData Information! Please input target design worksheet input/output data in the 'WSData' Sheet!"
        GoTo ExitFunc:
    End If
    
    Set designGroupWS = wsInteract.setWorksheet("DesignWsGroupData", isSkipErr:=True)
    
    If summarySheet Is Nothing Then Set sumWS = ActiveSheet
    Set coll = wsInteract.FindAll(sumWS.rows(1), "targetWS")
    If coll.count = 0 Then
        mTerminateMessageMain = "Cannot Find any 'targetWS' tag in the first row of the summary sheet. Please check!"
        Initialize = -1
        Exit Function
    End If

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
        GoTo ExitFunc
    End If
    
    If isOutputPDF Then
        Dim UI As New clsUIManager
        myFilePath = UI.GetFolderPath(DsSys.prop("OutputFolderPath", "path"))
        If myFilePath = vbNullString Then
            ret = -1
            GoTo ExitFunc
        End If
        DsSys.prop("OutputFolderPath", "path") = myFilePath & "\"
    End If
    
ExitFunc:
    GetUserInput = ret
End Function

Private Function SaveUserInput() As Integer
    Dim propGrpName As String
    
    propGrpName = sumWS.Name & mOperName
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
    DsSys.prop(propGrpName, "wsName") = sumWS.Name
    DsSys.prop(propGrpName, "isBoldedRows") = isBoldedOnly
    DsSys.prop(propGrpName, "isPrintPDF") = isOutputPDF
    DsSys.prop(propGrpName, "isIncludePrefix") = isIncludePrefix
End Function

Private Function LoadUserInput(cbIsBoldedOnly As msForms.checkbox, _
                                cbIsOutputPDF As msForms.checkbox, cbIsIncludePrefix As msForms.checkbox) As Integer
    Dim propGrpName As String
    
    propGrpName = sumWS.Name & mOperName
    '1.Check if there is existing save. If yes, try loading the data
    If Not DsSys.isPropGrpExist(propGrpName) Then
        g_log.WriteLogInDetailMode "Cannot find the propGrpName in Data_System Sheet"
        ret = -1
        GoTo ExitFunc
    End If
    
    cbIsBoldedOnly.value = DsSys.prop(propGrpName, "isBoldedRows")
    cbIsOutputPDF.value = DsSys.prop(propGrpName, "isPrintPDF")
    cbIsIncludePrefix.value = DsSys.prop(propGrpName, "isIncludePrefix")

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
    'console.Log df_sum.heads, "df_sum.heads"
    'console.Log df_sum.data("1", "1", True), "df_sum.data"
    If df_sum.CountRows = 0 Then GoTo ExitFunc
   
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
    
    g_log.WriteLog "Start reading design worksheets' Input/Output Data..."
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
        If Not designWorksheetsIO.Exists(nonEmptyCells(i).value) Then
            g_log.WriteLog "   Reading worksheet '" & nonEmptyCells(i).value & "' input/output data"
            Set designWorksheet = ReadSingleDesignWorksheetData(nonEmptyCells(i), refRow)
            
            designWorksheetsIO.Add designWorksheet.Name, designWorksheet
            'console.Log designWorksheet.WsInput, designWorksheet.Name & " INPUT"
            'console.Log designWorksheet.WsInput, designWorksheet.Name & " OUTPUT"
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
    Dim tempStr As String
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
        
        If value = vbNullString Then GoTo NextIteration1
        If Not wsInteract.IsCellAddress(value) Then
            g_log.RaiseWarning "Cell Address invalid. Worksheet Name: " & designObject.Name & ". Input: " & key & ".", InvalidAddress
            GoTo NextIteration1
        End If
        tempStr = wsInteract.GetAddressWorksheetName(value)
        If Not tempStr = vbNullString And Not tempStr = designObject.Name Then
            g_log.RaiseWarning "Cell Address invalid: Worksheet Name Not Match. Worksheet Name: " & designObject.Name & ". Input: " & key & ".", InvalidAddress
            GoTo NextIteration1
        End If
        
        ret = designObject.AddInput(key, value)
        If Not ret = 0 Then g_log.RaiseWarning "Duplicate Input Name. Worksheet Name: " & designObject.Name & ". Input: " & key & ".", DuplicateIOTagName
        g_log.WriteLog "   Input: '" & key & "' address sucessfully read"
NextIteration1:
    Next i
    
    ' Read and set Output values
    lRow = wsInteract.FindLastRow(refRow, rCol + 2, ws)
    For i = fRow To lRow
        key = ws.Cells(i, rCol + 2).value
        value = ws.Cells(i, rCol + 3).value
        
        If value = vbNullString Then GoTo NextIteration2
        
        If Not wsInteract.IsCellAddress(value) Then
            g_log.RaiseWarning "Cell Address invalid. Worksheet Name: " & designObject.Name & ". Output: " & key & ".", InvalidAddress
            GoTo NextIteration2
        End If
        
        tempStr = wsInteract.GetAddressWorksheetName(value)
        If Not tempStr = vbNullString And Not tempStr = designObject.Name Then
            g_log.RaiseWarning "Cell Address invalid: Worksheet Name Not Match. Worksheet Name: " & designObject.Name & ". Input: " & key & ".", InvalidAddress
            GoTo NextIteration1
        End If
        
        ret = designObject.AddOutput(key, value)
        If Not ret = 0 Then g_log.RaiseWarning "Duplicate Output Name. Worksheet Name: " & designObject.Name & ". Output: " & key & ".", DuplicateIOTagName
        g_log.WriteLog "   Output: '" & key & "' address sucessfully read"
NextIteration2:
    Next i
    
    Set ReadSingleDesignWorksheetData = designObject
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
            If Not targetWsTags.Exists(colTag.Name) Then targetWsTags.Add colTag.Name, colTag
        Else
            colTag.isOutputTag = True
            If Not columnTags.Exists(colTag.Name) Then columnTags.Add colTag.Name, colTag
        End If
        colTag.colNumInDF = df_sum.columnNum(colTag.Name)
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
Private Function ReadDesignGroupData() As Long
    Dim ws As Worksheet
    
    Set ws = Worksheets("DesignWsGroupData")
    Set designWsGroups = CreateObject("Scripting.Dictionary")

    Dim refRow As Long
    Dim col As Long
    Dim i As Long
    Dim nonEmptyCells() As Range
    Dim designGroup As DesignWsGroup
    
    ' Find all non-empty cells in the first row
    g_log.WriteLog "Start reading design worksheets group Input/Output Data..."
    nonEmptyCells = wsInteract.FindNonEmptyCell(ws.rows(1))
    
    If Not genFuncs.isInitialised(nonEmptyCells) Then
        'Write log
        GoTo ExitFunc
    End If
    
    ' Find the text 'rRow' in the first row
    refRow = wsInteract.getLocVar(ws, "rRow", False, isMustMatch:=False)
    If refRow = -1 Then
        'Write log
        GoTo ExitFunc
    End If
    
    
    ' Loop through each non-empty cell found in the first row
    For i = LBound(nonEmptyCells) To UBound(nonEmptyCells)
        If Not designWsGroups.Exists(nonEmptyCells(i).value) Then
            g_log.WriteLog "   Reading worksheet group'" & nonEmptyCells(i).value & "' data"
            Set designGroup = ReadSingleDesignGroupData(nonEmptyCells(i), refRow)
            If designGroup Is Nothing Then
                mTerminateMessageMain = "Check 'DesignWsGroupData' input! Some errors are found in the cells. The macro will be terminated."
                GoTo ExitFunc
            End If
            
            designWsGroups.Add designGroup.Name, designGroup
            'console.Log designGroup.TagsRelationships.item(1).tagRelation, designGroup.TagsRelationships.item(1).Name
            'console.Log designGroup.TagsRelationships.item(6).tagRelation, designGroup.TagsRelationships.item(6).Name
        Else
            g_log.RaiseWarning "Duplicate Design Worksheet Group in 'DesignWsGroupData'. Only the first sheet is considered.", duplicateDesignWorksheet 'Raise Warning
        End If
    Next i
    Exit Function
ExitFunc:
    ReadDesignGroupData = -1
End Function

Function ReadSingleDesignGroupData(refCell As Range, refRow As Long) As DesignWsGroup
    Dim ws As Worksheet
    Dim rCol As Long
    Dim fRow As Long, lRow As Long, lRowSameGroup As Long
    Dim key As String
    Dim value As String
    Dim designGroup As DesignWsGroup
    Dim ret As Integer
    Dim i As Long, j As Long
    
    Set ws = refCell.Parent
    rCol = refCell.column
    
    ' Create a new DesignWsIO object using refCell's value as name
    Set designGroup = New DesignWsGroup
    designGroup.Initialize refCell.value
    
    'Save data in a design worksheet group
    fRow = refRow + 1
    lRow = wsInteract.FindLastNonEmptyCell(ws.Columns(rCol)).row
    
    Dim wsName As String, designIOtag As String, summaryTag As String
    Dim tagsRelation As TagsRelationship
    For i = fRow To lRow
    
        If i = lRow Then
            lRowSameGroup = wsInteract.FindLastRow(i, rCol + 2, ws)
        Else
            lRowSameGroup = wsInteract.FindNextNonEmptyCell(ws.Cells(i, rCol)).row - 1
        End If
        
        Set tagsRelation = New TagsRelationship
        
        wsName = ws.Cells(i, rCol)
        g_log.WriteLog "      Reading worksheet '" & wsName & "' data"
        tagsRelation.Name = wsName
        tagsRelation.isPrint = ws.Cells(i, rCol + 1)
        
        'Save tag relation data of a worksheet in a design worksheet group
        For j = i To lRowSameGroup
            If VarType(ws.Cells(j, rCol + 2)) = vbError Then Exit Function
                
            designIOtag = ws.Cells(j, rCol + 2)
            summaryTag = ws.Cells(j, rCol + 3)
            ret = tagsRelation.Add(designIOtag, summaryTag)
            
        Next j
        designGroup.AddTagsRelationship tagsRelation
        i = lRowSameGroup
    Next i
    
    Set ReadSingleDesignGroupData = designGroup
End Function


Private Function TransferDataFromSummary() As Integer
    Dim i As Long, j As Long
    Dim targetWsTag As SummaryWsColTags, targetWs As DesignWsIO, wsName As String
    Dim designGroup As DesignWsGroup, tagsRelation As TagsRelationship
    Dim ret As Long, key As Variant
    For i = 1 To df_sum.CountRows
        If Not activeRows(i) Then GoTo NextIteration
        g_log.WriteLog "Current Row: " & i + fRow_summary
        
        For Each key In targetWsTags
            Set targetWsTag = targetWsTags(key)
            wsName = df_sum.idata(i, targetWsTag.colNumInDF)
            'Check if worksheets exists in WSData
            If wsName = vbNullString Or wsName = "-" Then GoTo NextWorksheet
            countTotal = countTotal + 1
            
            If (Not designWorksheetsIO.Exists(wsName)) And (Not designWsGroups.Exists(wsName)) Then
                g_log.WriteLog "  Cannot Find Design Worksheet: '" & wsName & "' Input/Output Data in 'WSData' Sheet."
                GoTo NextWorksheet
            End If
            
            'transfer data to worksheet
            If designWorksheetsIO.Exists(wsName) Then
                g_log.WriteLog "  Design Worksheet: '" & wsName & "'."
                Set targetWs = designWorksheetsIO(wsName)
                ret = TransferSingleDesignWorksheetData(targetWs, i)
                If ret = 0 Then countSucessful = countSucessful + 1
            Else
                Debug.Assert Not wsName = "FinPlatesBolts"
                g_log.WriteLog "  Design Worksheets Group: '" & wsName & "'."
                Set designGroup = designWsGroups(wsName)
                For Each tagsRelation In designGroup.TagsRelationships
                    g_log.WriteLog "    Design Worksheets: '" & tagsRelation.Name & "'."
                    'Debug.Assert Not (tagsRelation.Name = "Calculation Title" And wsName = "FinPlatesBolts")
                    Debug.Assert Not wsName = "FinPlatesBolts"
                    Set targetWs = designWorksheetsIO(tagsRelation.Name)
                    ret = TransferSingleDesignWorksheetData(targetWs, i, tagsRelation)
                Next
                If ret = 0 Then countSucessful = countSucessful + 1
            End If

NextWorksheet:
        Next key
NextIteration:
    Next i
End Function

Private Function TransferSingleDesignWorksheetData(targetWsIO As DesignWsIO, dfRowNum As Long, _
                                                Optional tagsRelation As TagsRelationship) As Integer
    Dim WsInputs As Object, WsOutputs As Object
    Dim key As Variant
    Dim targetWs As Worksheet
    Dim tag_Ws As String, tag_Summary As String
    Dim val As Variant
    Set WsInputs = targetWsIO.WsInput
    Set WsOutputs = targetWsIO.WsOutput
    
    Set targetWs = wsInteract.setWorksheet(targetWsIO.Name, isSkipErr:=True)
    If targetWs Is Nothing Then GoTo ExitFunc:
    
    For Each key In WsInputs.keys
        tag_Ws = key
        If tagsRelation Is Nothing Then
            tag_Summary = key
        Else
            tag_Summary = tagsRelation.tagRelation(key)
        End If
        If columnTags.Exists(tag_Summary) Then
            val = df_sum.idata(dfRowNum, columnTags(tag_Summary).colNumInDF)
            g_log.WriteLog "     Input Tag: '" & key & ":" & tag_Summary & "'; value: " & val

            targetWs.Range(WsInputs(tag_Ws)) = val
        Else
            If Not tag_Summary = "-" Then g_log.WriteLog "     Cannot find Input Tag in the summary sheet: '" & tag_Summary & "';"
        End If
    Next key
    
    targetWs.Calculate
    For Each key In WsOutputs.keys
        tag_Ws = key
        If tagsRelation Is Nothing Then
            tag_Summary = key
        Else
            tag_Summary = tagsRelation.tagRelation(key)
        End If
        If columnTags.Exists(tag_Summary) Then
            val = targetWs.Range(WsOutputs(tag_Ws))
            g_log.WriteLog "     Output Tag: '" & key & ":" & tag_Summary & "'; value: " & CStr(val)
            df_sum.idata(dfRowNum, columnTags(tag_Summary).colNumInDF) = val
            If Not activeOutputCols.Exists(tag_Summary) Then
                activeOutputCols.Add tag_Summary, columnTags(tag_Summary)
            End If
        Else
            If Not tag_Summary = "-" Then g_log.WriteLog "     Cannot find Output Tag in the summary sheet: '" & tag_Summary & "';"
        End If
    Next key
    
    Dim pdfName As String
    If isOutputPDF Then
        If Not tagsRelation Is Nothing Then
            If Not tagsRelation.isPrint Then
                Exit Function
            End If
        End If
        Dim ret2 As Integer
        strCount = VBA.Format(countPDF, "00")
        If isIncludePrefix Then
            pdfName = targetWs.Name & "-" & strCount
        Else
            pdfName = strCount
        End If
        g_log.WriteLog "     Try printing PDF: '" & pdfName
        ret2 = Create_PDF(targetWs, myFilePath, pdfName)
        If ret2 = 0 Then
            g_log.WriteLog "     Print PDF successful."
        Else
            g_log.RaiseWarning "      Could not create PDF file! Record Skipped. Please check file name", failToCreatePDF
        End If
        countPDF = countPDF + 1
    End If
    
    Exit Function
ExitFunc:
    TransferSingleDesignWorksheetData = -1
    g_log.RaiseWarning "   No Worksheet '" & targetWsIO.Name & "' is Found in the workbook. Skip this worksheet.", NoWorksheetFoundInWorkbook
End Function

Private Function TransferDesignGroupData(designGroup As DesignWsGroup) As Integer
    Dim tagsRelation As TagsRelationship
    
    For Each tagsRelation In designGroup.TagsRelationships
        
    Next
    
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
Private Function Create_PDF(myWS As Worksheet, myFilePath As String, myFileName As String) As Integer
    
    Dim strFile As String
    strFile = myFilePath & "\" & myFileName

        myWS.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            fileName:=strFile, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
    
    Exit Function
    
ErrHandler:
    
    Create_PDF = -1
End Function


