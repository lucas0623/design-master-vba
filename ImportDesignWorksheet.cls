VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImportDesignWorksheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Operation"

Private genFunc As New clsGeneralFunctions
Private wsInteract As New clsWorksheetsInteraction
Private UI As New clsUIManager
Private ListBoxSelected As msForms.listBox, ListBoxExcluded As msForms.listBox
Private wb_origin As Workbook, wb_target As Workbook
Private ws_WSData_ori As Worksheet, ws_WSData_tar As Worksheet

Sub Main()
    Dim filePath As String
    Dim wbName As String
    Dim ws_count As Long
    Dim wsList() As String, ws_selected() As String 'for import use
    Dim i As Long
    
    Dim ws_summary As New Collection, ws_designWS As New Collection 'for setting up the imported worksheet
    Dim ws_import As Worksheet
    Dim ret As Integer

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    'initialize
    Set wb_origin = ActiveWorkbook
    Set ws_WSData_ori = wsInteract.setWorksheet("WSData")
    Set ws_import = ActiveSheet

    'ask user input and open the worksheet
    filePath = UI.GetFilePath(".xlsx,.xls,.xlsm")
    If filePath = vbNullString Then Exit Sub
    'Set wb_target = genFunc.SetWorkbook(filePath)
    Set wb_target = Workbooks.Open(filePath)
    Set ws_WSData_tar = wsInteract.setWorksheet("WSData", isSkipErr:=True)
    wsList = GetWSList(wb_target)
    'arr = collectionToArray(ws_list)
    ret = CreateUserForm(wsList)
    If Not ret = 0 Then Exit Sub
    ws_selected = UI.GetListBoxArr(ListBoxSelected)
    
    'Copy the worksheets
    Application.DisplayAlerts = False
    For i = UBound(ws_selected) To 0 Step -1
        wb_target.Sheets(ws_selected(i)).Copy After:=ws_import
        If InStr(1, ActiveSheet.name, "summary", vbTextCompare) <> 0 Then
            ws_summary.Add (ActiveSheet.name)
        Else
            ws_designWS.Add (ActiveSheet.name)
            'Debug.Print ActiveSheet.Name
        End If
        'Debug.Print ActiveWorkbook.name & "," & ActiveSheet.name
        ReplaceString ActiveSheet, "https*[*]"
        ReplaceString ActiveSheet, "[*.xlsx]"
        ReplaceString ActiveSheet, "[*.xlsm]"
        ReplaceString ActiveSheet, "[*.xls]"
    Next i
    Application.DisplayAlerts = True
    'Add the macro button
'    For i = 1 To ws_summary.count
'        DeleteAllShapes Worksheets(ws_summary(i))
'        CopyShape Worksheets("summary"), Worksheets(ws_summary(i))
'    Next

    'copy all the WSData
    If Not ws_WSData_tar Is Nothing Then
        Dim rCol_tar As Long, rCol_ori As Long, lCol_tar As Long
        rCol_tar = 3
        lCol_tar = wsInteract.LastNonEmptyColumn(ws_WSData_tar, 1) + 3
        rCol_ori = wsInteract.LastNonEmptyColumn(ws_WSData_ori, 1) + 4
        If rCol_ori = 4 Then rCol_ori = 3
        With ws_WSData_tar
    
            .Range(.Columns(rCol_tar), .Columns(lCol_tar)).Copy _
                Destination:=ws_WSData_ori.Columns(rCol_ori)
    
        End With
    End If

    wb_target.Close False

    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
End Sub
'Private Sub DeleteAllShapes(myWS As Worksheet)
''PURPOSE: Remove All Shape Objects From The Active Worksheet
''SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
'
'Dim shp As Shape
'
'For Each shp In myWS.Shapes
'   shp.Delete
'Next shp
'
'End Sub
'
'Private Sub CopyShape(fromWS As Worksheet, toWS As Worksheet)
'    Dim shp1 As Shape
'    Dim posX As Double
'    posX = 0
'    For Each shp1 In fromWS.Shapes
'        'shp1.Copy
'        toWS.Paste
'        Set shp2 = toWS.Shapes(toWS.Shapes.count)
'        shp2.Top = toWS.Range("E2").Top
'        shp2.Left = toWS.Range("E2").Left + posX
'        posX = posX + shp1.Width + 5
'    Next shp1
'End Sub
Private Function GetWSList(WB As Workbook) As String()
    Dim numSheet As Integer, retArr() As String

    Dim count As Integer
    With WB
        numSheet = .Worksheets.count

        If ws_WSData_tar Is Nothing Then
            ReDim retArr(numSheet - 1)
        Else
            ReDim retArr(numSheet - 2)
        End If
        
        For i = 1 To numSheet
            If Not (ws_WSData_tar.name = .Worksheets(i).name) Then
                retArr(count) = .Worksheets(i).name
                count = count + 1
            End If
        Next i
    End With
    GetWSList = retArr
End Function
Private Function CreateUserForm(wsList() As String) As Integer
    'let user select the worksheets to be imported

    Dim mbtn1 As msForms.CommandButton, mbtn2 As msForms.CommandButton
    Dim selectedStr() As String
    Dim myUserForm As New UFBasic
    Dim cmdBtnEvent As New EventSelectionBoxMulti
    
    myUserForm.Initialize 350, True '3, genFuncs.CStr_arr(Split("General,Member Filter,Load Comb Filter", ",")), 400, False
    myUserForm.TitleBarCaption = "Select Worksheets for Import"
    
    myUserForm.AddSelectionBoxMulti wsList, ListBoxSelected, "SELECTED Sheets", height_LB:=200, _
                                    is_reListBox2:=True, reListBox2:=ListBoxExcluded, _
                                    title2:="NOT Import", _
                                    reCmdBtn1:=mbtn1, reCmdBtn2:=mbtn2
    cmdBtnEvent.Init myUserForm, ListBoxSelected, ListBoxExcluded, mbtn1, mbtn2

    
    myUserForm.Show
    If Not myUserForm.CloseState = 0 Then
        ret = -1
    End If
ExitFunction:
    CreateUserForm = ret
End Function
Private Function collectionToArray(c As Collection) As Variant()
    Dim a() As Variant: ReDim a(0 To c.count - 1)
    Dim i As Long
    For i = 1 To c.count
        a(i - 1) = c.item(i)
    Next
    collectionToArray = a
End Function


Private Sub ReplaceString(ws As Worksheet, str As String)
'
' Macro2 Macro
    ws.Cells.Replace What:=str, Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
End Sub


