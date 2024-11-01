VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataConvertorGSA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Operation.ImportData"

'********************************************************
'This module read GSA .csv file and write the data to the worksheet
'Arthor: Lucas LEUNG
'Update Log
'05 Jun 2023 - Initial Setup
'*******************************************************
Option Explicit
Implements IDataFormatConvertor

Private wsInteract As clsWorksheetsInteraction
Private genFunc As clsGeneralFunctions
Private UI As clsUIManager
Private ds_sys As DataSheetSystem

Private df_ele As clsDataFrame, df_force As clsDataFrame
Private df_joint As clsDataFrame

Private isTerminate As Boolean


Private Sub Class_Initialize()
    Set wsInteract = New clsWorksheetsInteraction
    Set genFunc = New clsGeneralFunctions
    Set UI = New clsUIManager
    Set ds_sys = New DataSheetSystem
End Sub

Private Function IDataFormatConvertor_GetUserInput() As String
    IDataFormatConvertor_GetUserInput = UI.GetFilePath(".csv", "Open GSA .csv output file", ds_sys.prop("ImportLog", "folderPath"))
End Function

Private Function IDataFormatConvertor_ReadData(filePath As String) As Integer
    'return -1 if 'cancel' the file picker dialog box
    Dim ret As Integer, i As Long
    Dim heads_ele As Variant, heads_force As Variant
    Dim ws_gsa As Worksheet, gsaWB As Workbook
    Dim rCol_table As Long
    'Validation of Data
    If filePath = vbNullString Then
        ret = -1
        GoTo TerminateFunc
    End If
    
    '1. Open csv File
     Workbooks.OpenText fileName:= _
        filePath, DataType:=xlDelimited, Semicolon:=True, Local:=True
    ds_sys.prop("ImportLog", "folderPath") = ActiveWorkbook.path & "\"
    
    '2. Initialize/ Format Setting
    'If Not mLog Is Nothing Then mLog.WriteLog ".csv file is opened."
    rCol_table = 1
    Set ws_gsa = ActiveSheet
    Set gsaWB = ActiveWorkbook
    
    '3. Read Data into dataframe
    Set df_ele = ReadGSATableToDF(ws_gsa, "START_TABLE Elements", True)
    If isTerminate Then
        ret = -1
        GoTo TerminateFunc
    End If
    Set df_force = ReadGSATableToDF(ws_gsa, "START_TABLE Beam and Spring Forces and Moments", True)
    Set df_joint = ReadGSATableToDF(ws_gsa, "START_TABLE Nodes", True)

    If Not df_joint Is Nothing Then Set df_joint = df_joint.Filter_byHeads(genFunc.CStr_arr(Array("Node", "x", "y", "z", "Restr.")))
    With df_joint
        For i = 1 To .CountRows
            If .idata(i, 5) = "free" Then
                .idata(i, 5) = False
            Else
                .idata(i, 5) = True
            End If
        Next i
    End With
    
    heads_ele = Array("Elem", "Property", "Topology", "Topology 2", "Length/ Area/Volume", "Type", "Orient. Angle")
    Set df_ele = df_ele.Filter_byHeads(genFunc.CStr_arr(heads_ele))
    'Set df_ele = df_ele.Filter_str("Type", "Beam")
    
    '4. Modify the force data frame to fit the format of this worksheet
    If Not df_force Is Nothing Then
'        myArr2 = WSInteract.SetArrayTwoDim(fRow_force, rCol_table, True, lRow_force - fRow_force + 1)
'        Set df_force = Factory.NewclsDataFrame(myArr2, True, False)
        If df_force.columnNum("Env") = -1 Then df_force.InsertEmptyColumn 4, "Env" 'insert Column if there is no 'Env' Column
        heads_force = Array("Elem", "Pos", "Case", "Env", "Fx", "Fz", "Fy", "Mxx", "Mzz", "Myy")
        Set df_force = df_force.Filter_byHeads(genFunc.CStr_arr(heads_force))
        
        'Modify the 'Pos' from % to True length
        Dim col_pos As Long, col_eleID As Long, col_length As Long
        Dim dict_eleLen As Object
        col_eleID = df_ele.columnNum("Elem")
        col_length = df_ele.columnNum("Length/ Area/Volume")
        col_pos = df_force.columnNum("Pos")
        Set dict_eleLen = CreateObject("Scripting.Dictionary")
        For i = 1 To df_ele.CountRows
            dict_eleLen.Add df_ele.idata(i, col_eleID), df_ele.idata(i, col_length)
        Next i
        For i = 1 To df_force.CountRows
            df_force.idata(i, col_pos) = dict_eleLen(df_force.idata(i, col_eleID)) * df_force.idata(i, col_pos)
        Next i
        
    End If
    gsaWB.Close False
    Exit Function
TerminateFunc:
    IDataFormatConvertor_ReadData = ret
    If Not gsaWB Is Nothing Then gsaWB.Close False
End Function

Private Function ReadGSATableToDF(ws As Worksheet, tableName As String, Optional isMustRead As Boolean = False) As clsDataFrame
     Dim rCol_table As Long, lRow As Long
     Dim fRow_table As Long, lRow_table As Long
     Dim myRng As Range
     Dim errMsg As String
     rCol_table = 1
     
     With ws
        lRow = .Range("A" & .rows.count).End(xlUp).row
        Set myRng = .Range(.Cells(1, rCol_table), .Cells(lRow, rCol_table)).Find(tableName)
        
        If myRng Is Nothing Then
            If isMustRead Then
                errMsg = "Cannot find the " & tableName & " Table in the file.  The Program will be terminated."
                isTerminate = True
            End If
            GoTo TerminateFunc
        End If
        fRow_table = myRng.row + 2
        If .Range(myRng, .Cells(lRow, rCol_table)).Find("Maxima") Is Nothing Then
            Set myRng = .Range(myRng, .Cells(lRow, rCol_table)).Find("END_TABLE")
        Else
            Set myRng = .Range(myRng, .Cells(lRow, rCol_table)).Find("Maxima")
        End If
        lRow_table = myRng.row - 1
    End With
    
    'There if there are empty cells within the table. (apply if 'fully populate' not tick). Only 'force' table applies.
    Dim rng As Variant
    For Each rng In ws.Range(ws.Cells(fRow_table, 1), ws.Cells(lRow_table, 1))
        If rng.Text = vbNullString Then
            errMsg = "Wrong Data Format! Please tick the 'Fully Populate Field' option in GSA during export."
            isTerminate = True
            GoTo TerminateFunc
        End If
    Next
    
    Dim myArr As Variant, df As clsDataFrame
    
    myArr = wsInteract.SetArrayTwoDim(fRow_table, rCol_table, True, lRow_table - fRow_table + 1)
    Set df = New clsDataFrame
    df.Init_byArr myArr, True, False
    Set ReadGSATableToDF = df
    Exit Function
TerminateFunc:
    If Not errMsg = vbNullString Then
        MsgBox errMsg
    End If
End Function


Private Property Get IDataFormatConvertor_DfEle() As clsDataFrame
    Set IDataFormatConvertor_DfEle = df_ele
End Property

Private Property Let IDataFormatConvertor_DfEle(value As clsDataFrame)
    Set df_ele = value
End Property

Private Property Get IDataFormatConvertor_DfForce() As clsDataFrame
    Set IDataFormatConvertor_DfForce = df_force
End Property

Private Property Let IDataFormatConvertor_DfForce(value As clsDataFrame)
    Set df_force = value
End Property

Private Property Get IDataFormatConvertor_DfJoint() As clsDataFrame
    Set IDataFormatConvertor_DfJoint = df_joint
End Property

Private Property Let IDataFormatConvertor_DfJoint(value As clsDataFrame)
    Set df_joint = value
End Property
