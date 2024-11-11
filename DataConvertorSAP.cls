VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataConvertorSAP"
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
Private model As New StrModel
Private ds_sys As DataSheetSystem

'data frame for temporarily use
Private df_SAPJtCoor As clsDataFrame, df_SAPJtRestrAssign As clsDataFrame
Private df_SAPFrameConn As clsDataFrame, df_SAPFrameSection As clsDataFrame, df_SAPFrameLocalAsix As clsDataFrame
Private df_SAPFrameForce As clsDataFrame
Private df_SAPLinkSection As clsDataFrame, df_SAPLinkForce As clsDataFrame, df_SAPLinkConn As clsDataFrame

'Data Frame for return
Private df_ele As clsDataFrame, df_force As clsDataFrame
Private df_joint As clsDataFrame

Private errMsg As String
Private isTerminate As Boolean


Private Sub Class_Initialize()
    Set wsInteract = New clsWorksheetsInteraction
    Set genFunc = New clsGeneralFunctions
    Set UI = New clsUIManager
    Set ds_sys = New DataSheetSystem
    
End Sub

Private Function IDataFormatConvertor_GetUserInput() As String
    IDataFormatConvertor_GetUserInput = UI.GetFilePath(".xlsx", "Open SAP .xlsx output file", ds_sys.prop("ImportLog", "folderPath"))
End Function

Public Function IDataFormatConvertor_ReadData(filePath As String) As Integer
    Dim ret As Integer
    
    ret = ReadSAPTables(filePath)
    If ret = -1 Then GoTo TerminateFunc
    
    WriteStrModelToDataFrame
    
    Exit Function
    
TerminateFunc:

    IDataFormatConvertor_ReadData = ret
'    If Not errMsg = vbNullString Then
'        MsgBox errMsg
'    End If
End Function

Private Function WriteStrModelToDataFrame() As Integer

    If Not df_SAPJtCoor Is Nothing Then Set df_joint = model.GetDataframe(obj_jt, "name", "x", "y", "z", "isRestraint")
    Set df_ele = model.GetDataframe(obj_frm, "Name", "section", "jtIName", "jtJName", "Length", "EleTypeStr", "localAxis")
    Dim df As clsDataFrame
    
    If Not df_SAPFrameForce Is Nothing Then Set df = df_SAPFrameForce
    If Not df Is Nothing And Not df_SAPLinkForce Is Nothing Then df.AddDataFrame df_SAPLinkForce
    If Not df Is Nothing Then Set df_force = df
End Function

Private Function ReadSAPTables(filePath As String) As Integer
    Dim ret As Integer
    Dim SapWB As Workbook, rng As Range

    'Validation of Data
    If filePath = vbNullString Then
        ret = -1
        GoTo TerminateFunc
    End If

    'Open File
    Workbooks.OpenText fileName:= _
        filePath, DataType:=xlDelimited, Semicolon:=True, Local:=True
    ds_sys.prop("ImportLog", "folderPath") = ActiveWorkbook.path & "\"
    
    Set SapWB = ActiveWorkbook
    
    ret = ReadJointCoorTable
    If ret = -1 Then
        errMsg = "'Joint Coordinate' table missing! Please check the input file"
        GoTo TerminateFunc
    Else
        FormJointObj
    End If
    
    ret = ReadFrameConnectivityTable
    If ret = -1 Then
        errMsg = "'Connectivity - Frame' table missing! Please check the input file"
        GoTo TerminateFunc
    Else
        FormFrameObj
    End If
    
    ret = ReadFrameSectionTable
    If ret = -1 Then
        errMsg = "'Frame Section Assignments' table missing! Please check the input file"
        GoTo TerminateFunc
    Else
        AssignFrameSection
    End If
    
    ret = ReadFrameLocalAxisTable
    If ret = -1 Then
        'errMsg = "'Frame Local Axis Assignment' table missing! Please check the input file"
        'GoTo TerminateFunc
    Else
        AssignFrameLocalAxis
    End If
    
    ret = ReadFrameForceTable
    If ret = -1 Then
        errMsg = "'Frame Force' table missing! Please check the input file"
        GoTo TerminateFunc
    End If
    
    ret = ReadLinkConnectivityTable
    If ret = -1 Then
        
    Else
        FormLinkObj
    End If
    
    ret = ReadLinkSectionTable
    If ret = -1 Then
        
    Else
        AssignLinkSection
    End If
    
    ret = ReadLinkForceTable
    If ret = -1 Then
        
    End If
    
    ret = ReadJointRestrAssignTable
    If ret = -1 Then
        errMsg = "'Joint Restraint Assignment' table missing! Please check the input file"
        GoTo TerminateFunc
    Else
        AssignJointRestraint
    End If
    
    If Not SapWB Is Nothing Then SapWB.Close False
    Exit Function
    
TerminateFunc:
    If Not SapWB Is Nothing Then SapWB.Close False
    ReadSAPTables = ret
    If Not errMsg = vbNullString Then
        MsgBox errMsg
    End If
End Function

Private Function ReadJointCoorTable() As Integer
    Dim ws As Worksheet, ret As Integer
    
    Set ws = setWorksheet("Joints Coordinate", isSkipErr:=True)
    If ws Is Nothing Then Set ws = setWorksheet("Joint Coordinates", isSkipErr:=True)
    If ws Is Nothing Then
        ret = -1
        GoTo ExitFunc:
    End If
    
    Set df_SAPJtCoor = New clsDataFrame
    df_SAPJtCoor.Init_ReadWorksheet ws, row_tag:=2, rRow:=3, headTags:=genFunc.CStr_arr(Array("Joint", "XorR", "Y", "Z"))
    If df_SAPJtCoor Is Nothing Then
        ret = -1
        g_log.WriteLog "Cannot identify some of the data inside the worksheet."
        GoTo ExitFunc:
    End If
    g_log.WriteLog "'" & ws.name & " ' Table is successfully read."
ExitFunc:
    ReadJointCoorTable = ret
End Function

Private Function FormJointObj() As Integer
    Dim i As Long, jt As StrJoint, ret As Integer
    Dim df As clsDataFrame
    Set df = df_SAPJtCoor
    For i = 1 To df.CountRows
        Set jt = New StrJoint
        jt.Init df.idata(i, 1), df.idata(i, 2), df.idata(i, 3), df.idata(i, 4)
        ret = model.AddStrObjToColl(jt, obj_jt)
    Next i
End Function

Private Function ReadJointRestrAssignTable() As Integer
    Dim ws As Worksheet, ret As Integer
    
    Set ws = setWorksheet("Joint Restraint Assignments", isSkipErr:=True)
    If ws Is Nothing Then
        ret = -1
        GoTo ExitFunc:
    End If
    
    Set df_SAPJtRestrAssign = New clsDataFrame
    df_SAPJtRestrAssign.Init_ReadWorksheet ws, row_tag:=2, rRow:=3, headTags:=genFunc.CStr_arr(Array("Joint"))
    If df_SAPJtRestrAssign Is Nothing Then
        ret = -1
        g_log.WriteLog "Cannot identify some of the data inside the worksheet."
        GoTo ExitFunc:
    End If
    g_log.WriteLog "'" & ws.name & " ' Table is successfully read."
ExitFunc:
    ReadJointRestrAssignTable = ret
End Function

Private Function AssignJointRestraint() As Integer
    Dim i As Long, jt As StrJoint
    Dim df As clsDataFrame
    Set df = df_SAPJtRestrAssign
    If Not df.isInitialized Then Exit Function
    For i = 1 To df.CountRows
        Set jt = model.GetStrObject(df.idata(i, 1), obj_jt)
        jt.isRestraint = True
    Next i
End Function

Private Function ReadFrameConnectivityTable() As Integer
    Dim ws As Worksheet, ret As Integer
    
    Set ws = setWorksheet("Connectivity - Frame", isSkipErr:=True)
    If ws Is Nothing Then
        ret = -1
        GoTo ExitFunc:
    End If
    
    Set df_SAPFrameConn = New clsDataFrame
    df_SAPFrameConn.Init_ReadWorksheet ws, row_tag:=2, rRow:=3, headTags:=genFunc.CStr_arr(Array("Frame", "JointI", "JointJ", "Length"))
    If df_SAPFrameConn Is Nothing Then
        ret = -1
        g_log.WriteLog "Cannot identify some of the data inside the worksheet."
        GoTo ExitFunc:
    End If
    g_log.WriteLog "'" & ws.name & " ' Table is successfully read."
ExitFunc:
    ReadFrameConnectivityTable = ret
End Function

Private Function FormFrameObj() As Integer
    Dim i As Long, frm As StrFrame, ret As Integer
    Dim jtI As StrJoint, jtJ As StrJoint
    Dim df As clsDataFrame
    Set df = df_SAPFrameConn
    For i = 1 To df.CountRows
        Set jtI = SetJoint(model, df.idata(i, 2))
        Set jtJ = SetJoint(model, df.idata(i, 3))
        Set frm = New StrFrame
        frm.Init df.idata(i, 1), jtI:=jtI, jtJ:=jtJ, Length:=df.idata(i, 4)
        ret = model.AddStrObjToColl(frm, obj_frm)
    Next i
End Function

Private Function FormLinkObj() As Integer
    Dim i As Long, frm As StrFrame, ret As Integer
    Dim jtI As StrJoint, jtJ As StrJoint
    Dim df As clsDataFrame
    Set df = df_SAPLinkConn
    For i = 1 To df.CountRows
        Set jtI = SetJoint(model, df.idata(i, 2))
        Set jtJ = SetJoint(model, df.idata(i, 3))
        Set frm = New StrFrame
        frm.Init df.idata(i, 1), jtI:=jtI, jtJ:=jtJ, Length:=df.idata(i, 4)
        ret = model.AddStrObjToColl(frm, obj_frm)
    Next i
End Function

Private Function ReadLinkConnectivityTable() As Integer
    Dim ws As Worksheet, ret As Integer
    Dim i As Long
    Set ws = setWorksheet("Connectivity - Link", isSkipErr:=True)
    If ws Is Nothing Then
        ret = -1
        GoTo ExitFunc:
    End If
    
    Set df_SAPLinkConn = New clsDataFrame
    df_SAPLinkConn.Init_ReadWorksheet ws, row_tag:=2, rRow:=3, headTags:=genFunc.CStr_arr(Array("Link", "JointI", "JointJ", "Length"))
    Set df_SAPLinkConn = ReviseLinkName(df_SAPLinkConn)
    If df_SAPLinkConn Is Nothing Then
        ret = -1
        g_log.WriteLog "Cannot identify some of the data inside the worksheet."
        GoTo ExitFunc:
    End If
    g_log.WriteLog "'" & ws.name & " ' Table is successfully read."
ExitFunc:
    ReadLinkConnectivityTable = ret
End Function

Private Function ReadFrameSectionTable() As Integer
    Dim ws As Worksheet, ret As Integer
    
    Set ws = setWorksheet("Frame Section Assignments", isSkipErr:=True)
    If ws Is Nothing Then
        ret = -1
        GoTo ExitFunc:
    End If
    
    Set df_SAPFrameSection = New clsDataFrame
    df_SAPFrameSection.Init_ReadWorksheet ws, row_tag:=2, rRow:=3, headTags:=genFunc.CStr_arr(Array("Frame", "AnalSect"))
    If df_SAPFrameSection Is Nothing Then
        ret = -1
        g_log.WriteLog "Cannot identify some of the data inside the worksheet."
        GoTo ExitFunc:
    End If
    g_log.WriteLog "'" & ws.name & " ' Table is successfully read."
ExitFunc:
    ReadFrameSectionTable = ret
End Function

Private Function ReadLinkSectionTable() As Integer
    Dim ws As Worksheet, ret As Integer
    
    Set ws = setWorksheet("Link Property Assignments", isSkipErr:=True)
    If ws Is Nothing Then
        ret = -1
        GoTo ExitFunc:
    End If
    
    Set df_SAPLinkSection = New clsDataFrame
    df_SAPLinkSection.Init_ReadWorksheet ws, row_tag:=2, rRow:=3, headTags:=genFunc.CStr_arr(Array("Link", "LinkProp"))
    Set df_SAPLinkSection = ReviseLinkName(df_SAPLinkSection)
    If df_SAPLinkSection Is Nothing Then
        ret = -1
        g_log.WriteLog "Cannot identify some of the data inside the worksheet."
        GoTo ExitFunc:
    End If
    g_log.WriteLog "'" & ws.name & " ' Table is successfully read."
ExitFunc:
    ReadLinkSectionTable = ret
End Function

Private Function AssignFrameSection() As Integer
    Dim i As Long, frm As StrFrame
    Dim df As clsDataFrame
    Set df = df_SAPFrameSection
    For i = 1 To df.CountRows
        Set frm = model.GetStrObject(df.idata(i, 1), obj_frm)
        If frm Is Nothing Then
            MsgBox "Unexpected Error! Cannot find frame " & df.idata(i, 1) & ". Please check your data file."
            End
        End If
        frm.section = df.idata(i, 2)
    Next i
End Function

Private Function AssignLinkSection() As Integer
    Dim i As Long, frm As StrFrame
    Dim df As clsDataFrame
    Set df = df_SAPLinkSection
    For i = 1 To df.CountRows
        Set frm = model.GetStrObject(df.idata(i, 1), obj_frm)
        If frm Is Nothing Then
            MsgBox "Unexpected Error! Cannot find frame " & df.idata(i, 1) & ". Please check your data file."
            End
        End If
        frm.section = df.idata(i, 2)
    Next i
End Function

Private Function ReadFrameLocalAxisTable() As Integer
    Dim ws As Worksheet, ret As Integer
    
    Set ws = setWorksheet("Frame Local Axes 1 - Typical", isSkipErr:=True)
    If ws Is Nothing Then
        ret = -1
        GoTo ExitFunc:
    End If
    
    Set df_SAPFrameLocalAsix = New clsDataFrame
    df_SAPFrameLocalAsix.Init_ReadWorksheet ws, row_tag:=2, rRow:=3, headTags:=genFunc.CStr_arr(Array("Frame", "Angle"))
    If df_SAPFrameLocalAsix Is Nothing Then
        ret = -1
        g_log.WriteLog "Cannot identify some of the data inside the worksheet."
        GoTo ExitFunc:
    End If
    g_log.WriteLog "'" & ws.name & " ' Table is successfully read."
ExitFunc:
    ReadFrameLocalAxisTable = ret
End Function

Private Function AssignFrameLocalAxis() As Integer
    Dim i As Long, frm As StrFrame
    Dim df As clsDataFrame
    Set df = df_SAPFrameLocalAsix
    If Not df.isInitialized Then Exit Function
    For i = 1 To df.CountRows
        Set frm = model.GetStrObject(df.idata(i, 1), obj_frm)
        frm.localAxis = df.idata(i, 2)
    Next i
End Function

Private Function ReadFrameForceTable() As Integer
    Dim ws As Worksheet, ret As Integer
    Dim lRow_SAPforce As Long, rng As Range
    
    Const fRow_SAP As Integer = 4
    Const col_stepDefault As Integer = 5
    
    Set ws = setWorksheet("Element Forces - Frames", isSkipErr:=True)
    If ws Is Nothing Then
        ret = -1
        GoTo ExitFunc:
    End If
    
    With ws
        If .FilterMode Then
            .ShowAllData
        End If
        If .Cells(fRow_SAP - 2, col_stepDefault) <> "StepType" Then
            .Columns(col_stepDefault).Insert
            .Cells(fRow_SAP, col_stepDefault) = "N/A"
            .Cells(fRow_SAP - 2, col_stepDefault) = "StepType"
            lRow_SAPforce = wsInteract.FindLastRow(fRow_SAP, 1)
            Set rng = Range(Cells(fRow_SAP, col_stepDefault), Cells(lRow_SAPforce, col_stepDefault))
            .Cells(fRow_SAP, col_stepDefault).AutoFill Destination:=rng
        End If
    End With
        
    Set df_SAPFrameForce = New clsDataFrame
    df_SAPFrameForce.Init_ReadWorksheet ws, row_tag:=2, rRow:=3, headTags:=genFunc.CStr_arr(Array("Frame", "Station", "OutputCase", "StepType", "P", "V2", "V3", "T", "M2", "M3", "FrameElem"))
    If df_SAPFrameForce Is Nothing Then
        ret = -1
        g_log.WriteLog "Cannot identify some of the data inside the worksheet."
        GoTo ExitFunc:
    End If
    g_log.WriteLog "'" & ws.name & " ' Table is successfully read."
ExitFunc:
    ReadFrameForceTable = ret
End Function

Private Function ReadLinkForceTable() As Integer
    Dim ws As Worksheet, ret As Integer
    Dim lRow_SAPforce As Long, rng As Range
    
    Const fRow_SAP As Integer = 4
    Const col_stepDefault As Integer = 6
    
    Set ws = setWorksheet("Element Forces - Links", isSkipErr:=True)
    If ws Is Nothing Then
        ret = -1
        GoTo ExitFunc:
    End If
    
    'Modify with 'StepType'
    With ws
        If .FilterMode Then
            .ShowAllData
        End If
        If .Cells(fRow_SAP - 2, col_stepDefault) <> "StepType" Then
            .Columns(col_stepDefault).Insert
            .Cells(fRow_SAP, col_stepDefault) = "N/A"
            .Cells(fRow_SAP - 2, col_stepDefault) = "StepType"
            lRow_SAPforce = wsInteract.FindLastRow(fRow_SAP, 1)
            Set rng = Range(Cells(fRow_SAP, col_stepDefault), Cells(lRow_SAPforce, col_stepDefault))
            .Cells(fRow_SAP, col_stepDefault).AutoFill Destination:=rng
        End If
    End With
    
    Set df_SAPLinkForce = New clsDataFrame
    df_SAPLinkForce.Init_ReadWorksheet ws, row_tag:=2, rRow:=3, headTags:=genFunc.CStr_arr(Array("Link", "Station", "OutputCase", "StepType", "P", "V2", "V3", "T", "M2", "M3", "LinkElem"))
    If df_SAPLinkForce Is Nothing Then
        ret = -1
        g_log.WriteLog "Cannot identify some of the data inside the worksheet."
        GoTo ExitFunc:
    End If
    'convert the I-end and J-end to absolute value
    Set df_SAPLinkForce = ReviseLinkName(df_SAPLinkForce)
    Dim link As StrFrame, i As Long
    Dim linkName As String
    For i = 1 To df_SAPLinkForce.CountRows
        If df_SAPLinkForce.idata(i, 2) = "I-End" Then
            df_SAPLinkForce.idata(i, 2) = 0
        Else
            linkName = df_SAPLinkForce.idata(i, 1)
            Set link = model.GetStrObject(linkName, obj_frm)
            df_SAPLinkForce.idata(i, 2) = link.Length
        End If
    Next i

    g_log.WriteLog "'" & ws.name & " ' Table is successfully read."
ExitFunc:
    ReadLinkForceTable = ret
End Function

Private Function ReviseLinkName(df As clsDataFrame) As clsDataFrame
    Dim reviseDF As New clsDataFrame
    Dim i As Long
    Set reviseDF = df
    For i = 1 To df.CountRows
        reviseDF.idata(i, 1) = "Link-" & CStr(df.idata(i, 1))
    Next i
    Set ReviseLinkName = reviseDF
End Function


Private Function setWorksheet(wsName As String, isSkipErr As Boolean) As Worksheet
    g_log.WriteLog "Try to Read '" & wsName & " ' Table"
    Dim ws As Worksheet
    Set ws = wsInteract.setWorksheet(wsName, isSkipErr:=isSkipErr)
    If ws Is Nothing Then
        g_log.WriteLog "Cannot Find '" & wsName & "' Table"
    End If
    
    Set setWorksheet = ws
End Function

Private Function SetJoint(model As StrModel, jtName As String) As StrJoint
    Dim jt As StrJoint
    Dim ret As Integer
    
    If model.IsStrObjectExist(jtName, obj_jt) Then
        Set jt = model.GetStrObject(jtName, obj_jt)
    Else
        g_log.RaiseWarning "Joint " & jtName & " coordinates data missing. The result of 'Process Model Data>Create Member' might be affected due to missing data!.", missingJointCoor
        Set jt = New StrJoint
        jt.Init jtName
        ret = model.AddStrObjToColl(jt, obj_jt)
    End If
    Set SetJoint = jt
End Function

Public Property Get IDataFormatConvertor_DfEle() As clsDataFrame
    Set IDataFormatConvertor_DfEle = df_ele
End Property

Property Let IDataFormatConvertor_DfEle(value As clsDataFrame)
    Set df_ele = value
End Property

Public Property Get IDataFormatConvertor_DfForce() As clsDataFrame
    Set IDataFormatConvertor_DfForce = df_force
End Property

Public Property Let IDataFormatConvertor_DfForce(value As clsDataFrame)
    Set df_force = value
End Property

Public Property Get IDataFormatConvertor_DfJoint() As clsDataFrame
    Set IDataFormatConvertor_DfJoint = df_joint
End Property

Public Property Let IDataFormatConvertor_DfJoint(value As clsDataFrame)
    Set df_joint = value
End Property



