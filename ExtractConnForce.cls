VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtractConnForce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Operation.ExtractConnectionForce")

Option Explicit

Private mCompleteMessageAddition As String
Private mTerminateMessageMain As String, mTerminateMsgAddition As String

Private genFunc As New clsGeneralFunctions
Private UFControl As New VMForceExtraction
Private mDsSys As New DataSheetSystem
Private mDsDesignData As New DataSheetSummary
Private mModel As New StrModel

Private frmForces As Collection 'extracted StrFrameForceForConn
Private ws_sum As Worksheet

Private ExtractFrmForceMethod As IExtractFrmForceMethod
Private lc() As String, MemberNames() As String, sections() As String 'for use in Userform
Public Enum ConnForceExtractMethod
    EachNodeAllForces
    CorrespondeceCases
End Enum

Public Property Get CompleteMessageAddition() As String
    CompleteMessageAddition = mCompleteMessageAddition
End Property

Public Property Get TerminateMessageMain() As String
    TerminateMessageMain = mTerminateMessageMain
End Property


Public Function Main(method As ConnForceExtractMethod) As Integer

    Dim ret As Integer

    ret = CheckDataBeforeStart
    If Not ret = 0 Then
        MsgBox "no data is found in the workbook! The macro will be terminated."
        GoTo ExitFunc
    End If
    g_log.WriteLog "Force Extraction (Connection) Started."

    '1. Get Data From Userform
    ret = ShowUFAndGetUserInput(method)
    If Not ret = 0 Then GoTo ExitFunc
    mDsDesignData.Initialize ws_sum.name

    '2. Form StrModel Objects
    ret = FormStrObj 'StrObj is formed after the userform due to performance issue
    If Not ret = 0 Then GoTo ExitFunc

    '3. Extract The forces according to user preference. Return as a collection of frame force objects (1 obj = 1 row result)
    g_log.WriteLogWithTime "Start extracting frame forces according to user perference from the Structrual Model Objects..."
    Set frmForces = SelectExtractionMethod(method).ExtractForce
    g_log.WriteLogWithTime "Successfully formed the frame force collection. Total number of frame force object = " & frmForces.count, True

    '4. Transform the frame force object to data frame object
    WriteData
    
    mCompleteMessageAddition = mCompleteMessageAddition & Chr(10) & _
                        "Data is Written on Rows " & mDsDesignData.startRowWritten & " to " & mDsDesignData.endRowWritten
    Application.StatusBar = False
    Exit Function
ExitFunc:
    Main = -1
End Function

Private Function CheckDataBeforeStart() As Integer
    If Not mDsSys.prop("WhatOperationCanProcess", "ExtractConnectionForce") Then
        CheckDataBeforeStart = -1
    End If
End Function

Private Function GetUserInput() As Integer
    Dim ret As Integer
    With UFControl
        Set ws_sum = .wsSum
'        LC = .LC
        MemberNames = .Nodes
        sections = .connTypes
'        PosFilterType = .PosFilterType
'        posFilterVal = .PosFilterValue
'        IsBySec = .IsBySec
'        isExtremeCase = .isExtremeCase
'        IsEachLC = .IsEachLC
'        If isExtremeCase Then ExtremeCase = .ExtremeCase
'        MaxMinEnv = .MaxMinEnv
    End With
End Function

Private Function FormStrObj() As Integer
    'Form Str Obj with sections, members and load comb filter applied
    Dim ret As Integer
    Application.StatusBar = "Start forming Structural Model Objects according to user preference...."
    g_log.WriteLogWithTime "Start forming Structural Model Objects according to user preference...."

    With mModel.Constructor
        ret = .FormJointObj

        ret = .FormFrmObj
        If Not ret = 0 Then GoTo ExitFunction

        ret = .FormFrmForceObj
        If Not ret = 0 Then GoTo ExitFunction

        ret = .FormMemberObj
        If Not ret = 0 Then GoTo ExitFunction

        ret = .FormConnTypeObjForForceExtract
        If Not ret = 0 Then GoTo ExitFunction

    End With
    g_log.WriteLogWithTime "Structural Model Objects are formed successfully.", True
    Exit Function
ExitFunction:
    FormStrObj = ret
End Function

Private Function WriteData() As Integer
    Application.StatusBar = "Transforming data for output...."
    g_log.WriteLog "Transforming data for output...."
    Dim df As clsDataFrame
    Set df = mModel.GetDataframe_fromColl(frmForces, "nodeName", "connTypeName", "frameSection", "frameName", "frameLength", _
                "frameJtIName", "frameJtJName", "memberName", "memberIFrameName", "memberJFrameName", _
                "preFrameName", "nextFrameName", "memberTotalLength", "pos_fromMemJtI_percent", "pos_fromMemJtJ_percent", _
                "pos_fromMemJtI", "pos_fromMemJtJ", "pos_fromEleJtI", _
                "pos_fromEleJtJ", "pos_fromEleJtI_percent", "pos_fromEleJtJ_percent", "loadcomb", "extremeCaseType", "stepType", _
                "P", "V2", "V3", "T", "M2", "M3")

    '5. Write the df to the sheet
    g_log.WriteLogWithTime "Writing extracted data to the worksheet '" & mDsDesignData.name & "'. Total number of data = " & df.CountRows
    With mDsDesignData
        .WriteDataframe df, .nodeName, .connType, .section, .eleName, .eleLen, .jtI, .jtJ, .memName, .fFrm, .lFrm, _
                .pFrm, .nFrm, .memTotalLen, .pos_fromMemJtI_percent, .pos_fromMemJtJ_percent, _
                .pos_fromMemJtI, .pos_fromMemJtJ, .pos_fromEleJtI, .pos_fromEleJtJ, _
                .pos_fromEleJtI_percent, .pos_fromEleJtJ_percent, .loadComb, .caseName, .stepType, .p, .V2, _
                .V3, .t, .M2, .M3
    End With
    g_log.WriteLogWithTime "All data written to the worksheet."
    
    mCompleteMessageAddition = "Total number of data extracted = " & df.CountRows & "."
    
End Function
Private Function SelectUF(method As ConnForceExtractMethod) As IView
    Dim uf As IView

    Select Case method
        Case EachNodeAllForces
            Set uf = New UFExtractConnForce
            uf.Initialize UFControl
            UFControl.Initialize "ExtractConnForceMethodAllNode"
        Case CorrespondeceCases
            Set uf = New UFExtractConnForceByCo
            uf.Initialize UFControl
            UFControl.Initialize "ExtractConnForceMethodCo"
    End Select

    Set SelectUF = uf
End Function

Private Function ShowUFAndGetUserInput(method As FrameForceExtractMethod) As Integer
    Dim uf As IView, ret As Integer
    
    Set uf = SelectUF(method)
    uf.Show
    
    If uf.CloseState = 0 Then
        ret = GetUserInput
        g_log.WriteLogWithTime "The Userform is closed and selected preferences are saved.", True
    Else
        ret = -1
        g_log.WriteLogWithTime "The Userform is closed and the procedure will be terminated."
        GoTo ExitFunction
    End If
    
ExitFunction:
    ShowUFAndGetUserInput = ret
End Function
Private Function SelectExtractionMethod(method As ConnForceExtractMethod) As IExtractFrmForceMethod
    Dim extractMethod As IExtractFrmForceMethod
    Select Case method
        Case EachNodeAllForces: Set extractMethod = New ExtractConnForceMethodAllNode
        Case CorrespondeceCases: Set extractMethod = New ExtractConnForceMethodCo
    End Select
    extractMethod.Initialize mModel, UFControl
    Set SelectExtractionMethod = extractMethod
End Function

