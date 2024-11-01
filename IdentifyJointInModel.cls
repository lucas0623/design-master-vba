VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IdentifyJointInModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Operation.ConnAnalysis")
Private mCompleteMessageAddition As String
Private mTerminateMessageMain As String, mTerminateMessageAddition As String

Private mModel As StrModel
Private mDsSys As DataSheetSystem


Public Function Main(Optional isShowMsgBox As Boolean = True) As Integer
    Dim ret As Integer
    
    ret = CheckDataBeforeStart
    If Not ret = 0 Then
        GoTo ExitFunction
    End If
    
    g_log.WriteLog "Model Connection Identification Started."
    FormStrObj
    IdentifyConn
    WriteData
    mDsSys.prop("isCreated", "isIdentifiedConn") = True
    g_log.WriteLog "Model Connection Identification Completed."
ExitFunction:
    Main = ret
End Function

Private Sub Class_Initialize()
    Set mModel = New StrModel
    Set mDsSys = New DataSheetSystem
End Sub

Private Function CheckDataBeforeStart() As Integer
    Application.Calculate
    If Not mDsSys.prop("WhatOperationCanProcess", "IdentifyRestraint") Then
        mTerminateMessageMain = "Not Sufficient Data to carry out the process!!"
        CheckDataBeforeStart = -1
    End If
End Function

Private Sub FormStrObj()
    Dim ret As Integer
    ret = mModel.Constructor.FormJointObj
    ret = mModel.Constructor.FormFrmObj
    ret = mModel.Constructor.FormMemberObj
End Sub

Private Sub IdentifyConn()
'Criteria for connection
'1. a joint with restraint
'2. a joint with more than 2 frames connecting to it.
'3. 2 frames -> different sections
'4. 2 frames -> 2 different members
'Splice joint with the same section cannot be identified by this sub

    Dim i As Long
    Dim jt As Variant, jts As Collection
    Dim frms As Collection
    
    Set jts = mModel.joints
    For Each jt In jts
        Set frms = jt.connectedFrames
        If frms Is Nothing Then
            g_log.WriteLog "Joint ID: " & jt.name & "; Nos. of Connected Frames = 0; isConnection = False"
            GoTo NextIteration
        End If
        If jt.isRestraint Then
            jt.isConn = True
        ElseIf frms.count > 2 Then
            jt.isConn = True
        ElseIf frms.count = 2 Then
            If Not (frms(1).section = frms(2).section) Or Not (frms(1).memberName = frms(2).memberName) Then
                jt.isConn = True
            End If
        End If
        
        'save restraint result to members properties
        g_log.WriteLog "Joint ID: " & jt.name & "; Nos. of Connected Frames = " & frms.count & "; isConnection = " & jt.isConn
NextIteration:
    Next

End Sub

Private Sub WriteData()
    Dim dsConn As oDataSheet
    Dim dsManager As DataSheetManager
    Set dsManager = New DataSheetManager
    Set dsConn = dsManager.DSJointConnectivity
    Dim df As clsDataFrame
    
    Set df = mModel.GetDataframe(obj_jt, "name", "connectedMembersStr", "connectedFramesStr", "connectedFramesSectionStr", "isRestraint", "isConn")
    dsConn.ClearData
    
    g_log.WriteLog "Writing result to Worksheets...."
    With dsConn.tagSelector
        dsConn.WriteDataframe df, True, False, .jtID, .memID, .eleID, .sections, .isRestraint, .isConn
    End With
End Sub

Public Property Get CompleteMessageAddition() As String
    CompleteMessageAddition = mCompleteMessageAddition
End Property

Public Property Get TerminateMessageMain() As String
    TerminateMessageMain = mTerminateMessageMain
End Property

Public Property Get TerminateMessageAddition() As String
    TerminateMessageAddition = mTerminateMessageAddition
End Property

