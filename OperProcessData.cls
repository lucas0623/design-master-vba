VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OperProcessData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Operation.ConnAnalysis")
Option Explicit

'Private mCompleteMsg As String
'Private genFunc As New clsGeneralFunctions
Private mDsSys As New DataSheetSystem
'Private mDsDesignData As New DataSheetSummary
'Private mModel As New StrModel
Private mCompleteMessageAddition As String
Private mTerminateMessageMain As String, mTerminateMessageAddition As String

Private UFControl As New VMProcessData
Private uf As IView
Private isCreateMember As Boolean, isIdentifyRestraint As Boolean
Private isMapBySection As Boolean, isMapByKinkAngle As Boolean
Private mCompleteMsg As String

Public Property Get CompleteMessageAddition() As String
    CompleteMessageAddition = mCompleteMessageAddition
End Property

Public Property Get TerminateMessageMain() As String
    TerminateMessageMain = mTerminateMessageMain
End Property

Public Property Get TerminateMessageAddition() As String
    TerminateMessageAddition = mTerminateMessageAddition
End Property

Public Function Main() As Integer
    Dim ret As Integer
    
    ret = CheckDataBeforeStart
    If Not ret = 0 Then
        mTerminateMessageMain = "no data is found in the workbook! The macro will be terminated."
        GoTo ExitFunction
    End If
    
    Set uf = New UFProcessData
    uf.Initialize UFControl
    UFControl.Initialize
    uf.Show
    
    If Not uf.CloseState = 0 Then GoTo ExitFunction
    
    isCreateMember = UFControl.isCreateMember
    isIdentifyRestraint = UFControl.isIdentifyRestraint
    isMapBySection = UFControl.isMapBySection
    isMapByKinkAngle = UFControl.isMapByKinkAngle
    
    
    If isCreateMember And Not isMapBySection And Not isMapByKinkAngle Then
        mTerminateMessageMain = "Create Member Setting: Please select at least 2 out of 3 Criteria. "
        ret = -1
        GoTo ExitFunction
    End If
    
    If isCreateMember Then
        Dim CreateMember As New IdentifyMember
        CreateMember.Initialize g_log
        If Not CreateMember.Main(isMapBySection, isMapByKinkAngle) = 0 Then
            mTerminateMessageMain = CreateMember.TerminateMessageMain
            GoTo ExitFunction
        End If
    End If
    
    If isIdentifyRestraint Then
        
        Dim IdentifyFrameOrientation As New IdentifyFrameOrientationType

        If Not IdentifyFrameOrientation.Main = 0 Then
            mTerminateMessageMain = IdentifyFrameOrientation.TerminateMessageMain
            GoTo ExitFunction
        End If
        
        Dim IdentifyRestraint As New IdentifyJointInModel
        If Not IdentifyRestraint.Main = 0 Then
            mTerminateMessageMain = IdentifyRestraint.TerminateMessageMain
            GoTo ExitFunction
        End If
    End If
    
    Exit Function
    
ExitFunction:
    Main = ret
End Function

Private Function CheckDataBeforeStart() As Integer
    If Not mDsSys.prop("WhatOperationCanProcess", "CreateMemberBySection") And _
        Not mDsSys.prop("WhatOperationCanProcess", "CreateMemerByKinkAngle") And _
        Not mDsSys.prop("WhatOperationCanProcess", "IdentifyRestraint") Then
        CheckDataBeforeStart = -1
    End If
End Function

Public Property Get completeMsg() As String
    completeMsg = mCompleteMsg
End Property
