VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IdentifyMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Operation.ConnAnalysis"


'********************************************************
'This module is the main function for frame connectivity analysis
'Input: All Str Model Obj
'Output: Form Member object, and then write to the worksheet
'Arthor: Lucas LEUNG
'Update Log

'04 Apr 2023 - Initial
'*******************************************************
Private mCompleteMessageAddition As String
Private mTerminateMessageMain As String, mTerminateMessageAddition As String

Private mLog As clsLog
Private mDS As New DataSheetManager, mDsSys As New DataSheetSystem
'Private ds_jt As oDataSheet
'Private ds_frm As oDataSheet
'Private ds_frmForce As oDataSheet
Private mModel As New StrModel
Private mDefaultAngle As Double 'the angle set for recgonizing continuous member (for byDir only)

Public Property Get CompleteMessageAddition() As String
    CompleteMessageAddition = mCompleteMessageAddition
End Property

Public Property Get TerminateMessageMain() As String
    TerminateMessageMain = mTerminateMessageMain
End Property

Public Property Get TerminateMessageAddition() As String
    TerminateMessageAddition = mTerminateMessageAddition
End Property
Public Function Main(isMatchSection As Boolean, isLimitKinkAngle As Boolean) As Integer
    
    'Form Str Obj
'    If Not isBySec Then ReadJoints
'    ReadFrameConn
    Dim ret As Integer
    
    If Not CheckDataBeforeStart(isMatchSection, isLimitKinkAngle) = 0 Then
        mTerminateMessageMain = "No Sufficient Data in the Workbook! The macro will be terminated."
        ret = -1
        GoTo ExitFunction
    End If
    
    If Not mLog Is Nothing Then mLog.WriteLog "Frame Connectivity Analysis Started."
    FormStrObj
    
    Dim coll_Frms As New Collection 'frms to be matched. Clone from the model.frames.
    Dim coll_matchedFrms As New Collection 'matched frames in the cMember object.
    Dim cFrm As StrFrame
    
    Dim cMember As StrMember, coll_members As New Collection
    
    Dim i As Long
    Dim numFrms As Long
    'Dim count As Long ', str_member As String
    
    Set coll_Frms = CloneCollection(mModel.frames)
    
    numFrms = coll_Frms.count
    
    Do Until coll_Frms.count = 0
        Application.StatusBar = "Progress for Connectivity Analysis = " & coll_Frms.count & "/" & numFrms
        If Not mLog Is Nothing Then mLog.WriteLog "Progress for Connectivity Analysis = " & coll_Frms.count & "/" & numFrms
        Set cMember = New StrMember
        Set cFrm = New StrFrame
        Set cFrm = coll_Frms(1)
        TraceFrame cFrm, True, cMember, coll_matchedFrms, isMatchSection, isLimitKinkAngle, mDefaultAngle 'find the previous connected frames and save to 'cMember'
        TraceFrame cFrm, False, cMember, coll_matchedFrms, isMatchSection, isLimitKinkAngle, mDefaultAngle 'find the end connected frames and save to 'cMember'
        
        'Create member if needed.
'        If IsKeyExist(coll_members, cMember.name) Then
'            count = 1
'            str_member = cMember.name & "(" & count & ")"
'            Do Until Not IsKeyExist(coll_members, str_member)
'                coll_members.Add cMember, str_member
'            Loop
'        Else
        coll_members.Add cMember, cMember.name
        If Not mLog Is Nothing Then mLog.WriteLog "Member '" & cMember.name & "' created."
'        End If

        'save the connectivity result back to the frames obj & remove the already matched frames
        'Set coll_matchedFrms = cMember.frames
        With cMember
            For i = 1 To .frames.count
                Set mModel.GetStrObject(.frames(i).name, obj_frm).member = cMember
                If isKeyExist(coll_Frms, .frames(i).name) Then coll_Frms.Remove .frames(i).name
            Next i
        End With
    Loop
    Application.StatusBar = False
    Set mModel.members = coll_members
    
    WriteResult
    g_log.WriteLog "Member Creation Completed."
    g_log.WriteLog ""

    mDsSys.prop("isCreated", "member") = True
ExitFunction:
    Main = ret
End Function

Public Sub Initialize(logFile As clsLog)
    Set mLog = logFile
    mDefaultAngle = CDbl(mDsSys.prop("Setting_Conn_ByDir", "kinkAngle"))
End Sub

Private Function CheckDataBeforeStart(isMatchSection As Boolean, isLimitKinkAngle As Boolean) As Integer
    Application.Calculate
    If isMatchSection And Not mDsSys.prop("WhatOperationCanProcess", "CreateMemberBySection") Then
        CheckDataBeforeStart = -1
    End If
    If isLimitKinkAngle And Not mDsSys.prop("WhatOperationCanProcess", "CreateMemberByKinkAngle") Then
        CheckDataBeforeStart = -1
    End If
    If CheckDataBeforeStart = -1 Then mCompleteMsg = "Not Sufficient Data to carry out the process!!"
End Function

Private Sub FormStrObj()
    mModel.Constructor.FormJointObj
    mModel.Constructor.FormFrmObj
End Sub

Private Sub TraceFrame(fFrm As StrFrame, traceDirI As Boolean, _
                        cMember As StrMember, matchedFrms As Object, isMatchSection As Boolean, isLimitKinkAngle As Boolean, limitKinkAngle As Double)
    'This Sub trace the frame by section. The sub will set the cMember
    'Dim matchedFrms As New Collection
    Dim cFrm As StrFrame, pFrm As StrFrame, cJt As StrJoint
    Dim connectedFrms As New Collection
    Dim isFound As Boolean 'to determine if the first/last frame is found
    Dim theta As Double
    Dim i As Long

    isFound = False
    
    Set cFrm = fFrm
    
    If Not isKeyExist(matchedFrms, cFrm.name) Then
        cMember.AddFrame cFrm, traceDirI
        matchedFrms.Add cFrm, cFrm.name
    End If
    
    Do Until isFound

        If traceDirI Then
            Set cJt = cFrm.jtI
            Set connectedFrms = cJt.connectedFrames_JtJ
        Else
            Set cJt = cFrm.jtJ
            Set connectedFrms = cJt.connectedFrames_JtI
        End If

        '1. find  the frames that connected to the joint. see if sucess. if failed to look up  -> it is the first/ last frame.
        If connectedFrms.count = 0 Then Exit Do
        
        For i = 1 To connectedFrms.count
            If IsSameMember(cFrm, connectedFrms(i), traceDirI, isMatchSection, isLimitKinkAngle, limitKinkAngle) Then
                If Not isKeyExist(matchedFrms, connectedFrms(i).name) Then
                    cMember.AddFrame connectedFrms(i), traceDirI
                    matchedFrms.Add connectedFrms(i), connectedFrms(i).name
                    Set cFrm = connectedFrms(i)
                    GoTo NextDo
                End If
            End If
        Next i
        isFound = True
NextDo:
    Loop

End Sub

Private Function IsSameMember(cFrm As StrFrame, connectedFrm As StrFrame, traceDirI As Boolean, _
                                isMatchSection As Boolean, isLimitKinkAngle As Boolean, limitKinkAngle As Double) As Boolean
    Dim ret As Boolean
    If isMatchSection Then
        ret = IsSameSection(cFrm, connectedFrm)
        If Not ret Then GoTo ExitFunction
    End If
    
    If isLimitKinkAngle Then ret = IsKinkAngleWithinLimit(cFrm, connectedFrm, traceDirI, limitKinkAngle)
ExitFunction:
    IsSameMember = ret
End Function

Private Function IsSameSection(cFrm As StrFrame, connectedFrm As StrFrame) As Boolean
    If connectedFrm.section = cFrm.section Then IsSameSection = True
End Function

Private Function IsKinkAngleWithinLimit(cFrm As StrFrame, connectedFrm As StrFrame, traceDirI As Boolean, limitKinkAngle As Double) As Boolean
    Dim theta As Double
    If traceDirI Then
        theta = AngleBetween3Points(cFrm.jtJ.x, cFrm.jtJ.y, cFrm.jtJ.z, _
                                cFrm.jtI.x, cFrm.jtI.y, cFrm.jtI.z, _
                                connectedFrm.jtI.x, connectedFrm.jtI.y, _
                                connectedFrm.jtI.z, False)
    Else
        theta = AngleBetween3Points(cFrm.jtI.x, cFrm.jtI.y, cFrm.jtI.z, _
                                cFrm.jtJ.x, cFrm.jtJ.y, cFrm.jtJ.z, _
                                connectedFrm.jtJ.x, connectedFrm.jtJ.y, _
                                connectedFrm.jtJ.z, False)
    End If
    If theta > (180 - limitKinkAngle) Then IsKinkAngleWithinLimit = True
End Function


Private Sub WriteResult()
    
    Dim df As clsDataFrame
    Set df = mModel.GetDataframe(obj_member, "name", "iFrameName", "jFrameName", "TotalLength", "eleList")
    
    Dim ds_member As oDataSheet
    Set ds_member = mDS.DSMemberData
    ds_member.ClearData
    
    g_log.WriteLog "Writing result to the worksheets...."
    With ds_member.tagSelector
        ds_member.WriteDataframe df, True, False, .memID, .fFrame, .lFrame, .totalLen, .eleList
    End With
    
    Dim ds_frm As oDataSheet
    Set ds_frm = mDS.DSFrameData
    Set df = mModel.GetDataframe(obj_frm, "memberName")
    With ds_frm.tagSelector
        ds_frm.WriteColumn df.iColumn(1), True, .memID
    End With
    
    Dim ds_frmForce As oDataSheet
    Set ds_frmForce = mDS.DSFrameForce
    With ds_frmForce.tagSelector
        ds_frmForce.FillColumns .memID
    End With
End Sub

Private Function CloneCollection(coll) As Object
    Dim newColl As New Collection
  
    Dim i As Long
    For i = 1 To coll.count
        newColl.Add coll(i), coll(i).name
    Next i
    
    Set CloneCollection = newColl
End Function

Private Function AngleBetween3Points(x1 As Double, y1 As Double, z1 As Double, _
                                    x2 As Double, y2 As Double, z2 As Double, _
                                    x3 As Double, y3 As Double, z3 As Double, _
                                    Optional isRad As Boolean = True) As Double
    Dim a As Double, b As Double, c As Double, theta As Double
    Dim sum As Double
    a = Dist_2Points(x1, y1, z1, x2, y2, z2)
    b = Dist_2Points(x2, y2, z2, x3, y3, z3)
    c = Dist_2Points(x1, y1, z1, x3, y3, z3)
    
    If a = 0 Or b = 0 Then Exit Function
    sum = (a ^ 2 + b ^ 2 - c ^ 2) / (2 * a * b)
    If sum < -1 And sum > -1.000000001 Then sum = -1
    If sum > 1 And sum < 1.000000001 Then sum = 1
    AngleBetween3Points = Application.WorksheetFunction.Acos(sum)
    
    If isRad Then
        Exit Function
    Else
        AngleBetween3Points = AngleBetween3Points * 180 / PI
    End If
End Function

Private Function Dist_2Points(x1 As Double, y1 As Double, z1 As Double, _
                              x2 As Double, y2 As Double, z2 As Double) As Double
    Dist_2Points = ((x2 - x1) ^ 2 + (y2 - y1) ^ 2 + (z2 - z1) ^ 2) ^ (1 / 2)
End Function

Private Function PI() As Double
    PI = 3.14159265358979
End Function

'***********************Functions for Collection**************************************
Private Function isKeyExist(coll As Collection, sKey As String) As Boolean
    Dim obj As Object
    On Error GoTo NotExist:
    Set obj = coll(sKey)
    isKeyExist = True
    Exit Function
NotExist:
    isKeyExist = False
    On Error GoTo -1
    'Debug.Print "Err Num: =" & Err.Number
End Function



