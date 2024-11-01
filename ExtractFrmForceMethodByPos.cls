VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtractFrmForceMethodByPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Operation.ExtractFrameForce"

Implements IExtractFrmForceMethod
Private mUFControl As VMForceExtraction
Private lc() As String, MemberNames() As String, sections() As String 'for use in Userform
Private PosFilterType As PositionFilterType, posFilterVal As Double
Private IsBySec As Boolean 'true = by section, false = by member
Private isExtremeCase As Boolean, IsEachLC As Boolean
Private ExtremeCase() As Boolean
Private MaxMinEnv() As Boolean
Private pModel As StrModel

Private Sub Class_Initialize()
    IsBySec = False
    isExtremeCase = False
    IsEachLC = True
End Sub
Private Sub IExtractFrmForceMethod_Initialize(model As StrModel, UFControl As Object)
    Set mUFControl = UFControl
    Set pModel = model
End Sub

Private Function GetUserInput() As Integer
    Dim ret As Integer
    With mUFControl
        'Set ws_sum = .ws_sum
        lc = .lc
        MemberNames = .MemberNames
        sections = .sections
        PosFilterType = .PosFilterType
        posFilterVal = .PosFilterValue
        IsEachLC = .IsEachLC
        'If isExtremeCase Then ExtremeCase = .ExtremeCase
        MaxMinEnv = .MaxMinEnv
    End With
End Function

Private Function ExtractFrmForcesByPos() As Collection

    GetUserInput

    'Logic: 1. Loop Each Section/Member, filter frms. 2. (Loop each LC) 3.Filter FrameForces 4.Operation
    Dim frmsByMem As New Collection, frm As Object
    Dim frmForcesByPos As New Collection, frmForcesAtPos As New Collection
    Dim frmForces As New Collection
    Dim reFrmForces As New Collection 'extracted frm force
    Dim i As Long, j As Long, k As Long, uLimit As Long
    'Looping for each section/member
'    If IsBySec Then
'        uLimit = UBound(Sections)
'    Else
        uLimit = UBound(MemberNames)
'    End If
    
    g_log.WriteLog "Current Procedure: OperExtractFrmForceCo.ExtractFrmForcesByPos"
    g_log.WriteLog "IsBySec = " & CStr(IsBySec)
    g_log.WriteLog "Total number of sections/members = " & uLimit + 1
    For i = 0 To uLimit 'looping for each section/ member
        g_log.WriteLog "Start extracting forces for no. " & i + 1 & "/" & uLimit + 1 & " sections/members."
        Application.StatusBar = "Start extracting forces for no. " & i + 1 & "/" & uLimit + 1 & " sections/members."

        'filter by section/ membernames
'        If IsBySec Then
'            g_log.WriteLog "   Section name = " & Sections(i)
'            Set cFrms = pModel.frames(sectionFilter:=Sections(i))
'        Else
        g_log.WriteLog "   Member name = " & MemberNames(i)
        Set frmsByMem = pModel.member(MemberNames(i)).frames
'        End If
        
        If frmsByMem.count = 0 Then
            g_log.RaiseWarning "Cannot find the above section/member. Record Skip.", frmForceExtract_objNotFound
            GoTo NextIteration
        End If
        For k = 0 To UBound(lc)
        For Each frm In frmsByMem 'for each element in the member
            
                Set frmForcesByPos = GetFrmForcesCollByPos(frm, loadCombFilter:=lc(k), posType:=PosFilterType, posVal:=posFilterVal)
                For Each frmForcesAtPos In frmForcesByPos
                    Set frmForces = pModel.frmForceAnalyser.EnvMaxMin(frmForcesAtPos, MaxMinEnv)  'Operation (Man Min Env)
                    AddColltoColl reFrmForces, frmForces 'Save Data
                    g_log.WriteLog "   " & frmForces.count & " nos. of frame forces results are saved. (Cumulative Total: = " & reFrmForces.count & ")"
                Next
            
        Next
        Next k
NextIteration:
    Next i
    
    Set ExtractFrmForcesByPos = reFrmForces
End Function


Private Function GetFrmForcesCollByPos(frm As StrFrame, Optional loadCombFilter As Variant, Optional stepTypeFilter As Variant, _
                                Optional posType As PositionFilterType, Optional posVal As Double = -1, Optional groupByProp As String = "station") As Collection
    'This Function filter the frm force objects. Input = a collection of Frame Objects. Output = a Collection of Frame Force Objects
    Dim coll_frmForce As New Collection
    Dim tempColl As New Collection
    Dim i As Long
    
    If posType = NoFilter Then
        Set tempColl = frm.GetFrameForces(loadCombFilter:=loadCombFilter, stepTypeFilter:=stepTypeFilter)
    ElseIf posType = BothEnd Then
        Set tempColl = frm.GetFrameForces_BothEnd(loadCombFilter:=loadCombFilter, stepTypeFilter:=stepTypeFilter)
    ElseIf posType = EndI Then
        Set tempColl = frm.GetFrameForces_EndI(loadCombFilter:=loadCombFilter, stepTypeFilter:=stepTypeFilter)
    ElseIf posType = EndJ Then
        Set tempColl = frm.GetFrameForces_EndJ(loadCombFilter:=loadCombFilter, stepTypeFilter:=stepTypeFilter)
    ElseIf posType = SpecifiedPosAbs Then
        Set tempColl = frm.GetFrameForces(loadCombFilter:=loadCombFilter, stepTypeFilter:=stepTypeFilter, SpecifiedStation:=posFilterVal)
    End If
    AddColltoColl coll_frmForce, tempColl
    
    Set GetFrmForcesCollByPos = GroupCollByProp(coll_frmForce, groupByProp)
End Function

Private Function GroupCollByProp(coll As Collection, propName As String) As Collection
    'Return as a collection of collection
    Dim i As Long, key As String
    Dim coll_grouped As New Collection
    Dim keys As New Collection
    For i = 1 To coll.count
        key = CallByName(coll(i), propName, VbGet)
        If IsKeyExistVar(keys, key) Then
            coll_grouped(key).Add coll(i)
        Else
            'Dim tempColl As New Collection
            keys.Add key, key
            'tempColl.Add coll(i)
            coll_grouped.Add New Collection, key
            coll_grouped(key).Add coll(i)
        End If
    Next i
    Set GroupCollByProp = coll_grouped
End Function

Private Function IsKeyExistVar(coll As Collection, sKey As String) As Boolean
    'Input shall be colleciton of OBJECT
    Dim var As Variant
    On Error GoTo NotExist:
    var = coll(sKey)
    IsKeyExistVar = True
    Exit Function
NotExist:
    IsKeyExistVar = False
    On Error GoTo -1
    'Debug.Print "Err Num: =" & Err.Number
End Function

Private Sub AddColltoColl(coll As Collection, coll2 As Collection)
    Dim i As Long
    For i = 1 To coll2.count
        coll.Add coll2(i)
    Next i
End Sub

Private Function IExtractFrmForceMethod_ExtractForce() As Collection
    Set IExtractFrmForceMethod_ExtractForce = ExtractFrmForcesByPos
End Function


