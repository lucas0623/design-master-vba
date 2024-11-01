VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtractFrmForceMethodCo"
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
        IsBySec = .IsBySec
        isExtremeCase = True
        IsEachLC = .IsEachLC
        If isExtremeCase Then ExtremeCase = .ExtremeCase
        MaxMinEnv = .MaxMinEnv
    End With
End Function

Private Function ExtractCorrespondenceFrmForces() As Collection

    GetUserInput

    'Logic: 1. Loop Each Section/Member, filter frms. 2. (Loop each LC) 3.Filter FrameForces 4.Operation
    Dim cFrms As New Collection, cFrmForces As New Collection, tempFrmForces As New Collection
    Dim reFrmForces As New Collection 'extracted frm force
    Dim i As Long, j As Long, k As Long, uLimit As Long
    'Looping for each section/member
    If IsBySec Then
        uLimit = UBound(sections)
    Else
        uLimit = UBound(MemberNames)
    End If
    
    g_log.WriteLog "Current Procedure: OperExtractFrmForceCo.ExtractCorrespondenceFrmForces"
    g_log.WriteLog "IsBySec = " & CStr(IsBySec)
    g_log.WriteLog "Total number of sections/members = " & uLimit + 1
    For i = 0 To uLimit 'looping for each section/ member
         g_log.WriteLog "Start extracting forces for no. " & i + 1 & "/" & uLimit + 1 & " sections/members."
         Application.StatusBar = "Start extracting forces for no. " & i + 1 & "/" & uLimit + 1 & " sections/members."
         g_log.WriteLogWithTime "   Start filtering objects for force extraction"
        'filter by section/ membernames
        If IsBySec Then
            g_log.WriteLog "   Section name = " & sections(i)
            Set cFrms = pModel.frmsBySection(sections(i))
            'pModel.FilterCollOfObjWithKey cFrms, sections(i)
            Set cFrms = pModel.FilterCollOfObj(cFrms, "memberName", MemberNames)
        Else
            g_log.WriteLog "   Member name = " & MemberNames(i)
            Set cFrms = pModel.frmsByMember(MemberNames(i))
            Set cFrms = pModel.FilterCollOfObj(cFrms, "section", sections)
        End If
        
        If cFrms.count = 0 Then
            g_log.RaiseWarning "Cannot find the above section/member. Record Skip.", frmForceExtract_objNotFound
            GoTo NextIteration
        End If
        
        If IsEachLC Then
            For j = 0 To UBound(lc)
                g_log.WriteLogWithTime "   Current load combination/case = " & lc(j)
                g_log.WriteLogWithTime "   Filtering Load combination/case."
                Set cFrmForces = GetFrmForcesColl(cFrms, loadCombFilter:=lc(j), posType:=PosFilterType, posVal:=posFilterVal) 'Filter Frame Forces
                g_log.WriteLogWithTime "   Filter Completed. Remaining Number of Frame Force Object = " & cFrmForces.count
                Set cFrmForces = pModel.frmForceAnalyser.FindExtremeForceWithMaxMinEnv(cFrmForces, ExtremeCase, MaxMinEnv) 'Operation (Man Min Env)
                AddColltoColl reFrmForces, cFrmForces 'Save Data
                g_log.WriteLogWithTime "   " & cFrmForces.count & " nos. of frame forces results are saved. (Cumulative Total: = " & reFrmForces.count & ")"
            Next j
        Else
            g_log.WriteLogWithTime "   Filtering Load combination/case = " & lc(j)
            Set cFrmForces = GetFrmForcesColl(cFrms, loadCombFilter:=lc, posType:=PosFilterType, posVal:=posFilterVal) 'Filter Frame Forces
            g_log.WriteLogWithTime "   Filter Completed. Remaining Number of Frame Force Object = " & cFrmForces.count
            Set cFrmForces = pModel.frmForceAnalyser.FindExtremeForceWithMaxMinEnv(cFrmForces, ExtremeCase, MaxMinEnv) 'Operation (Man Min Env)
            AddColltoColl reFrmForces, cFrmForces 'Save Data
            g_log.WriteLogWithTime "   " & cFrmForces.count & " nos. of frame forces results are saved. (Cumulative Total: = " & reFrmForces.count & ")"
        End If
NextIteration:
    Next i
    
    Set ExtractCorrespondenceFrmForces = reFrmForces
End Function


Private Function GetFrmForcesColl(frames As Collection, Optional loadCombFilter As Variant, Optional stepTypeFilter As Variant, _
                                Optional posType As PositionFilterType, Optional posVal As Double = -1) As Collection
    'This Function filter the frm force objects. Input = a collection of Frame Objects. Output = a Collection of Frame Force Objects
    Dim coll_frmForce As New Collection
    Dim tempColl As New Collection
    Dim i As Long, frame As StrFrame
    
    For Each frame In frames
        If posType = NoFilter Then
            Set tempColl = frame.GetFrameForces(loadCombFilter:=loadCombFilter, stepTypeFilter:=stepTypeFilter)
        ElseIf posType = BothEnd Then
            Set tempColl = frame.GetFrameForces_BothEnd(loadCombFilter:=loadCombFilter, stepTypeFilter:=stepTypeFilter)
        ElseIf posType = EndI Then
            Set tempColl = frame.GetFrameForces_EndI(loadCombFilter:=loadCombFilter, stepTypeFilter:=stepTypeFilter)
        ElseIf posType = EndJ Then
            Set tempColl = frame.GetFrameForces_EndJ(loadCombFilter:=loadCombFilter, stepTypeFilter:=stepTypeFilter)
        ElseIf posType = SpecifiedPosAbs Then
            Set tempColl = frame.GetFrameForces(loadCombFilter:=loadCombFilter, stepTypeFilter:=stepTypeFilter, SpecifiedStation:=posFilterVal)
        End If
        AddColltoColl coll_frmForce, tempColl
    Next
    
    Set GetFrmForcesColl = coll_frmForce
End Function

Private Sub AddColltoColl(coll As Collection, coll2 As Collection)
    Dim i As Long
    For i = 1 To coll2.count
        coll.Add coll2(i)
    Next i
End Sub

Private Function IExtractFrmForceMethod_ExtractForce() As Collection
    Set IExtractFrmForceMethod_ExtractForce = ExtractCorrespondenceFrmForces
End Function


