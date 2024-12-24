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
Private genFuncs As New clsGeneralFunctions
Private mUFControl As VMForceExtraction
Private lc() As String, MemberNames() As String, sections() As String 'for use in Userform
Private lcExcluded() As String, MemberNamesExcluded() As String, sectionsExcluded() As String 'for use in Userform
Private PosFilterType As PositionFilterType, posFilterVal As Double
Private IsBySec As Boolean 'true = by section, false = by member
Private isExtremeCase As Boolean, IsEachLC As Boolean, isEnvMaxMin As Boolean, isEachPos As Boolean
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
        lcExcluded = .lcExcl
        MemberNamesExcluded = .MemberNamesExcl
        PosFilterType = .PosFilterType
        sectionsExcluded = .SectionsExcl
        posFilterVal = .PosFilterValue
        IsBySec = .IsBySec
        isExtremeCase = True
        IsEachLC = .IsEachLC
        If isExtremeCase Then ExtremeCase = .ExtremeCase
        isEnvMaxMin = .isEnvMaxMin
        MaxMinEnv = .MaxMinEnv
        isEachPos = .isEachPos
    End With
End Function

Private Function ExtractCorrespondenceFrmForces() As Collection

    GetUserInput

    'Logic: 1. Loop Each Section/Member, filter frms. 2. (Loop each LC) 3.Filter FrameForces 4.Operation
    Dim filteredFrmForces As Collection, groupedFrmForces As Object
    Dim key As Variant, keys As Variant
    Dim cFrms As New Collection, cFrmForces As New Collection
    Dim reFrmForces As New Collection 'extracted frm force
    Dim frmForceAnalyser As StrFrmForceAnalyser
    Dim i As Long, n As Long
    'Looping for each section/member
    g_log.WriteLog "Filtering Frame Forces according to user input..."
    Set filteredFrmForces = FilterFrmForces
    
    g_log.WriteLog "Grouping frame forces according to user input"
    Set groupedFrmForces = GroupingFrmForces(filteredFrmForces)
    
    g_log.WriteLog "Grouping frame forces..."
    Set frmForceAnalyser = pModel.frmForceAnalyser(filteredFrmForces)
    
    keys = GetSortedKeys(groupedFrmForces)
    n = UBound(keys) - LBound(keys) + 1
    g_log.WriteLog "Total number of sections/members = " & UBound(keys) - LBound(keys) + 1
    For i = LBound(keys) To UBound(keys) 'looping for each section/ member
         g_log.WriteLog "Start extracting forces for no. " & i + 1 & "/" & n & " grouped result. key = " & keys(i)
         Application.StatusBar = "Start extracting forces for no. " & i + 1 & "/" & n & " grouped result."

        Set cFrmForces = groupedFrmForces(keys(i))
        If isEnvMaxMin Then
            Set cFrmForces = frmForceAnalyser.FindExtremeForceWithMaxMinEnv(cFrmForces, ExtremeCase, MaxMinEnv) 'Operation (Man Min Env)
        Else
            Set cFrmForces = frmForceAnalyser.FindExtremeForce(cFrmForces, ExtremeCase)
        End If
        AddColltoColl reFrmForces, cFrmForces 'Save Data
NextIteration:
    Next i
    
    Set ExtractCorrespondenceFrmForces = reFrmForces
End Function

Private Function FilterFrmForces() As Collection
'This function filter frame forces from ALL frame forces based on user input.
'The return frame forces will be used to
    Dim filteredFrmForces As New Collection
    Set filteredFrmForces = pModel.frmForces
    'ReDim filterBool(0)
    'Filter Section
    g_log.WriteLog "Number of Frame Force Objects = " & filteredFrmForces.count
    If genFuncs.isInitialised(sectionsExcluded) Then Set filteredFrmForces = pModel.FilterCollOfObj(filteredFrmForces, "sectionName", sections)
    g_log.WriteLog "Number of Frame Force Objects after applying section filter = " & filteredFrmForces.count
    'Filter Members
    If genFuncs.isInitialised(MemberNamesExcluded) Then Set filteredFrmForces = pModel.FilterCollOfObj(filteredFrmForces, "memberName", MemberNames)
    g_log.WriteLog "Number of Frame Force Objects after applying member filter = " & filteredFrmForces.count
    'Filter load combinations
    If genFuncs.isInitialised(lcExcluded) Then Set filteredFrmForces = pModel.FilterCollOfObj(filteredFrmForces, "loadComb", lc)
    g_log.WriteLog "Number of Frame Force Objects after applying load combination filter = " & filteredFrmForces.count
    'Filter Position
    'filterBool(0) = True
    If PosFilterType = BothEnd Then
        Set filteredFrmForces = pModel.FilterCollOfObj(filteredFrmForces, "isMemberEnd", Array(True))
    ElseIf PosFilterType = EndI Then
        Set filteredFrmForces = pModel.FilterCollOfObj(filteredFrmForces, "isMemberEndI", Array(True))
    ElseIf PosFilterType = EndJ Then
        Set filteredFrmForces = pModel.FilterCollOfObj(filteredFrmForces, "isMemberEndJ", Array(True))
    End If
    g_log.WriteLog "Number of Frame Force Objects after applying position filter = " & filteredFrmForces.count
    Set FilterFrmForces = filteredFrmForces
End Function

Private Function GroupingFrmForces(frmForces As Collection) As Object
'Save the filter frmForces into a dictionary according to user input for further processing
    Dim dict As Object
    Dim key As Variant, keyPrefix As Variant, frmForce As Variant
    Dim coll As Collection
    Set dict = CreateObject("Scripting.Dictionary")
    For Each frmForce In frmForces
    
        If IsBySec Then
            keyPrefix = frmForce.frame.section
        Else
            keyPrefix = frmForce.memberName
        End If
        
        If isEachPos And IsEachLC Then
            key = keyPrefix & frmForce.loadComb & frmForce.pos_fromMemJtI_percent
        ElseIf isEachPos And Not IsEachLC Then
            key = keyPrefix & frmForce.pos_fromMemJtI_percent
        ElseIf Not isEachPos And IsEachLC Then
            key = keyPrefix & frmForce.loadComb
        Else
            key = keyPrefix
        End If
        
        If dict.Exists(key) Then
            dict(key).Add frmForce
        Else
            Set coll = New Collection
            coll.Add frmForce
            dict.Add key, coll
            
        End If
    Next
    
    Set GroupingFrmForces = dict
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

Function GetSortedKeys(dict As Object) As Variant
    Dim keys As Object
    Dim key As Variant
    Dim sortedArray() As Variant
    
    ' Create a new ArrayList using late binding
    Set keys = CreateObject("System.Collections.ArrayList")
    
    ' Add each key from the dictionary to the ArrayList
    For Each key In dict.keys
        keys.Add key
    Next key
    
    ' Sort the keys
    keys.Sort
    
    ' Convert the sorted ArrayList to an array
    sortedArray = keys.ToArray
    
    ' Return the sorted array
    GetSortedKeys = sortedArray
End Function


