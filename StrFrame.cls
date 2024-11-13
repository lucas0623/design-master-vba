VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StrFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Structural Model"
'@Folder("Structural Model Object")

'********************************************************
'This module represent the Structural Joint/Node Object
'Arthor: Lucas LEUNG
'Update Log

'04 Apr 2023 - Initial
'*******************************************************

Option Explicit
'this is the variable for the sub frame
Private mName As String ', mSubFrameName As String
Private pLength As Double
Private pJtI As New StrJoint, pJtJ As New StrJoint
Private mOrientationType As EleOrientationType
Private pSection As String
Private pLocalAxis As Double
Private pEndOffsetI As Double, pEndOffsetJ As Double
Private pRigidFac As Double
Private pReleaseI(1 To 6) As Boolean, pReleaseJ(1 To 6) As Boolean
Private pMaterial As String
Private pMember As StrMember
Private mFrmID As Long
Private mEleType As EleType


Private EnumEleOrientationType As New EnumEleOrientationType
Private EnumEleType As New EnumEleType

Private Enum DefReleaseType
    Moment = 1
    MomentAndTorsion = 2
    Axial = 3
End Enum


'Private pFrmLoads As New Collection
'Private pOrientation As Double

'For FrameForce
Private mFrameForces As New Collection
Private mFrameForcesByLC As New Collection ', mFrameForcesByStep As New Collection, mFrameForcesByPos As New Collection 'Collection of Collection, key = Pos.
'Private mFrameForcesByLCnStep As New Collection, mFrameForcesByLCnPos As New Collection, mFrameForcesByStepnPos As New Collection
'Private mFrameForcesByLCnStepnPos As New Collection

Public Sub Init(Name As String, Optional section As String, Optional jtI As StrJoint, _
                Optional jtJ As StrJoint, Optional ID As Long = 0, Optional Length As Double = 0)
    mName = Name
    pSection = section
    If Not jtI Is Nothing Then
        Set pJtI = jtI
        pJtI.AddConnectedFrm Me
    End If
    
    If Not jtJ Is Nothing Then
        Set pJtJ = jtJ
        pJtJ.AddConnectedFrm Me
    End If
    
    If Length > 0 Then
        pLength = Length
    Else
        pLength = ((jtI.x - jtJ.x) ^ 2 + (jtI.y - jtJ.y) ^ 2 + (jtI.z - jtJ.z) ^ 2) ^ (1 / 2)
    End If
    
End Sub

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(str As String)
    mName = str
End Property

'Public Property Get subFrameName() As String
'    subFrameName = mSubFrameName
'End Property
'
'Public Property Let subFrameName(str As String)
'    mSubFrameName = str
'End Property

Public Property Get jtI() As StrJoint
    Set jtI = pJtI
End Property
Public Property Get jtJ() As StrJoint
    Set jtJ = pJtJ
End Property
Public Property Get jtIName() As String
    jtIName = pJtI.Name
End Property

Public Property Get jtJName() As String
    jtJName = pJtJ.Name
End Property

Public Property Get section() As String
    section = pSection
End Property

Public Property Let section(str As String)
    pSection = str
End Property
Public Property Get EleType() As EleType
    EleType = mEleType
End Property

Public Property Let EleType(rhs As EleType)
    mEleType = rhs
End Property

Public Property Get EleTypeStr() As String
    EleTypeStr = EnumEleType.EnumToStr(mEleType)
End Property

Public Property Let EleTypeStr(str As String)
    mOrientationType = EnumEleType.StrToEnum(str)
End Property

Public Property Get orientationType() As EleOrientationType
    orientationType = mOrientationType
End Property

Public Property Let orientationType(rhs As EleOrientationType)
    mOrientationType = rhs
End Property

Public Property Get OrientationTypeStr() As String
    OrientationTypeStr = EnumEleOrientationType.EnumToStr(mOrientationType)
End Property

Public Property Let OrientationTypeStr(str As String)
    mOrientationType = EnumEleOrientationType.StrToEnum(str)
End Property

Public Property Get localAxis() As Double
    localAxis = pLocalAxis
End Property

Public Property Let localAxis(angle As Double)
    pLocalAxis = angle
End Property

Public Property Get Length() As Double
    Length = pLength
End Property

Public Property Get endOffLenI() As Double
    endOffLenI = pEndOffsetI
End Property

Public Property Get rigidFac() As Double
    rigidFac = pRigidFac
End Property

'Public Property Get fRelease_str(i As Long) As String
'    If pfRelease(i) Then
'        fRelease_str = "Yes"
'    Else
'        fRelease_str = "No"
'    End If
'End Property

'Public Property Get isRelease() As Boolean
'    Dim rele As Boolean
'    For Each rele In pRelease
'        If rele Then
'            isRelease = True
'            Exit For
'        Else
'            isRelease = False
'        End If
'    Next
'End Property
Public Property Get isAnyRelease_str(Optional hvReleaseText As String = "RLS", Optional noReleaseText As String = "NO_RLS") As String
    Dim i As Integer
    Dim isAnyRelease As Boolean
    For i = 1 To 6
        If pReleaseI(i) Or pReleaseJ(i) Then
            isAnyRelease = True
            Exit For
        End If
    Next i
    
    If isAnyRelease Then
        isAnyRelease_str = hvReleaseText
    Else
        isAnyRelease_str = noReleaseText
    End If
    
End Property
Public Property Get release_str(isEndI As Boolean, Optional sRelease As String = "R", Optional sFix As String = "F") As String
    Dim i As Integer, cRelease() As Boolean
    If isEndI Then
        cRelease = pReleaseI
    Else
        cRelease = pReleaseJ
    End If
        
    Dim str As String
    For i = 1 To 6
        If cRelease(i) Then
            str = str & sRelease
        Else
            str = str & sFix
        End If
    Next i
    release_str = str
End Property
Public Property Let releaseI(str As String)
    Dim rele() As Boolean
    rele = SetRelease(str)
    If UBound(rele) = 0 Then
        Err.Raise 5011, "clsStrFrame.ReleaseI", "Wrong Format of Frame Restraint Assignment Input"
        'g_Log.RaiseWarning "Joint '" & name & "' Restraint Assignment Data in Wrong Format. Plesae Review.", jtRestraintInputWrong
        Exit Property
    End If
    
    Dim i As Integer
    For i = 1 To 6
        pReleaseI(i) = rele(i)
    Next i
End Property
Public Property Let releaseJ(str As String)
    Dim rele() As Boolean
    rele = SetRelease(str)
    If UBound(rele) = 0 Then
        Err.Raise 5011, "clsStrFrame.ReleaseJ", "Wrong Format of Frame Restraint Assignment Input"
        'g_Log.RaiseWarning "Joint '" & name & "' Restraint Assignment Data in Wrong Format. Plesae Review.", jtRestraintInputWrong
        Exit Property
    End If
    
    Dim i As Integer
    For i = 1 To 6
        pReleaseJ(i) = rele(i)
    Next i
End Property
Public Function SetRelease(str As String) As Boolean()
    Dim i As Integer, release() As Boolean
    
    ReDim release(1 To 6)
    Select Case GetDefReleaseTypeVal(str)
    Case Moment
        str = "000011"
    Case MomentAndTorsion
        str = "000111"
    Case Axial
        str = "100000"
    End Select
    

    If str = vbNullString Then
        For i = 1 To 6
            release(i) = False
        Next i
    Else
        For i = 1 To 6
            If Mid(str, i, 1) = "1" Then
                release(i) = True
            ElseIf Mid(str, i, 1) = "0" Then
                release(i) = False
            Else
                'g_Log.RaiseWarning "Joint '" & name & "' Restraint Assignment Data in Wrong Format. Plesae Review.", jtRestraintInputWrong
                GoTo Err
            End If
        Next i
    End If

    SetRelease = release
    Exit Function
Err:
    ReDim release(0)
    SetRelease = release
End Function


Public Property Get ID() As Long
    ID = mFrmID
End Property
Public Property Let ID(index As Long)
    mFrmID = index
End Property
Public Property Get material() As String
    material = pMaterial
End Property

Property Get member() As StrMember
    Set member = pMember
End Property

Property Set member(mem As StrMember)
    Set pMember = mem
End Property

Property Get memberName() As String
    memberName = pMember.Name
End Property

Property Get frmForces() As Collection
    Set frmForces = mFrameForces
End Property

'Public Property Get frameForcesByPos() As Object
'    If mFrameForcesByPos.count = 0 Then Set mFrameForcesByPos = GroupCollByProp(mFrameForces, "station")
'    Set frameForcesByPos = mFrameForcesByPos
'End Property

Public Property Get frameForcesByLC() As Object
    If mFrameForcesByLC.count = 0 Then Set mFrameForcesByLC = GroupCollByProp(mFrameForces, "loadcomb")
    Set frameForcesByLC = mFrameForcesByLC
End Property

Private Function GetDefReleaseTypeVal(str) As DefReleaseType
    Select Case str
        Case "Moment": GetDefReleaseTypeVal = Moment
        Case "MomentAndTorsion": GetDefReleaseTypeVal = MomentAndTorsion
        Case "Moment And Torsion": GetDefReleaseTypeVal = MomentAndTorsion
        Case "Axial": GetDefReleaseTypeVal = Axial
    End Select
End Function

'Public Sub AssignLoad(loadPat As String, dir As String, val As Double, Optional loadGrp As String = vbNullString)
'    Dim frmLoad As clsStrFrameLoad
'    Set frmLoad = New clsStrFrameLoad
'
'    frmLoad.Init loadPat, Me, dir, val, loadGrp
'    pFrmLoads.Add frmLoad
'End Sub

'Property Get RecentAssignedLoad() As clsStrFrameLoad
'    Set RecentAssignedLoad = pFrmLoads.Item(pFrmLoads.count)
'End Property


'*****************************************************************************************************
'***************************************for Frame Force Control***************************************
'*****************************************************************************************************

Public Function AddFrameForceToColl(frmForce As StrFrameForce)
    mFrameForces.Add frmForce
End Function
Public Function GetFrameForces_EndI(Optional loadCombFilter As Variant, Optional stepTypeFilter As Variant) As Collection
    Set GetFrameForces_EndI = GetFrameForces(loadCombFilter, stepTypeFilter, 0)
End Function

Public Function GetFrameForces_EndJ(Optional loadCombFilter As Variant, Optional stepTypeFilter As Variant) As Collection
    Set GetFrameForces_EndJ = GetFrameForces(loadCombFilter, stepTypeFilter, Me.Length)
End Function

Public Function GetFrameForces_BothEnd(Optional loadCombFilter As Variant, Optional stepTypeFilter As Variant) As Collection
    Dim coll As New Collection
    coll.Add GetFrameForces_EndI(loadCombFilter, stepTypeFilter)
    coll.Add GetFrameForces_EndJ(loadCombFilter, stepTypeFilter)
    Set GetFrameForces_BothEnd = coll
End Function

'Public Function GetFrameForces(Optional loadCombFilter As Variant, Optional stepTypeFilter As Variant, _
'                                Optional SpecifiedStation As Double = -1) As Collection
'    Dim i As Long, j As Long
'    Dim coll As New Collection 'for saving the result frame forces
'    Set coll = mFrameForces
'
'
'    If isInitialised(loadCombFilter) And Not isInitialised(stepTypeFilter) And SpecifiedStation = -1 Then
'        If mFrameForcesByLC.count = 0 Then Set mFrameForcesByLC = GroupCollByProp(mFrameForces, "loadcomb")
'        Set coll = FilterCollOfObjWithKey(mFrameForcesByLC, CStr_arr(loadCombFilter))
'    ElseIf isInitialised(loadCombFilter) And isInitialised(stepTypeFilter) And SpecifiedStation = -1 Then
'        If mFrameForcesByStep.count = 0 Then Set mFrameForcesByStep = GroupCollByProp(mFrameForces, "stepType")
'        Set coll = FilterCollOfObjWithKey(mFrameForcesByStep, CStr_arr(stepTypeFilter))
'    ElseIf isInitialised(loadCombFilter) And isInitialised(stepTypeFilter) Then
'        If mFrameForcesByLCnStep.count = 0 Then Set mFrameForcesByLCnStep = GroupCollByProp(mFrameForces, "stepType")
'        Set coll = FilterCollOfObjWithKey(mFrameForcesByLCnStep, CStr_arr(loadCombFilter))
'    End If
'
'    If Not SpecifiedStation = -1 Then
'        Dim coll2 As New Collection, coll3 As New Collection
'        Dim stepTypes() As String, loadCombs() As String, count As Long
'        Dim re_coll As New Collection
'        If isInitialised(loadCombFilter) Then
'            loadCombs = CStr_arr(loadCombFilter)
'        Else
'            loadCombs = GetLoadCombs(coll)
'        End If
'
'        If isInitialised(stepTypeFilter) Then
'            stepTypes = CStr_arr(stepTypeFilter)
'        Else
'            stepTypes = GetStepTypes(coll)
'        End If
'
'        For i = 0 To UBound(loadCombs)
'            Set coll2 = FilterCollOfObj(coll, "loadComb", CStr_arr(loadCombs(i)))
'            For j = 0 To UBound(stepTypes)
'                Set coll3 = FilterCollOfObj(coll2, "stepType", CStr_arr(stepTypes(j)))
'                If Not coll3.count = 0 Then re_coll.Add GetFrameForceAtStation(coll3, SpecifiedStation)
'            Next j
'        Next i
'
'        Set coll = re_coll
'    End If
'
'    Set GetFrameForces = coll
'End Function

Public Function GetFrameForces(Optional loadCombFilter As Variant, Optional stepTypeFilter As Variant, _
                                Optional SpecifiedStation As Double = -1, Optional subFrameNameFilter As Variant) As Collection
    Dim i As Long, j As Long
    Dim coll As New Collection 'for saving the result frame forces
    Set coll = mFrameForces

'    If isInitialised(loadCombFilter) Then
'        If mFrameForcesByLC.count = 0 Then Set mFrameForcesByLC = GroupCollByProp(mFrameForces, "loadcomb")
'        Set coll = FilterCollOfObjWithKey(mFrameForcesByLC, CStr_arr(loadCombFilter))
'    End If
    If isInitialised(subFrameNameFilter) Then
        Set coll = FilterCollOfObj(coll, "subFrameName", CStr_arr(subFrameNameFilter))
    End If
    
    If isInitialised(loadCombFilter) Then
        If IsArray(loadCombFilter) Then
            Set coll = FilterCollOfObj(coll, "loadcomb", CStr_arr(loadCombFilter))
        Else
            Set coll = Me.frameForcesByLC(loadCombFilter)
        End If
    End If

    If isInitialised(stepTypeFilter) Then
        Set coll = FilterCollOfObj(coll, "stepType", CStr_arr(stepTypeFilter))
    End If

    If Not SpecifiedStation = -1 Then
        Dim coll2 As New Collection, coll3 As New Collection, coll4 As New Collection
        Dim stepTypes() As String, loadCombs() As String, count As Long
        Dim re_coll As New Collection
        If isInitialised(loadCombFilter) Then
            loadCombs = CStr_arr(loadCombFilter)
        Else
            loadCombs = GetLoadCombs(coll)
        End If

        If isInitialised(stepTypeFilter) Then
            stepTypes = CStr_arr(stepTypeFilter)
        Else
            stepTypes = GetStepTypes(coll)
        End If
        
        'Filter method 1, cannot interpolate station, return several results at same station
'        For i = 0 To UBound(loadCombs)
'            Set coll2 = FilterCollOfObj(coll, "loadComb", CStr_arr(loadCombs(i)))
'            For j = 0 To UBound(stepTypes)
'                Set coll3 = FilterCollOfObj(coll2, "stepType", CStr_arr(stepTypes(j)))
'                Set coll4 = FilterCollOfObj(coll3, "station", CStr_arr(SpecifiedStation)) 'this does not interpolate frame force
'                If Not coll4.count = 0 Then AddColltoColl re_coll, coll4
'                'If Not coll3.count = 0 Then re_coll.Add GetFrameForceAtStation(coll3, SpecifiedStation)
'            Next j
'        Next i
        
        'Filter Method 2, can interpolate station, but cannot deal with multiple results at the same station
        For i = 0 To UBound(loadCombs)
            Set coll2 = FilterCollOfObj(coll, "loadComb", CStr_arr(loadCombs(i)))
            For j = 0 To UBound(stepTypes)
                Set coll3 = FilterCollOfObj(coll2, "stepType", CStr_arr(stepTypes(j)))
                If Not coll3.count = 0 Then re_coll.Add GetFrameForceAtStation(coll3, SpecifiedStation)
            Next j
        Next i

        Set coll = re_coll
    End If

    Set GetFrameForces = coll
End Function

Public Function GetFrameForceAtStation(cFrmForce As Collection, station As Double) As StrFrameForce
    'Input =: the collection of the frame force of certain load combination AND certain stepType
    Dim i As Integer
    Dim Dist As Double 'distance between specified station and the existing station points
    Dim frmForce_smaller As StrFrameForce, dist_smaller As Double
    Dim frmForce_larger As StrFrameForce, dist_larger As Double
    Dim frmForce As StrFrameForce
    Dim isInterpolate As Boolean
    
    If station < 0 Or station > cFrmForce(1).frame.Length Then Exit Function
    
    dist_smaller = -99999999
    dist_larger = 99999999
    isInterpolate = True
    
    For i = 1 To cFrmForce.count
        'Check if there is existing station
        If cFrmForce(i).station = station Then
            Set frmForce = cFrmForce(i)
            isInterpolate = False
            Exit For
        End If
        
        'find the previous station
        Dist = cFrmForce(i).station - station

        If Dist < 0 And Dist > dist_smaller Then
            dist_smaller = Dist
            Set frmForce_smaller = cFrmForce(i)
        End If
        
        'find the next station
        If Dist > 0 And Dist < dist_larger Then
            dist_larger = Dist
            Set frmForce_larger = cFrmForce(i)
        End If
    Next i
    
    'interpolate for the result if needed
    If isInterpolate Then
        Set frmForce = InterpolateForces(frmForce_smaller, frmForce_larger, station)
    End If
    
    'Return Result
    Set GetFrameForceAtStation = frmForce
End Function

'Public Function GetFrameForcePair() As Collection
'    'Get the frame force with the same load combination and same station
'
'End Function

'Public Function EnvForcesAtPos(loadcomb As String, stepType As String, position As Double, isEnvForce() As Boolean) As clsStrFrameForce
'    'isEnvForce(0 to 5) represent the option to env max/min force pair for the P/V2/V3/T/M2/M3
'    'True = Env (Get the ABS Max Result) --> suitable when the direction of the force is not important
'    'Combinations will be created for all the 'false' options.
'
'End Function


Private Function InterpolateForces(F1 As StrFrameForce, F2 As StrFrameForce, station As Double) As StrFrameForce
    Dim frmForce As StrFrameForce
    Dim forces(5) As Double, rLen As Double
    Dim i As Integer
    
    Set frmForce = New StrFrameForce
    rLen = station - F1.station
    For i = 0 To 5
        forces(i) = Interpolate(F1.station, F1.force(i), F2.station, F2.force(i), rLen)
    Next i
    
    frmForce.Init F1.frame, station, F1.loadComb, F1.stepType, forces
    Set InterpolateForces = frmForce
End Function

Private Function GetStepTypes(coll As Collection) As String()
    Dim stepTypes() As String, i As Long
    ReDim stepTypes(coll.count - 1)
    For i = 1 To coll.count
        stepTypes(i - 1) = coll(i).stepType
    Next i
    stepTypes = CStr_arr(DeDupeOneDimArray(stepTypes, False))
    GetStepTypes = stepTypes
End Function
Private Function GetLoadCombs(coll As Collection) As String()
    Dim lc() As String, i As Long
    ReDim lc(coll.count - 1)
    For i = 1 To coll.count
        lc(i - 1) = coll(i).loadComb
    Next i
    lc = CStr_arr(DeDupeOneDimArray(lc, False))
    GetLoadCombs = lc
End Function
Private Function Interpolate(x1 As Double, y1 As Double, _
                        x2 As Double, y2 As Double, _
                        rLen As Double) As Double
                        
    Dim s As Double, ratio As Double
    Dim coor(1) As Double
    s = ((x2 - x1) ^ 2 + (y2 - y1) ^ 2) ^ (1 / 2)
    ratio = rLen / s
    Interpolate = y1 + (y2 - y1) * ratio
End Function

Private Function isInitialised(ByRef a As Variant) As Boolean
'This sub check if an ARRAY is initialized.
    isInitialised = False
    On Error GoTo ErrHandler
    If IsArray(a) Then
        If Not UBound(a) = -1 Then
            isInitialised = True
        End If
    ElseIf Not a = vbNullString Then
        isInitialised = True
    End If

    Exit Function
ErrHandler:
    isInitialised = False
End Function

Public Function FilterCollOfObj(coll As Collection, propName As String, criteria() As String) As Collection
    Dim i As Long
    Dim coll_filtered As New Collection
    Dim critColl As Object
    Dim key As Variant
    Set critColl = CreateObject("Scripting.Dictionary")

    For i = LBound(criteria) To UBound(criteria)
        critColl.Add criteria(i), criteria(i)
    Next i
    For i = 1 To coll.count

        key = CStr(CallByName(coll(i), propName, VbGet))
        If critColl.Exists(key) Then
            coll_filtered.Add coll(i)
        End If
    Next i
    Set FilterCollOfObj = coll_filtered
End Function

Public Function FilterCollOfObjWithKey(coll As Collection, keys() As String) As Collection
    Dim i As Long
    Dim coll_filtered As New Collection
    For i = LBound(keys) To UBound(keys)
        If isKeyExist(coll, keys(i)) Then
            coll_filtered.Add coll(keys(i))
        End If
    Next i
    Set FilterCollOfObjWithKey = FlattenCollOfColl(coll_filtered)
End Function

Private Function IsInArr(str As Variant, arr As Variant) As Boolean
    Dim i As Long
    If IsArray(arr) Then
        For i = LBound(arr) To UBound(arr)
            If str = arr(i) Then
                IsInArr = True
                Exit Function
            End If
        Next i
    Else
        If str = arr Then
            IsInArr = True
            Exit Function
        End If
    End If
    IsInArr = False
End Function

Private Function CStr_arr(var As Variant) As String()
    Dim arr() As String, i As Long
    If IsArray(var) Then
        ReDim arr(LBound(var) To UBound(var))
        For i = LBound(var) To UBound(var)
            arr(i) = CStr(var(i))
        Next i
    Else
        ReDim arr(0)
        arr(0) = CStr(var)
    End If
    CStr_arr = arr
End Function

Private Function DeDupeOneDimArray(vArray As Variant, Optional isExcludeEmpty As Boolean = True) As Variant
  Dim oDict As Object, i As Long
  Set oDict = CreateObject("Scripting.Dictionary")
  For i = LBound(vArray) To UBound(vArray)
    oDict(vArray(i)) = True
  Next i
  
  On Error Resume Next
  If isExcludeEmpty Then oDict.Remove (vbNullString)
  On Error GoTo 0
  
  DeDupeOneDimArray = oDict.keys()
End Function

Private Sub QuickSortObj(vArray As Variant, inLow As Long, inHi As Long, field As String)
    'sorting class object
    Dim pivot   As Variant
    Dim tmpSwap As Object
    Dim tmpLow  As Long
    Dim tmpHi   As Long
    
    tmpLow = inLow
    tmpHi = inHi
    
    pivot = CallByName(vArray((inLow + inHi) \ 2), field, VbGet)
    
    While (tmpLow <= tmpHi)
    
     While (CallByName(vArray(tmpLow), field, VbGet) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend
    
     While (pivot < CallByName(vArray(tmpHi), field, VbGet) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend
    
     If (tmpLow <= tmpHi) Then
        Set tmpSwap = vArray(tmpLow)
        Set vArray(tmpLow) = vArray(tmpHi)
        Set vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If
    
    Wend
    
    If (inLow < tmpHi) Then QuickSortObj vArray, inLow, tmpHi, field
    If (tmpLow < inHi) Then QuickSortObj vArray, tmpLow, inHi, field

End Sub

'Sub BubbleSort(propName As String, numericSort As Boolean, Optional ascendingOrder As Boolean = True)
'
'    Dim i As Integer
'    Dim j As Integer
'    Dim temp As Object
'
'    With Me
'        For i = 1 To .count - 1
'            For j = i + 1 To .count
'                If (numericSort And ascendingOrder And CallByName(.Item(i), propName, VbGet) > CallByName(.Item(j), propName, VbGet)) _
'                Or (Not (numericSort) And ascendingOrder And StrComp(CallByName(.Item(i), propName, VbGet), CallByName(.Item(j), propName, VbGet)) = 1) _
'                Or (numericSort And Not (ascendingOrder) And CallByName(.Item(i), propName, VbGet) < CallByName(.Item(j), propName, VbGet)) _
'                Or (Not (numericSort) And Not (ascendingOrder) And StrComp(CallByName(.Item(i), propName, VbGet), CallByName(.Item(j), propName, VbGet)) = -1) Then
'                    Set temp = .Item(j)
'                    Call .Remove(j)
'                    Portfolio.Add temp, Before:=i
'                End If
'            Next j
'        Next i
'    End With
'
'End Sub


Private Sub AddColltoColl(coll As Collection, coll2 As Collection)
    Dim i As Long
    For i = 1 To coll2.count
        coll.Add coll2(i)
    Next i
End Sub

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

Private Function GroupCollByProps(coll As Collection, ParamArray propNames() As Variant) As Collection
    'Return as a collection of collection
    Dim i As Long, key As String, j As Long
    Dim coll_grouped As New Collection
    Dim keys As New Collection
    For i = 1 To coll.count
        key = ""
        For j = LBound(propNames) To UBound(propNames)
            key = key & CallByName(coll(i), propNames(j), VbGet)
        Next j
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
    Set GroupCollByProps = coll_grouped
End Function

Private Function isKeyExist(coll As Collection, sKey As String) As Boolean
    'Input shall be colleciton of OBJECT
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

Private Function FlattenCollOfColl(collOfColl As Collection) As Collection
    Dim result As Collection
    Dim coll As Variant
    Dim item As Variant
    
    Set result = New Collection
    
    ' Iterate through each collection in the input collection of collections
    For Each coll In collOfColl
        ' Iterate through each item in the current collection
        For Each item In coll
            ' Add the item to the result collection
            result.Add item
        Next item
    Next coll
    
    Set FlattenCollOfColl = result
End Function


Private Function CDbl_arr(var As Variant) As Double()
    Dim arr() As Double, i As Long
    ReDim arr(LBound(var) To UBound(var))
    For i = LBound(var) To UBound(var)
        arr(i) = CDbl(var(i))
    Next i
    CDbl_arr = arr
End Function


