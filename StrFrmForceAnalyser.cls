VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StrFrmForceAnalyser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Structural Model"

'********************************************************
'This module creates the frame force objects with MaxMinEnv.
'Input = frame force object collection of certain frame
'Output is the permutation of the frame force
'Arthor: Lucas LEUNG
'Update Log

'24 Aug 2023 - Initial
'*******************************************************
Private mModel As StrModel
Private caseDesc As String

Private Enum extremeCaseType
    maxP = 0
    MinP = 1
    MaxV2 = 2
    MinV2 = 3
    MaxV3 = 4
    MinV3 = 5
    MaxT = 6
    MinT = 7
    MaxM2 = 8
    MinM2 = 9
    MaxM3 = 10
    MinM3 = 11
    NoExtremeCase = 12
End Enum

Public Sub Initialize(model As StrModel)
    Set mModel = model
End Sub
Public Function FindExtremeForceWithMaxMinEnv(coll_frmForce As Collection, ExtremeCase() As Boolean, _
                            isMaxMinEnv() As Boolean) As Collection
    'This is the main function for extract correpondence caes.
    'Input: coll_frmForce = collection of frame force objects
    'Logic: 1. locate the pos of the extreme case in the coll_frmForce object
    '       2. Env Max Min, Create Permutation. Then Return the collection of frame force object.
    
    Dim i As Integer
    Dim tempColl As New Collection, reColl As New Collection
    
    For i = LBound(ExtremeCase) To UBound(ExtremeCase)
        If ExtremeCase(i) Then
            Set tempColl = FindFrameForcePair(LocateExtremeForce(coll_frmForce, IndexToExtremeCaseType(i)), IndexToExtremeCaseType(i)) 'Filter out Extreme Force Location
            g_log.WriteLogWithTime "   Start Creating Permutations for Max/Min Force Pair"
            AddColltoColl reColl, CreateFrameForcesPermutations(tempColl, IndexToExtremeCaseType(i), isMaxMinEnv)
            g_log.WriteLogWithTime "   Complete Creating Permutations for Max/Min Force Pair"
        End If
    Next i
    Set FindExtremeForceWithMaxMinEnv = reColl
End Function

Public Function FindExtremeForce(coll_frmForce As Collection, ExtremeCase() As Boolean) As Collection
    'This is the main function for extract correpondence caes.
    'Input: coll_frmForce = collection of frame force objects
    'Logic: 1. locate the pos of the extreme case in the coll_frmForce object
    '       2. Env Max Min, Create Permutation. Then Return the collection of frame force object.
    
    Dim i As Integer
    Dim tempColl As New Collection
    Dim objForce As StrFrameForce, maxForce As StrFrameForce
    'Form frm Force Obj from the returned 2D array
    
    For i = LBound(ExtremeCase) To UBound(ExtremeCase)
        If ExtremeCase(i) Then
            Set objForce = New StrFrameForce
            Set maxForce = LocateExtremeForce(coll_frmForce, IndexToExtremeCaseType(i))
            With maxForce
                objForce.Init2 .frame, .station, .loadComb, .stepType, _
                            .p, .V2, .V3, .t, .M2, .M3
            End With
            objForce.extremeCaseType = maxForce.extremeCaseType
            tempColl.Add objForce
        End If
    Next i
    Set FindExtremeForce = tempColl
End Function

Private Function LocateExtremeForce(frmForces As Collection, enum_extremeCaseType As extremeCaseType) As StrFrameForce
    'this function is used to locate the extreme case of frame force objects
    'return should be 1/2 nos of frame force objects (1 if that is not ENV load comb, 2 = env load comb, max and min rows
    Dim i As Long
    Dim forceIndex As Integer, isMax As Boolean
    Dim cExtremeForce As StrFrameForce ', cFrmForce As New Collection
    Dim frmForce As StrFrameForce
    forceIndex = ExtremeCaseToForceIndex(enum_extremeCaseType)

    If enum_extremeCaseType Mod 2 = 0 Then
        isMax = True
    Else
        isMax = False
    End If
    
    'to find out the extreme force (Max P, Min P, etc...) from a collection of frm forces (all frm forces)
    Set cExtremeForce = frmForces(1)

    g_log.WriteLogWithTime GetExtremeCaseName(enum_extremeCaseType)
    If isMax Then
        For Each frmForce In frmForces
            If frmForce.force(forceIndex) > cExtremeForce.force(forceIndex) Then Set cExtremeForce = frmForce

        Next
    Else
        For Each frmForce In frmForces
            If frmForce.force(forceIndex) < cExtremeForce.force(forceIndex) Then Set cExtremeForce = frmForce
            
        Next
    End If

    
    g_log.WriteLogWithTime "   " & GetExtremeCaseName(enum_extremeCaseType) & " value:" & cExtremeForce.force(forceIndex)
    g_log.WriteLog "   Frame Name: " & cExtremeForce.frameName & _
                    ", Sub Element Name: " & cExtremeForce.subFrameName & _
                    ", station: " & cExtremeForce.pos_fromEleJtI
    
'    cFrmForce.Add cExtremeForce
'    Dim nextFrmForce As StrFrameForce
'    If Not locAtExtreme = frmForces.count Then
'        Set nextFrmForce = frmForces(locAtExtreme + 1)
'        If cExtremeForce.frameName = nextFrmForce.frameName And cExtremeForce.station = nextFrmForce.station And cExtremeForce.stepType <> nextFrmForce.station Then
'            cFrmForce.Add nextFrmForce
'        End If
'    End If
    'to find out the
'    With cExtremeForce
'        Set cFrmForce = mModel.frmForcesOneLcAtPos(.subFrameName, .loadComb, .pos_fromMemJtI_percent)
'        If cFrmForce.count > 2 Then g_log.RaiseWarning "WARNING: Please check your data. There are more than 2 results at the same location of an element.", unexpectedResult
'    End With
    

    cExtremeForce.extremeCaseType = GetExtremeCaseName(enum_extremeCaseType)

    
    Set LocateExtremeForce = cExtremeForce
End Function

Private Function FindFrameForcePair(frmForce As StrFrameForce, enum_extremeCaseType As extremeCaseType) As Collection
    Dim coll As New Collection
    With frmForce
        Set coll = mModel.frmForcesOneLcAtPos(.subFrameName, .loadComb, .pos_fromMemJtI_percent)
        If coll.count > 2 Then g_log.RaiseWarning "WARNING: Please check your data. There are more than 2 results at the same location of an element.", unexpectedResult
    End With
    
    For i = 1 To coll.count
        coll(i).extremeCaseType = GetExtremeCaseName(enum_extremeCaseType)
    Next i
    Set FindFrameForcePair = coll
End Function
Private Function CreateFrameForcesPermutations(frmForceAtPos As Collection, enum_extremeCaseType As extremeCaseType, isMaxMinEnv() As Boolean) As Collection
    'This function input = frm force at of certain frame at certain position and certain load comb. (1 or 2 objs)
    'output = Permuation of forces
    'maxMinEnv(0 to 5) represent the option to env max/min force pair for the P/V2/V3/T/M2/M3
    'True = Env (Get the ABS Max Result) --> suitable when the direction of the force is not important
    'Combinations will be created for all the 'false' options.
    Dim i As Integer
    Dim maxForce As StrFrameForce, minForce As StrFrameForce
    Dim forces_Origin As Variant 'Array of array, to contain the forces for creation of permutation
    Dim forces_Processed As Variant
    Dim objForce As StrFrameForce
    Dim coll_objForce As New Collection
    Dim caseDesc() As String
    
    If frmForceAtPos.count = 1 Then 'for data without max/min (load case/ not env load comb)
        Set objForce = New StrFrameForce
        'frmForceAtPos(1).extremeCaseType = GetExtremeCaseName(enum_extremeCaseType)
        With frmForceAtPos(1)
            objForce.Init2 .frame, .station, .loadComb, "", .p, .V2, .V3, .t, .M2, .M3, .subFrameName
        End With
        objForce.extremeCaseType = GetExtremeCaseName(enum_extremeCaseType)
        coll_objForce.Add objForce
        Set CreateFrameForcesPermutations = coll_objForce
        Exit Function
    End If
    
    If UCase(frmForceAtPos(1).stepType) = "MAX" Then
        Set maxForce = frmForceAtPos(1)
        Set minForce = frmForceAtPos(2)
    Else
        Set maxForce = frmForceAtPos(2)
        Set minForce = frmForceAtPos(1)
    End If
    
    'Get the force array for creation of permutation
    ReDim forces_Origin(5)
    For i = 0 To 5
        If ExtremeCaseToForceIndex(enum_extremeCaseType) = i Then
            If ExtremeCaseToIsMax(enum_extremeCaseType) Then
                forces_Origin(i) = Array(maxForce.force(i))
            Else
                forces_Origin(i) = Array(minForce.force(i))
            End If
        ElseIf isMaxMinEnv(i) Then
            If Abs(maxForce.force(i)) > Abs(minForce.force(i)) Then
                forces_Origin(i) = Array(maxForce.force(i))
            Else
                forces_Origin(i) = Array(minForce.force(i))
            End If
        Else
            forces_Origin(i) = Array(maxForce.force(i), minForce.force(i))
        End If
    Next i
    
    forces_Processed = Permutation(forces_Origin, caseDesc)
    
    'Form frm Force Obj from the returned 2D array
    For i = 0 To UBound(forces_Processed)
        Set objForce = New StrFrameForce
        objForce.Init2 maxForce.frame, maxForce.station, maxForce.loadComb, caseDesc(i), _
                        CDbl(forces_Processed(i, 0)), CDbl(forces_Processed(i, 1)), CDbl(forces_Processed(i, 2)), _
                        CDbl(forces_Processed(i, 3)), CDbl(forces_Processed(i, 4)), CDbl(forces_Processed(i, 5)), maxForce.subFrameName
        objForce.extremeCaseType = GetExtremeCaseName(enum_extremeCaseType)
        coll_objForce.Add objForce
    Next i
    
    Set CreateFrameForcesPermutations = coll_objForce
End Function

Public Function EnvMaxMin(frmForceAtPos As Collection, isMaxMinEnv() As Boolean) As Collection
    'This function input = frm force at of certain frame at certain position and certain load comb. (1 or 2 objs)
    'output = Permuation of forces
    'maxMinEnv(0 to 5) represent the option to env max/min force pair for the P/V2/V3/T/M2/M3
    'True = Env (Get the ABS Max Result) --> suitable when the direction of the force is not important
    'Combinations will be created for all the 'false' options.
    Dim i As Integer
    Dim maxForce As StrFrameForce, minForce As StrFrameForce
    Dim forces_Origin As Variant 'Array of array, to contain the forces for creation of permutation
    Dim forces_Processed As Variant
    Dim objForce As StrFrameForce
    Dim coll_objForce As New Collection
    Dim caseDesc() As String
    
    If frmForceAtPos.count = 1 Then 'for data without max/min (load case/ not env load comb)
        Set EnvMaxMin = frmForceAtPos
        
'        'frmForceAtPos(1).extremeCaseType = GetExtremeCaseName(enum_extremeCaseType)
'        With frmForceAtPos(1)
'            objForce.Init2 .frame, .station, .loadcomb, "", .p, .V2, .V3, .t, .M2, .M3
'        End With
'
'        coll_objForce.Add objForce
'        Set EnvMaxMinWithExtremeCase = coll_objForce
        Exit Function
    End If
    
    If UCase(frmForceAtPos(1).stepType) = "MAX" Then
        Set maxForce = frmForceAtPos(1)
        Set minForce = frmForceAtPos(2)
    Else
        Set maxForce = frmForceAtPos(2)
        Set minForce = frmForceAtPos(1)
    End If
    
    'Get the force array for creation of permutation
    ReDim forces_Origin(5)
    For i = 0 To 5
        If isMaxMinEnv(i) Then
            If Abs(maxForce.force(i)) > Abs(minForce.force(i)) Then
                forces_Origin(i) = Array(maxForce.force(i))
            Else
                forces_Origin(i) = Array(minForce.force(i))
            End If
        Else
            forces_Origin(i) = Array(maxForce.force(i), minForce.force(i))
        End If
    Next i
    
    forces_Processed = Permutation(forces_Origin, caseDesc)
    
    'Form frm Force Obj from the returned 2D array
    For i = 0 To UBound(forces_Processed)
        Set objForce = New StrFrameForce
        objForce.Init2 maxForce.frame, maxForce.station, maxForce.loadComb, caseDesc(i), _
                        CDbl(forces_Processed(i, 0)), CDbl(forces_Processed(i, 1)), CDbl(forces_Processed(i, 2)), _
                        CDbl(forces_Processed(i, 3)), CDbl(forces_Processed(i, 4)), CDbl(forces_Processed(i, 5))
        coll_objForce.Add objForce
    Next i
    
    Set EnvMaxMin = coll_objForce
End Function


Private Function Permutation(InputArr As Variant, caseDesc() As String) As Variant
    'input as array of array. 1st layer of array = groups, 2nd array = items in the same group
    Dim arr As Variant, numOfArr As Long 'arr = InputArr(Base 1), numOfArr = num of group
    Dim index() As Long 'array index
    Dim num() As Long 'UBound of each item in a group
    Dim totalNum As Long 'total number of permutation
    Dim returnArr As Variant
    Dim i As Long, j As Long
    
    'Change arr from base 0 to base 1
    ReDim arr(1 To UBound(InputArr) + 1)
    numOfArr = UBound(arr)
    For i = 0 To UBound(InputArr)
        arr(i + 1) = ChangeArrBase(InputArr(i), 1)
    Next i
    
    ReDim index(1 To numOfArr)
    ReDim num(1 To numOfArr)
    
    totalNum = 1
    
    For i = 1 To numOfArr
        num(i) = UBound(arr(i))
        totalNum = totalNum * num(i)
        index(i) = 1
        
    Next i
    index(numOfArr) = 0
    
    ReDim returnArr(0 To totalNum - 1, 0 To numOfArr - 1)
    ReDim caseDesc(0 To totalNum - 1)
    
    For i = 1 To totalNum
        'caseDesc(i - 1) = "MaxMinEnv-"
        
        'assign the index of array
        For j = numOfArr To 1 Step -1
            index(j) = index(j) + 1
            If index(j) > num(j) Then
                index(j) = 1
            Else
                Exit For
            End If
        Next j
        
        'assign the value to return arr
        For j = 1 To numOfArr
            returnArr(i - 1, j - 1) = arr(j)(index(j))
            
            If num(j) = 2 Then
                If index(j) = 1 Then
                    caseDesc(i - 1) = "MaxMinEnv-" & "Max" & GetForceString(j) & " "
                Else
                    caseDesc(i - 1) = "MaxMinEnv-" & "Min" & GetForceString(j) & " "
                End If
            End If
        Next j
    Next i
    
    Permutation = returnArr
End Function

Private Function ExtremeCaseToForceIndex(enum_extremeCaseType As extremeCaseType) As Integer
    Dim forceIndex As Integer
    Select Case enum_extremeCaseType
        Case maxP: forceIndex = 0
        Case MinP: forceIndex = 0
        Case MaxV2: forceIndex = 1
        Case MinV2: forceIndex = 1
        Case MaxV3: forceIndex = 2
        Case MinV3: forceIndex = 2
        Case MaxT: forceIndex = 3
        Case MinT: forceIndex = 3
        Case MaxM2: forceIndex = 4
        Case MinM2: forceIndex = 4
        Case MaxM3: forceIndex = 5
        Case MinM3: forceIndex = 5
        Case NoExtremeCase: forceIndex = 11
    End Select
    ExtremeCaseToForceIndex = forceIndex
End Function

Private Function ExtremeCaseToIsMax(enum_extremeCaseType As extremeCaseType) As Boolean
    Dim isMax As Boolean
    If enum_extremeCaseType Mod 2 = 0 Then
        isMax = True
    Else
        isMax = False
    End If
    ExtremeCaseToIsMax = isMax
End Function

Private Function GetForceString(i As Long) As String
    Select Case i
        Case 1: GetForceString = "P"
        Case 2: GetForceString = "V2"
        Case 3: GetForceString = "V3"
        Case 4: GetForceString = "T"
        Case 5: GetForceString = "M2"
        Case 6: GetForceString = "M3"
    End Select
End Function


Private Function GetExtremeCaseName(enum_extremeCaseType As extremeCaseType) As String
    Select Case enum_extremeCaseType
        Case maxP: GetExtremeCaseName = "Max P"
        Case MinP: GetExtremeCaseName = "Min P"
        Case MaxV2: GetExtremeCaseName = "Max V2"
        Case MinV2: GetExtremeCaseName = "Min V2"
        Case MaxV3: GetExtremeCaseName = "Max V3"
        Case MinV3: GetExtremeCaseName = "Min V3"
        Case MaxT: GetExtremeCaseName = "Max T"
        Case MinT: GetExtremeCaseName = "Min T"
        Case MaxM2: GetExtremeCaseName = "Max M2"
        Case MinM2: GetExtremeCaseName = "Min M2"
        Case MaxM3: GetExtremeCaseName = "Max M3"
        Case MinM3: GetExtremeCaseName = "Min M3"
    End Select
End Function

Private Function IndexToExtremeCaseType(index As Integer) As extremeCaseType
    Select Case index
        Case 0: IndexToExtremeCaseType = maxP
        Case 1: IndexToExtremeCaseType = MinP
        Case 2: IndexToExtremeCaseType = MaxV2
        Case 3: IndexToExtremeCaseType = MinV2
        Case 4: IndexToExtremeCaseType = MaxV3
        Case 5: IndexToExtremeCaseType = MinV3
        Case 6: IndexToExtremeCaseType = MaxT
        Case 7: IndexToExtremeCaseType = MinT
        Case 8: IndexToExtremeCaseType = MaxM2
        Case 9: IndexToExtremeCaseType = MinM2
        Case 10: IndexToExtremeCaseType = MaxM3
        Case 11: IndexToExtremeCaseType = MinM3
    End Select
End Function

Private Function ChangeArrBase(arr As Variant, base As Integer) As Variant
    Dim i As Long, tempArr As Variant
    Dim n As Long, count As Long
    n = UBound(arr) - LBound(arr) + 1
    ReDim tempArr(base To base + n - 1)
    For i = LBound(arr) To UBound(arr)
        tempArr(base + count) = arr(i)
        count = count + 1
    Next i
    ChangeArrBase = tempArr
End Function
Private Sub AddColltoColl(coll As Collection, coll2 As Collection)
    Dim i As Long
    For i = 1 To coll2.count
        coll.Add coll2(i)
    Next i
End Sub


