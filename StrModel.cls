VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StrModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Structural Model"

'********************************************************
'This module contains all structural objects
'Arthor: Lucas LEUNG
'Update Log

'07 Jun 2023 - Initial
'*******************************************************

Option Explicit

Private isRenumberNameToID As Boolean

Private mJts As New Collection, mJts_ID As New Collection 'for pJts, the key = name of the object (string), ID is long
Private mFrms As New Collection, mFrms_ID As New Collection
Private mFrmsBySection As New Collection, mFrmsByMember As New Collection 'Different Key for filter use. Collection of Collection
Private mLinks As New Collection, mLinks_ID As New Collection
Private mShells As New Collection, mShells_ID As New Collection
Private mMembers As New Collection
Private mConns As New Collection
Private mJtLoads As New Collection
Private mFrmLoads As New Collection
Private mLoadPats As New Collection, mLoadPats_ID As New Collection
Private mLoadCases As New Collection, mLoadCases_ID As New Collection
Private mLoadCombs As New Collection, mLoadCombs_ID As New Collection

Private pID_joint As Long, pID_ele As Long, pID_loadPat As Long

'For Stroing and Extraction of Forces
Private mFrmForces As New Collection, mFrmForcesBySection As New Collection
Private mFrmForcesByMember As New Collection, mFrmForcesByLc As New Collection
Private mFrmForceAnalyser As StrFrmForceAnalyser
Private mStrModelConstructor As StrModelConstructor

'Private pBorelogs() As clsBoreLog
'Private id_joint(3) As Long, id_frame(3) As Long
'
'Private pBridges() As clsBridge
'Private pPiles() As clsPile

'Private pIsExportSucess As Boolean
'Object = Have Name and ID. Share similar method in conversion of name/ID/Identification

Public Enum StructureObjectType
    obj_jt = 1
    obj_frm = 2
    obj_link = 3
    obj_shell = 4
    obj_loadPat = 5
    obj_loadCase = 6
    obj_loadComb = 7
    obj_member = 8
    obj_connection = 9
    obj_jtLoad = 11
    obj_frmLoad = 12
    obj_frmTemperatureLoad = 13
    obj_frmGravityLoad = 14
    obj_frmForce = 21
End Enum


Private pID() As Long
Private Enum idType
    jt = 0
    frm = 1
    link = 2
End Enum
Private Sub Class_Initialize()
    pID_joint = 0
    pID_ele = 0
    pID_loadPat = 0
    isRenumberNameToID = True
End Sub

Public Function GetID(cStrObjType As StructureObjectType) As Long
    'This Sub to Get Unique Joint name
    If cStrObjType = obj_jt Then
        pID_joint = pID_joint + 1
        GetID = pID_joint
    ElseIf cStrObjType = obj_frm Or cStrObjType = obj_link Then
        pID_ele = pID_ele + 1
        GetID = pID_ele
    ElseIf cStrObjType = obj_loadPat Then
        pID_loadPat = pID_loadPat + 1
        GetID = pID_loadPat
    End If
End Function

'Public Sub ExportGWA()
'
'    If isRenumberNameToID Then
'        RenumberNametoID_AllObjs
'    Else
'        CopyNametoID_AllObjs
'    End If
'
'    'On Error GoTo Err
'    Dim writer As clsStrModelExportGSA
'    Set writer = New clsStrModelExportGSA
'    writer.Init Me
'    writer.WriteFile
'    pIsExportSucess = True
'Err:
'End Sub

'Private Sub RenumberNametoID_AllObjs()
'    'Convert all the joints and frame to suit for model input (eg. GSA/ MnameAS, need integer)
'    RenumberNametoID obj_jt
'    RenumberNametoID obj_frm
'    RenumberNametoID obj_link
'    RenumberNametoID obj_loadPat
'
'End Sub

'Private Sub RenumberNametoID(cStrObjType As StructureObjectType)
'    'Convert all the joints and frame to suit for model input (eg. GSA/ MnameAS, need integer)
'    Dim i As Long
'    Dim cColl As Object, cColl_ID As Object
'    Dim id As Long
'
'    Set cColl = GetStrObjColl(cStrObjType)
'    Set cColl_ID = GetStrObjColl(cStrObjType, False)
'    For i = 1 To cColl.count
'        With cColl.Item(i)
'            If .id = 0 Then
'                id = GetID(cStrObjType)
'            Else
'                id = .id
'            End If
'            If Me.IsStrObjectExist(CStr(id), cStrObjType, False) Then
'                Dim strTypeName As String
'                strTypeName = GetStrObjTypeName(cStrObjType)
'
'                Do Until Not Me.IsStrObjectExist(CStr(id), cStrObjType, False)
'                    id = GetID(cStrObjType)
'                Loop
'                g_log.RaiseWarning "Fail to convert " & strTypeName & " '" & cColl.Item(i).name _
'                        & "' to ID. The ID is renumberred to '" & id & "'.", failToConvertNameToID
'            End If
'            .id = id
'            cColl_ID.Add cColl.Item(i), CStr(id)
'        End With
'    Next i
'End Sub
'
'Private Sub CopyNametoID_AllObjs()
'
'    CopyNametoID obj_jt
'    CopyNametoID obj_frm
'    CopyNametoID obj_link
'    CopyNametoID obj_loadPat
'
'End Sub

'Private Sub CopyNametoID(cStrObjType As StructureObjectType)
'
'    Dim i As Long
'    Dim cColl As Collection, cColl_ID As Object
'
'    On Error GoTo Err
'
'    Set cColl = GetStrObjColl(cStrObjType, True)
'    Set cColl_ID = GetStrObjColl(cStrObjType, False)
'    For i = 1 To cColl.count
'        With cColl.Item(i)
'            If .id = 0 Then .id = CLng(.name)
'            cColl_ID.Add cColl.Item(i), CStr(.id)
'        End With
'    Next i
'    Exit Sub
'Err:
'    Dim id As Long, strTypeName As String
'    'With cColl.Item(i)
'    id = GetID(cStrObjType)
'    Do Until Not Me.IsStrObjectExist(CStr(id), cStrObjType, False)
'        id = GetID(cStrObjType)
'    Loop
'    cColl.Item(i).id = id
'    strTypeName = GetStrObjTypeName(cStrObjType)
'    g_log.RaiseWarning "Fail to convert " & strTypeName & " '" & cColl.Item(i).name _
'                        & "' to ID. The ID is renumberred to '" & id & "'.", failToConvertNameToID
'    'End If
'    Resume Next
'End Sub

'Public Property Get isExportSucess() As Boolean
'    isExportSucess = pIsExportSucess
'End Property

Public Property Get joints(Optional nodeList As String) As Object
    If nodeList = vbNullString Then
        Set joints = mJts
    End If
End Property

Public Property Get frames(Optional sectionFilter As Variant, Optional memberFilter As Variant) As Object
    
    Dim coll As New Collection
    Set coll = mFrms
    If isInitialised(sectionFilter) Then
        Set coll = FilterCollOfObj(mFrms, "section", CStr_arr(sectionFilter))
    End If
    
    If isInitialised(memberFilter) Then
        Set coll = FilterCollOfObj(coll, "memberName", CStr_arr(memberFilter))
    End If
    
'    If isInitialised(sectionFilter) And Not isInitialised(memberFilter) Then
'        Set coll = FilterCollOfObjWithKey(mFrmsBySection, CStr_arr(sectionFilter))
'    ElseIf isInitialised(memberFilter) And Not isInitialised(sectionFilter) Then
'        Set coll = FilterCollOfObjWithKey(mFrmsByMember, CStr_arr(memberFilter))
'    ElseIf isInitialised(memberFilter) And isInitialised(sectionFilter) Then
'        Set coll = FilterCollOfObjWithKey(mFrmsBySection, CStr_arr(memberFilter))
'        Dim coll2 As New Collection
'        Set coll2 = GroupCollByProp(coll, "memberName")
'        Set coll = FilterCollOfObjWithKey(coll2, CStr_arr(sectionFilter))
'    Else
'        Set coll = mFrms
'    End If
    
    Set frames = coll
End Property

Public Property Get frmsBySection() As Collection
    If mFrmsBySection.count = 0 Then
        Set mFrmsBySection = GroupCollByProp(mFrms, "section")
    End If
    Set frmsBySection = mFrmsBySection
End Property

Public Property Get frmsByMember() As Collection
    If mFrmsByMember.count = 0 Then
        Set mFrmsByMember = GroupCollByProp(mFrms, "memberName")
    End If
    Set frmsByMember = mFrmsByMember
End Property

Public Property Get links() As Object
    Set links = mLinks
End Property

Public Property Get member(Name As String) As Object
    Set member = mMembers(Name)
 End Property
 
Public Property Get members() As Object
    Set members = mMembers
 End Property

Public Property Set members(coll_members As Object)
    Set mMembers = coll_members
End Property

Public Property Get conns() As Object
    Set conns = mConns
End Property
Public Property Get loadPats() As Object
    Set loadPats = mLoadPats
End Property

Public Property Get jtLoads() As Object
    Set jtLoads = mJtLoads
End Property

Public Property Get frmLoads() As Object
    Set frmLoads = mFrmLoads
End Property

Public Property Get frmForces() As Object
    Set frmForces = mFrmForces
End Property

'***********************Structural Model Object Base Methods and Functions***************************************
Public Function GetStrObjTypeName(strObjectType As StructureObjectType) As String
    Select Case strObjectType
        Case obj_jt: GetStrObjTypeName = "Joint"
        Case obj_frm: GetStrObjTypeName = "Frame"
        Case obj_link: GetStrObjTypeName = "Link"
        Case obj_shell: GetStrObjTypeName = "Shell"
        Case obj_member: GetStrObjTypeName = "Member"
        Case obj_connection: GetStrObjTypeName = "Connection"
        Case obj_loadPat: GetStrObjTypeName = "Load Pattern"
        Case obj_loadCase: GetStrObjTypeName = "Load Case"
        Case obj_loadComb: GetStrObjTypeName = "Load Combination"
        Case obj_jtLoad: GetStrObjTypeName = "Joint Load"
        Case obj_frmLoad: GetStrObjTypeName = "Frame Distributed Load"
        Case obj_frmTemperatureLoad: GetStrObjTypeName = "Frame Temperature Load"
        Case obj_frmGravityLoad: GetStrObjTypeName = "Frame Gravity Load"
    End Select
End Function

Public Function GetStrObjColl(strObjectType As StructureObjectType, Optional isByName As Boolean = True) As Collection
    'Get all objects of certain type of structural object
    If isByName Then
        Select Case strObjectType
            Case obj_jt: Set GetStrObjColl = mJts
            Case obj_frm: Set GetStrObjColl = mFrms
            Case obj_link: Set GetStrObjColl = mLinks
            Case obj_shell: Set GetStrObjColl = mShells
            Case obj_member: Set GetStrObjColl = mMembers
            Case obj_connection: Set GetStrObjColl = mConns
            Case obj_loadPat: Set GetStrObjColl = mLoadPats
            Case obj_loadCase: Set GetStrObjColl = mLoadCases
            Case obj_loadComb: Set GetStrObjColl = mLoadCombs
            Case obj_jtLoad: Set GetStrObjColl = mJtLoads
            Case obj_frmLoad: Set GetStrObjColl = mFrmLoads
            Case obj_frmTemperatureLoad: Set GetStrObjColl = mFrmLoads
            Case obj_frmGravityLoad: Set GetStrObjColl = mFrmLoads
            Case obj_frmForce: Set GetStrObjColl = mFrmForces
        End Select
    Else
        Select Case strObjectType
            Case obj_jt: Set GetStrObjColl = mJts_ID
            Case obj_frm: Set GetStrObjColl = mFrms_ID
            Case obj_link: Set GetStrObjColl = mLinks_ID
            Case obj_shell: Set GetStrObjColl = mShells_ID
            Case obj_loadPat: Set GetStrObjColl = mLoadPats_ID
            Case obj_loadCase: Set GetStrObjColl = mLoadCases_ID
            Case obj_loadComb: Set GetStrObjColl = mLoadCombs_ID
        End Select
    End If
End Function

Public Function SetNewStrObj(strObjectType As StructureObjectType) As Object
    Dim obj As Object
    Select Case strObjectType
        Case obj_jt: Set obj = New StrJoint
        Case obj_frm: Set obj = New StrFrame
        'Case obj_link: Set obj = New clsStrLink
        'Case obj_shell: Set obj = New clsStrJoint
        Case obj_member: Set obj = New StrMember
        'Case obj_connection: Set obj = New clsStrConnection
        'Case obj_loadCase: Set obj = New clsStrJoint
        'Case obj_loadComb: Set obj = New clsStrJoint
        
    End Select
    Set SetNewStrObj = obj
End Function

Public Function AddStrObjToColl(obj As Object, strObjectType As StructureObjectType) As Integer
    Dim coll As Collection
    Dim ret As Integer
    Set coll = GetStrObjColl(strObjectType)
    'Debug.Print CallByName(obj, "name", VbGet)
    If strObjectType < obj_jtLoad Then
        If isKeyExist(coll, CallByName(obj, "name", VbGet)) Then
            ret = -1
        Else
            coll.Add obj, CallByName(obj, "name", VbGet)
        End If
    Else
        coll.Add obj
    End If
    AddStrObjToColl = ret
End Function

Public Function GetStrObject(sKey As String, strObjectType As StructureObjectType) As Object
    'Get an object by name properties
    Dim coll As Collection
    Set coll = GetStrObjColl(strObjectType)
    Set GetStrObject = GetObjectFromCollection(coll, sKey)
End Function

Public Function IsStrObjectExist(sKey As String, strObjectType As StructureObjectType, Optional isByName As Boolean = True) As Boolean
    IsStrObjectExist = isKeyExist(GetStrObjColl(strObjectType, isByName), sKey)
End Function

Public Function GetDataframe(strObjectType As StructureObjectType, ParamArray propName() As Variant) As clsDataFrame
    'Get the dataframe representing the data of certain type of obj.
    Dim i As Long, j As Long
    Dim heads() As String, data As Variant
    Dim coll As Collection
    
    
    Set coll = GetStrObjColl(strObjectType)
    ReDim heads(1 To UBound(propName) + 1)
    ReDim data(1 To coll.count, 1 To UBound(propName) + 1)
    
    For j = 0 To UBound(propName)
        heads(j + 1) = propName(j)
    Next j
    
    For i = 1 To coll.count
        For j = 1 To UBound(heads)
            data(i, j) = CallByName(coll(i), heads(j), VbGet)
        Next j
    Next i
    
    Dim df As clsDataFrame
    Set df = New clsDataFrame
    df.Init_byArr data, False, False
    df.heads = heads
    Set GetDataframe = df
End Function

Public Function GetDataframe_fromColl(coll As Collection, ParamArray propName() As Variant) As clsDataFrame
    'Get the dataframe representing the data of certain type of obj.
    Dim i As Long, j As Long
    Dim heads() As String, data As Variant
    
    ReDim heads(1 To UBound(propName) + 1)
    ReDim data(1 To coll.count, 1 To UBound(propName) + 1)
    
    For j = 0 To UBound(propName)
        heads(j + 1) = propName(j)
    Next j
    
    For i = 1 To coll.count
        For j = 1 To UBound(heads)
            data(i, j) = CallByName(coll(i), heads(j), VbGet)
        Next j
    Next i
    
    Dim df As clsDataFrame
    Set df = New clsDataFrame
    df.Init_byArr data, False, False
    df.heads = heads
    Set GetDataframe_fromColl = df
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

        key = CallByName(coll(i), propName, VbGet)
        If critColl.Exists(key) Then
            coll_filtered.Add coll(i)
        End If
    Next i
    Set FilterCollOfObj = coll_filtered
End Function
'
'Public Function FilterCollOfObjWithKey(coll As Collection, keys() As String) As Collection
'    Dim i As Long
'    Dim coll_filtered As New Collection
'    For i = LBound(keys) To UBound(keys)
'        If isKeyExist(coll, keys(i)) Then
'            coll_filtered.Add coll(keys(i))
'        End If
'    Next i
'    Set FilterCollOfObjWithKey = FlattenCollOfColl(coll_filtered)
'End Function

'***********************Functions for Collection**************************************
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
    'Input shall be collection
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

Private Function GetObjectFromCollection(coll As Collection, sKey As String) As Object
    'return nothing if not found
    Dim obj As Object
    On Error Resume Next
       
    'then read by name
    Set obj = coll(sKey)
    If Err.Number = 0 Then
        Set GetObjectFromCollection = obj
        Exit Function
    End If
    
    Set GetObjectFromCollection = Nothing
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

'Private Function FlattenCollOfColl(collOfColl As Collection) As Collection
'    Dim result As Collection
'    Dim coll As Variant
'    Dim item As Variant
'
'    Set result = New Collection
'
'    ' Iterate through each collection in the input collection of collections
'    For Each coll In collOfColl
'        ' Iterate through each item in the current collection
'        For Each item In coll
'            ' Add the item to the result collection
'            result.Add item
'        Next item
'    Next coll
'
'    Set FlattenCollOfColl = result
'End Function
'*************For Force Extraction & Analysis*******************
Property Get frmForceAnalyser() As StrFrmForceAnalyser
    If mFrmForceAnalyser Is Nothing Then Set mFrmForceAnalyser = New StrFrmForceAnalyser
    Set frmForceAnalyser = mFrmForceAnalyser
End Property

Property Get Constructor() As StrModelConstructor
    If mStrModelConstructor Is Nothing Then
        Set mStrModelConstructor = New StrModelConstructor
        mStrModelConstructor.Initialize Me
    End If
    Set Constructor = mStrModelConstructor
End Property


