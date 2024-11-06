VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StrJoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Structural Model"

'********************************************************
'This module represent the Structural Joint/Node Object
'Arthor: Lucas LEUNG
'Update Log

'04 Apr 2023 - Initial
'*******************************************************

Option Explicit
Private mName As String 'the ID used in the model
Private mJtID As Long 'the ID used in the WS
Private mx As Double, my As Double, mz As Double
Private mLocalAxis As Double
Private mConstraint As Collection 'collection of joint object
Private mConnectedFrms As New Collection 'collection of frame connected to the jt
Private mConnectionsTypes As New Collection
Private mIsRestraint As Boolean, mIsConn As Boolean
'Private pSupportType As String 'Pin/Fixed/etc

Private Enum DefSupportType
    Roller = 1
    Pinned = 2
    Fixed = 3
End Enum

'Private Enum SupportDir
'    L1 = 1
'    L2 = 2
'    L3 = 3
'    R1 = 4
'    R2 = 5
'    R3 = 6
'End Enum

Private mRestraint(1 To 6) As Boolean
'pRestraint(1) = U1
'pRestraint(2) = U2
'pRestraint(3) = U3
'pRestraint(4) = R1
'pRestraint(5) = R2
'pRestraint(6) = R3

Private mSpring(1 To 6) As Double
'pSpring(1) = U1 [F/L]
'pSpring(2) = U2 [F/L]
'pSpring(3) = U3 [F/L]
'pSpring(4) = R1 [FL/rad]
'pSpring(5) = R2 [FL/rad]
'pSpring(6) = R3 [FL/rad]

Private mJtLoads As New Collection

'Private pDistLoadPat As String
'Private pDistLoad(6) As Double

Public Sub Init(name As String, Optional x As Double = 0, Optional y As Double = 0, Optional z As Double = 0, Optional ID As Long = 0)
    mName = name
    mx = x
    my = y
    mz = z
    mJtID = ID
End Sub
Private Function GetDefSupportTypeVal(str) As DefSupportType
    str = UCase(str)
    Select Case str
        Case "ROLLER": GetDefSupportTypeVal = Roller
        Case "PINNED": GetDefSupportTypeVal = Pinned
        Case "PIN": GetDefSupportTypeVal = Pinned
        Case "FIXED": GetDefSupportTypeVal = Fixed
        Case "FIX": GetDefSupportTypeVal = Fixed
    End Select
End Function

'Private Function GetSupportDirVal(str) As SupportDir
'    Select Case str
'        Case "L1": GetSupportDirVal = L1
'        Case "L2": GetSupportDirVal = L2
'        Case "L3": GetSupportDirVal = L3
'        Case "R1": GetSupportDirVal = R1
'        Case "R2": GetSupportDirVal = R2
'        Case "R3": GetSupportDirVal = R3
'    End Select
'End Function

'Sub Init_Pile(name As String, x As Double, y As Double, z As Double, Optional supportType As String = "", _
'         Optional constraint As clsStrJoint, Optional Spring1 As Double = 0, Optional Spring2 As Double = 0, _
'         Optional localAxis As Double = 0, Optional distLoadPat As String = "", Optional distLoadZ As Double = 0, Optional Name As String = "")
'    Dim i As Long
'    For i = 1 To 6
'        pSpring(i) = 0
'        pRestraint(i) = False
'        pDistLoad(i) = 0
'    Next i
'
'    pName = name
'    px = x
'    py = y
'    pz = z
'    pConstraint = constraint 'to be modified
'    'pSupportType = supportType
'    pLocalAxis = localAxis
'    pDistLoadPat = distLoadPat 'to be modified
'    pDistLoad(3) = distLoadZ 'to be modified
'    pID = Name
'
'    pSpring(1) = Spring1
'    pSpring(2) = Spring2
'
'    'to be modified
''    If pSupportType <> "" Then
''        Select Case pSupportType
''            Case "Fixed"
''                pRestraint(1) = True
''                pRestraint(2) = True
''                pRestraint(3) = True
''                pRestraint(4) = True
''                pRestraint(5) = True
''                pRestraint(6) = True
''            Case "Pinned"
''                pRestraint(1) = True
''                pRestraint(2) = True
''                pRestraint(3) = True
''                pRestraint(4) = True
''            Case "Roller"
''                pRestraint(3) = True
''            Case "Horizontal Load Only"
''                pRestraint(1) = True
''                pRestraint(2) = True
''        End Select
''    End If
'End Sub

Public Sub Init_all()

End Sub

Public Property Get name() As String
    name = mName
End Property

Public Property Let name(str As String)
    mName = str
End Property

Public Property Get ID() As Long
    ID = mJtID
End Property

Public Property Let ID(index As Long)
    mJtID = index
End Property
Public Property Get isHvSpr() As Boolean
    Dim i As Integer, spr As Boolean
    For i = 1 To 6
        If Not mSpring(i) = 0 Then
            spr = True
            Exit For
        End If
    Next i
    isHvSpr = spr
End Property


Public Property Get springArr() As Double()
    springArr = mSpring()
End Property

Public Property Get spring(i As Long) As Double
    spring = mSpring(i)
End Property

Public Property Let spring(i As Long, val As Double)
    mSpring(i) = val
End Property

Public Property Get x() As Double
    x = mx
End Property

Public Property Let x(locX As Double)
    mx = locX
End Property

Public Property Get y() As Double
    y = my
End Property

Public Property Let y(locY As Double)
    my = locY
End Property

Public Property Get z() As Double
    z = mz
End Property

Public Property Let z(locZ As Double)
    mz = locZ
End Property

Public Property Get constraint() As Collection 'collection of joint
    Set constraint = mConstraint
End Property

Public Property Let constraint(jts As Collection)
    Set mConstraint = jts
End Property

Public Property Get isRestraint() As Boolean
    isRestraint = mIsRestraint
End Property

Public Property Let isRestraint(val As Boolean)
    mIsRestraint = val
End Property

Public Property Get isConn() As Boolean
    isConn = mIsConn
End Property

Public Property Let isConn(val As Boolean)
    mIsConn = val
End Property
'Property Get supportType() As String
'    supportType = pSupportType
'End Property

'Property Get connType() As String
'    connType = pConnType
'End Property

Public Property Get localAxis() As Double
    localAxis = mLocalAxis
End Property

Public Property Let localAxis(deg As Double)
    mLocalAxis = deg
End Property


Public Property Get restraintArr_str() As String
    'format will be like 000000(free), 111000(pinned), 111111(fixed)
    Dim i As Integer, str As String
    Dim freeText As String
    Dim fixText As String
    freeText = "0"
    fixText = 1
    For i = 1 To 6
        If mRestraint(i) = True Then
            str = str & fixText
        Else
            str = str & freeText
        End If
    Next i
    restraintArr_str = str
End Property

Public Property Let restraintArr_str(str As String)
    Dim i As Integer
    Select Case GetDefSupportTypeVal(str)
    Case Roller
        For i = 1 To 6
            mRestraint(i) = False
        Next i
        mRestraint(3) = True
    Case Pinned
        For i = 1 To 3
            mRestraint(i) = True
        Next i
        For i = 4 To 6
            mRestraint(i) = False
        Next i
    Case Fixed
        For i = 1 To 6
            mRestraint(i) = True
        Next i
    Case Else
        If str = vbNullString Then
            For i = 1 To 6
                mRestraint(i) = False
            Next i
        Else
            For i = 1 To 6
                If Mid(str, i, 1) = "F" Then
                    mRestraint(i) = True
                ElseIf Mid(str, i, 1) = "0" Then
                    mRestraint(i) = False
                Else
                    'g_Log.RaiseWarning "Joint '" & name & "' Restraint Assignment Data in Wrong Format. Plesae Review.", jtRestraintInputWrong
                    Err.Raise 5010, "clsStrJoint.restraintArr_str", "Wrong Format of Joint Restraint Assignment Input"
                    Exit Property
                End If
            Next i
        End If
    End Select
End Property

Public Property Get restraintArr() As Boolean()
    restraintArr = mRestraint()
End Property

Public Property Get Restraint(i As Long) As Boolean
    Restraint = mRestraint(i)
End Property

Public Property Get restraint_str(i As Long) As String
    If mRestraint(i) Then
        restraint_str = "Yes"
    Else
        restraint_str = "No"
    End If
End Property

'Property Get distLoadPat() As String
'    distLoadPat = pDistLoadPat
'End Property
'
'Property Get distLoad(i As Long) As Double
'    distLoad = pDistLoad(i)
'End Property
'
'Property Get RecentAssignedLoad() As clsStrJointLoad
'    Set RecentAssignedLoad = pJtLoads.Item(pJtLoads.count)
'End Property

'Public Property Get load(index As Integer) As clsStrJointLoad
'
'End Property

'Public Property Get loads() As Object
'    Set loads = pJtLoads
'End Property
'
'Public Sub AddJtLoadObj(jtLoad As clsStrJointLoad)
'    pJtLoads.Add jtLoad
'End Sub
Public Sub Move(adj_x As Double, adj_y As Double, adj_z As Double)
    mx = mx - adj_x
    my = my - adj_y
    mz = mz - adj_z
End Sub

'Public Sub AssignLoad(loadPat As String, dir As String, val As Double, Optional loadType As String = "joint", Optional loadGrp As String = vbNullString)
'    Dim jtLoad As clsStrJointLoad
'    Set jtLoad = New clsStrJointLoad
'    If loadType = vbNullString Then loadType = "joint"
'    jtLoad.Init loadPat, Me, dir, val, loadType, loadGrp
'    pJtLoads.Add jtLoad
'End Sub

'Sub AssignConstraint()
'    'input shall be a collection of joint object
'End Sub

Property Get connectedFrames() As Object
    If mConnectedFrms.count > 0 Then
        Set connectedFrames = mConnectedFrms
    End If
End Property

Property Get connectedFrames_JtI() As Object
    'return the frames that the joint I of the frames = the joint
    Dim i As Long
    Dim frames_jtI As New Collection
    
    If Not mConnectedFrms.count = 0 Then
        For i = 1 To mConnectedFrms.count
            If mConnectedFrms(i).jtI.name = mName Then
                frames_jtI.Add mConnectedFrms(i), mConnectedFrms(i).name
            End If
        Next i
    End If
    Set connectedFrames_JtI = frames_jtI
End Property

Property Get connectedFrames_JtJ() As Object
    'return the frames that the joint J of the frames = the joint
    Dim i As Long
    Dim frames_jtJ As New Collection
    
    If Not mConnectedFrms.count = 0 Then
        For i = 1 To mConnectedFrms.count
            If mConnectedFrms(i).jtJ.name = mName Then
                frames_jtJ.Add mConnectedFrms(i), mConnectedFrms(i).name
            End If
        Next i
    End If
    Set connectedFrames_JtJ = frames_jtJ
End Property

Sub AddConnectedFrm(frm As StrFrame)
    mConnectedFrms.Add frm, frm.name
End Sub

Property Get connectedFramesStr(Optional prefix As String = "'") As String
    Dim frm As Variant, retStr As String
    If mConnectedFrms.count > 0 Then
        For Each frm In mConnectedFrms
            If Not retStr = vbNullString Then
                retStr = retStr & "," & frm.name
            Else
                retStr = prefix & frm.name
            End If
        Next
    End If
    connectedFramesStr = retStr
End Property

Property Get connectedFramesSectionStr(Optional prefix As String = "'") As String
    Dim frm As Variant, retStr As String
    If mConnectedFrms.count > 0 Then
        For Each frm In mConnectedFrms
            If Not retStr = vbNullString Then
                retStr = retStr & "," & frm.section
            Else
                retStr = prefix & frm.section
            End If
        Next
    End If
    connectedFramesSectionStr = retStr
End Property

Property Get SortedConnectedFramesSectionStr(Optional prefix As String = "'", _
                                            Optional isRemoveDuplicate As Boolean = False) As String
    Dim arrL As Object, frm As StrFrame
    Dim i As Long, retStr As String
    Set arrL = CreateObject("System.Collections.ArrayList")
    
    If mConnectedFrms.count = 0 Then Exit Property
    
    If isRemoveDuplicate Then
        For Each frm In mConnectedFrms
            If Not arrL.contains(frm.section) Then arrL.Add frm.section
        Next
    Else
        For Each frm In mConnectedFrms
            arrL.Add frm.section
        Next
    End If
    
    arrL.Sort
   
    For i = 0 To arrL.count - 1
        If Not retStr = vbNullString Then
            retStr = retStr & "," & arrL.item(i)
        Else
            retStr = prefix & arrL.item(i)
        End If
    Next i

    SortedConnectedFramesSectionStr = retStr
End Property

Property Get numOfConnectedFrames(orientationType As EleOrientationType) As Long
    Dim frm As StrFrame, count As Long
    If mConnectedFrms.count = 0 Then Exit Property
    For Each frm In mConnectedFrms
        If frm.orientationType = orientationType Then count = count + 1
    Next
    numOfConnectedFrames = count
End Property

Property Get connectedMembersStr(Optional prefix As String = "'") As String
    Dim frm As Variant, retStr As String
    If mConnectedFrms.count > 0 Then
        For Each frm In mConnectedFrms
            If Not retStr = vbNullString Then
                retStr = retStr & "," & frm.memberName
            Else
                retStr = prefix & frm.memberName
            End If
        Next
    End If
    connectedMembersStr = retStr
End Property

Public Function AddConnectionTypes(conn As StrConnectionType)
    mConnectionsTypes.Add conn, conn.name
End Function

Property Get connectionsTypes() As Object
    If mConnectionsTypes.count > 0 Then
        Set connectionsTypes = mConnectionsTypes
    End If
End Property

Property Get matchedConnectionStr() As String
    Dim conn As Variant, retStr As String
    If mConnectionsTypes.count > 0 Then
        For Each conn In mConnectionsTypes
            If Not retStr = vbNullString Then
                retStr = retStr & "," & conn.name
            Else
                retStr = "'" & conn.name
            End If
        Next
    End If
    If retStr = vbNullString Then retStr = "-"
    matchedConnectionStr = retStr
End Property


