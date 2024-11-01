VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StrFrameForceForConn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Structural Model"
Private mConnType As StrConnectionType
Private mNode As StrJoint
Private mFrmForce As StrFrameForce

'These props are copy from the object. may assigned empty when 'mConnType.isFullyPopulate = false
Private mConnTypeName As String, mNodeName As String
Private mLc As String


Public Sub Initialize(connType As StrConnectionType, node As StrJoint, frmForce As StrFrameForce, _
            Optional outputConnTypeName As String = "-", Optional outputNodeName As String = "-", _
            Optional outputLoadCombName As String = "-")
    Set mConnType = connType
    Set mNode = node
    Set mFrmForce = frmForce
    
    If outputConnTypeName = "-" Then
        mConnTypeName = connType.name
    Else
        mConnTypeName = outputConnTypeName
    End If
    
    If outputNodeName = "-" Then
        mNodeName = node.name
    Else
        mNodeName = outputNodeName
    End If
    
    If outputLoadCombName = "-" Then
        mLc = frmForce.loadComb
    Else
        mLc = outputLoadCombName
    End If
End Sub

Public Property Get connType() As StrConnectionType
    Set connType = mConnType
End Property

Public Property Let connType(value As StrConnectionType)
    Set mConnType = value
End Property

' Public Get and Let properties for mNode
Public Property Get node() As StrJoint
    Set node = mNode
End Property

Public Property Let node(value As StrJoint)
    Set mNode = value
End Property

' Public Get and Let properties for mFrmForce
Public Property Get frmForce() As StrFrameForce
    Set frmForce = mFrmForce
End Property

Public Property Let frmForce(ByVal value As StrFrameForce)
    Set mFrmForce = value
End Property

Public Property Get nodeName() As String
    nodeName = mNodeName
End Property

Public Property Get connTypeName() As String
    connTypeName = mConnTypeName
End Property

Public Property Get frame() As StrFrame
    Set frame = mFrmForce.frame
End Property

Public Property Get frameName() As String
    frameName = mFrmForce.frameName
End Property

Public Property Get frameSection() As String
    frameSection = mFrmForce.frameSection
End Property

Public Property Get frameLength() As Double
    frameLength = mFrmForce.frameLength
End Property

Public Property Get frameJtIName() As String
    frameJtIName = mFrmForce.frameJtIName
End Property

Public Property Get frameJtJName() As String
    frameJtJName = mFrmForce.frameJtJName
End Property

Public Property Get memberName() As String
    memberName = mFrmForce.memberName
End Property

Public Property Get memberIFrameName() As String
    memberIFrameName = mFrmForce.memberIFrameName
End Property

Public Property Get memberJFrameName() As String
    memberJFrameName = mFrmForce.memberJFrameName
End Property

Public Property Get preFrameName() As String
    preFrameName = mFrmForce.preFrameName
End Property

Public Property Get nextFrameName() As String
    nextFrameName = mFrmForce.nextFrameName
End Property

Public Property Get memberTotalLength() As Double
    memberTotalLength = mFrmForce.memberTotalLength
End Property

Public Property Get pos_fromMemJtI_percent() As Double
    pos_fromMemJtI_percent = mFrmForce.pos_fromMemJtI_percent
End Property

Public Property Get pos_fromMemJtJ_percent() As Double
    pos_fromMemJtJ_percent = mFrmForce.pos_fromMemJtJ_percent
End Property

Public Property Get pos_fromMemJtI() As Double
    pos_fromMemJtI = mFrmForce.pos_fromMemJtI
End Property

Public Property Get pos_fromMemJtJ() As Double
    pos_fromMemJtJ = mFrmForce.pos_fromMemJtJ
End Property

Public Property Get pos_fromEleJtI() As Double
    pos_fromEleJtI = mFrmForce.pos_fromEleJtI
End Property

Public Property Get pos_fromEleJtJ() As Double
    pos_fromEleJtJ = mFrmForce.pos_fromEleJtJ
End Property

Public Property Get pos_fromEleJtI_percent() As Double
    pos_fromEleJtI_percent = mFrmForce.pos_fromEleJtI_percent
End Property

Public Property Get pos_fromEleJtJ_percent() As Double
    pos_fromEleJtJ_percent = mFrmForce.pos_fromEleJtJ_percent
End Property

Public Property Get station() As Double
    station = mFrmForce.station
End Property

Public Property Get loadComb() As String
    loadComb = mLc
End Property

Public Property Get stepType() As String
    stepType = mFrmForce.stepType
End Property


Public Property Get forces() As Double()
    forces = mFrmForce.forces
End Property

Public Property Get force(index As Integer) As Double
    force = mFrmForce.force(index)
End Property

Public Property Get p() As Double
    p = mFrmForce.p
End Property

Public Property Get V2() As Double
    V2 = mFrmForce.V2
End Property

Public Property Get V3() As Double
    V3 = mFrmForce.V3
End Property

Public Property Get t() As Double
    t = mFrmForce.t
End Property

Public Property Get M2() As Double
    M2 = mFrmForce.M2
End Property

Public Property Get M3() As Double
    M3 = mFrmForce.M3
End Property

'Public Property Get description() As String
'    description = pDesc
'End Property
'
'Public Property Let description(text As String)
'    pDesc = text
'End Property

Public Property Get extremeCaseType() As String
    extremeCaseType = mFrmForce.extremeCaseType
End Property

Public Property Get maxMinEnvPerm() As String
    maxMinEnvPerm = mFrmForce.maxMinEnvPerm
End Property

