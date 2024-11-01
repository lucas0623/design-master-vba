VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MemberRestrTypeClassifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Structural Model")
Option Explicit

Private FrameTypeClassifier As New FrameTypeClassifier
Private Geom2D As New lib2DCoorGeom
Private DefAngleBetweenVertMajAndX As Double 'Model Default Angle Between The major axis of a vertical element to x-axis
Private DefAngleBetweenHorMajAndZ As Double 'Model Default Angle Between The major axis of a horizontal element to vertical plane
Private LimitAngleBetweenMajAxis As Double

Public Function Classify(member As StrMember, jt As StrJoint) As RestraintType
    Dim connectedFrms As New Collection
    Dim frm As StrFrame
    
    Set connectedFrms = member.connectedFrames(jt)
    
    
    
    
End Function

'Private Function ChooseMethod(member As StrMember, jt As StrJoint) As RestraintType
'    Dim cFrameType As FrameType, ret As RestraintType
'    'Get the frame orientation type of the member
'    For Each frm In connectedFrms
'        If frm.memberName = member.Name Then
'            cFrameType = frm.OrientationType
'            Exit For
'        End If
'    Next
'    Set connectedFrms = member.connectedFrames(jt)
'
'    Select Case cFrameType
'        Case Vertical: ret = VerticalMainMemberAllFrames
'        Case Diagonal: ret = HorizontalMainMemberAllFrames
'        Case Horizontal: ret = DiagonalMainMemberAllFrames
'    End Select
'
'End Function

Private Function VerticalMainMemberAllFrames(member As StrMember, connectedFrms As Collection) As RestraintType

End Function

Private Function VerticalMainMemberOneFrame(member As StrMember, frmInMember As StrFrame, frm As StrFrame) As RestraintType
    Dim ret As RestraintType, angleToX As Double, angleBetweenMemMajAndFrm As Double
    Dim localAxis As Double
    If frm.orientationType = Horizontal Or frm.orientationType = Diagonal Then
        angleToX = Geom2D.AngleToHorizontal(frm.jtI.x, frm.jtI.y, frm.jtJ.x, frm.jtJ.y)
        localAxis = Math.Abs(frmInMember.localAxis) Mod 90
        angleBetweenMemMajAndFrm = Math.Abs(angleToX - localAxis - DefAngleBetweenVertMajAndX)
        
        If angleBetweenMemMajAndFrm < LimitAngleBetweenMajAxis Then
            ret = MajorOnly
        ElseIf (90 - angleBetweenMemMajAndFrm) > LimitAngleBetweenMajAxis Then
            ret = MinorOnly
        Else
            ret = Undefined
        End If
    Else
        ret = Undefined
    End If
    
End Function

Private Function HorizontalMainMemberOneFrame(member As StrMember, frmInMember As StrFrame, frm As StrFrame) As RestraintType


End Function

Private Function DiagonalMainMemberOneFrame(member As StrMember, frmInMember As StrFrame, frm As StrFrame) As RestraintType
    DiagonalMainMemberOneFrame = HorizontalMainMemberOneFrame(member, frmInMember, frm)
End Function

Private Function SumRestraintType(restr1 As RestraintType, restr2 As RestraintType) As RestraintType
'if MajorOnly + MI
    If (restr1 = MajorOnly And restr2 = MinorOnly) Or (restr2 = MajorOnly And restr1 = MinorOnly) Then
        SumRestraintType = MajorAndMinor
    ElseIf restr1 = Undefined Then
        SumRestraintType = restr2
    Else
        SumRestraintType = restr1
    End If
End Function
