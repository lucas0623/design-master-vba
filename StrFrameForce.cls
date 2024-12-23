VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StrFrameForce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Structural Model"

'********************************************************
'This module represent the Structural Joint/Node Object
'Arthor: Lucas LEUNG
'Update Log

'09 Jun 2023 - Initial
'*******************************************************

Option Explicit
'this is the variable for the sub frame
Private pFrm As StrFrame, mSubFrameName As String
Private pStation As Double
Private pLoadComb As String
Private pStepType As String
Private pForces(5) As Double
Private pDesc_ExtremeCaseType As String
Private pDesc_MaxMinEnvPerm As String

Public Sub Init(frm As StrFrame, station As Double, loadComb As String, _
                stepType As String, forces() As Double, Optional isSaveToFrmColl As Boolean = True, _
                Optional subFrameName As String)
    Set pFrm = frm
    pStation = station
    pLoadComb = loadComb
    pStepType = stepType
    
    Dim i As Integer
    For i = 0 To 5
        pForces(i) = forces(i)
    Next i
    If subFrameName = vbNullString Then
        mSubFrameName = frm.Name
    Else
        mSubFrameName = subFrameName
    End If
End Sub

Public Sub Init2(frm As StrFrame, station As Double, loadComb As String, _
                stepType As String, p As Double, V2 As Double, V3 As Double, _
                t As Double, M2 As Double, M3 As Double, Optional subFrameName As String)
    Set pFrm = frm
    pStation = station
    pLoadComb = loadComb
    pStepType = stepType
    pForces(0) = p
    pForces(1) = V2
    pForces(2) = V3
    pForces(3) = t
    pForces(4) = M2
    pForces(5) = M3
    If subFrameName = vbNullString Then
        mSubFrameName = frm.Name
    Else
        mSubFrameName = subFrameName
    End If
End Sub

Public Property Get frame() As StrFrame
    Set frame = pFrm
End Property

Public Property Get frameName() As String
    frameName = pFrm.Name
End Property

Public Property Get subFrameName() As String
    subFrameName = mSubFrameName
End Property

Public Property Get frameSection() As String
    frameSection = pFrm.section
End Property

Public Property Get frameLength() As Double
    frameLength = pFrm.Length
End Property

Public Property Get frameJtIName() As String
    frameJtIName = pFrm.jtI.Name
End Property

Public Property Get frameJtJName() As String
    frameJtJName = pFrm.jtJ.Name
End Property

Public Property Get memberName() As String
    memberName = pFrm.member.Name
End Property

Public Property Get memberIFrameName() As String
    memberIFrameName = pFrm.member.iFrameName
End Property

Public Property Get memberJFrameName() As String
    memberJFrameName = pFrm.member.jFrameName
End Property

Public Property Get sectionName() As String
    sectionName = pFrm.section
End Property


Public Property Get preFrameName() As String
    Dim frm As StrFrame
    Set frm = pFrm.member.preFrame(pFrm.Name)
    If frm Is Nothing Then
        preFrameName = "N/A"
    Else
        preFrameName = frm.Name
    End If
End Property

Public Property Get nextFrameName() As String
    Dim frm As StrFrame
    Set frm = pFrm.member.nextFrame(pFrm.Name)
    If frm Is Nothing Then
        nextFrameName = "N/A"
    Else
        nextFrameName = frm.Name
    End If
End Property

Public Property Get memberTotalLength() As Double
    memberTotalLength = pFrm.member.TotalLength
End Property

Public Property Get pos_fromMemJtI_percent() As Double
    pos_fromMemJtI_percent = (pFrm.member.RelativeLength(pFrm.Name, True) + pStation) / memberTotalLength
End Property

Public Property Get pos_fromMemJtJ_percent() As Double
    pos_fromMemJtJ_percent = 1 - Me.pos_fromMemJtI_percent
End Property

Public Property Get pos_fromMemJtI() As Double
    pos_fromMemJtI = pFrm.member.RelativeLength(pFrm.Name, True) + pStation
End Property

Public Property Get pos_fromMemJtJ() As Double
    pos_fromMemJtJ = Me.memberTotalLength - Me.pos_fromMemJtI
End Property

Public Property Get pos_fromEleJtI() As Double
    pos_fromEleJtI = pStation
End Property

Public Property Get pos_fromEleJtJ() As Double
    pos_fromEleJtJ = pFrm.Length - pStation
End Property

Public Property Get pos_fromEleJtI_percent() As Double
    pos_fromEleJtI_percent = (Me.pos_fromEleJtI / pFrm.Length)
End Property

Public Property Get pos_fromEleJtJ_percent() As Double
    pos_fromEleJtJ_percent = (Me.pos_fromEleJtJ / pFrm.Length)
End Property

Public Property Get station() As Double
    station = pStation
End Property

Public Property Get loadComb() As String
    loadComb = pLoadComb
End Property

Public Property Get stepType() As String
    stepType = pStepType
End Property


Public Property Get forces() As Double()
    forces = pForces
End Property

Public Property Get force(index As Integer) As Double
    force = pForces(index)
End Property

Public Property Get p() As Double
    p = pForces(0)
End Property

Public Property Get V2() As Double
    V2 = pForces(1)
End Property

Public Property Get V3() As Double
    V3 = pForces(2)
End Property

Public Property Get t() As Double
    t = pForces(3)
End Property

Public Property Get M2() As Double
    M2 = pForces(4)
End Property

Public Property Get M3() As Double
    M3 = pForces(5)
End Property

'Public Property Get description() As String
'    description = pDesc
'End Property
'
'Public Property Let description(text As String)
'    pDesc = text
'End Property

Public Property Get extremeCaseType() As String
    extremeCaseType = pDesc_ExtremeCaseType
End Property

Public Property Let extremeCaseType(Text As String)
    pDesc_ExtremeCaseType = Text
End Property

Public Property Get maxMinEnvPerm() As String
    maxMinEnvPerm = pDesc_MaxMinEnvPerm
End Property

Public Property Let maxMinEnvPerm(Text As String)
    pDesc_MaxMinEnvPerm = Text
End Property

Public Property Get isMemberEndI() As Boolean
    If Me.pos_fromMemJtI_percent = 0 Then
        isMemberEndI = True
    Else
        isMemberEndI = False
    End If
End Property

Public Property Get isMemberEndJ() As Boolean
    If Me.pos_fromMemJtI_percent = 1 Then
        isMemberEndJ = True
    Else
        isMemberEndJ = False
    End If
End Property

Public Property Get isMemberEnd() As Boolean
    If isMemberEndI Or isMemberEndJ Then
        isMemberEnd = True
    Else
        isMemberEnd = False
    End If
End Property
