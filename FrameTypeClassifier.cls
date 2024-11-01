VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FrameTypeClassifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Structural Model")

Private DsSys As New DataSheetSystem
Private Geom3D As New lib3DCoorGeom

Private degVert As Double, degHori As Double


Private Sub Class_Initialize()
    degVert = DsSys.prop("ElementTypeIdentification", "DegToHoriPlane_Vert")
    degHori = DsSys.prop("ElementTypeIdentification", "DegToHoriPlane_Hori")
End Sub
Public Function EnumToStr(FrameType As EleOrientationType) As String
    Dim ret As String
    Select Case FrameType
        Case Vertical: ret = "Vertical"
        Case Horizontal: ret = "Horizontal"
        Case Diagonal: ret = "Diagonal"
    End Select
    EnumToStr = ret
End Function

Public Function StrToEnum(str As String) As EleOrientationType
    Dim ret As EleOrientationType
    Select Case str
        Case "Vertical": ret = Vertical
        Case "Horizontal": ret = Horizontal
        Case "Diagonal": ret = Diagonal
    End Select
    StrToEnum = ret
End Function

Public Function Classify(frame As StrFrame) As EleOrientationType
    Dim angle As Double, frmType As EleOrientationType
    
    With frame
        angle = Geom3D.AngleToHorizontalPlane(.jtI.x, .jtI.y, .jtI.z, .jtJ.x, .jtJ.y, .jtJ.z, False)
        If angle < degHori Then
            frmType = Horizontal
        ElseIf angle > degVert Then
            frmType = Vertical
        Else
            frmType = Diagonal
        End If
    End With
    
    Classify = frmType
    
End Function

