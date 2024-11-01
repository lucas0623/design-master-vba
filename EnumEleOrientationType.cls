VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EnumEleOrientationType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Structural Model")
Public Enum EleOrientationType
    Vertical
    Horizontal
    Diagonal
End Enum

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
