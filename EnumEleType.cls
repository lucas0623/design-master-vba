VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EnumEleType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Structural Model")
Public Enum EleType
    Beam
    link
End Enum

Public Function EnumToStr(FrameType As EleType) As String
    Dim ret As String
    Select Case FrameType
        Case Beam: ret = "Beam"
        Case link: ret = "Link"
    End Select
    EnumToStr = ret
End Function

Public Function StrToEnum(str As String) As EleType
    Dim ret As EleType
    Select Case str
        Case "Beam": ret = Beam
        Case "Link": ret = link
    End Select
    StrToEnum = ret
End Function

