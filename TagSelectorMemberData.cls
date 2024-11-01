VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TagSelectorMemberData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("DataSheet.ModelData")
Option Explicit
Public Property Get memID() As String
    memID = "MemberID"
End Property

Public Property Get fFrame() As String
    fFrame = "fFrame"
End Property

Public Property Get lFrame() As String
    lFrame = "lFrame"
End Property

Public Property Get totalLen() As String
    totalLen = "totalLen"
End Property

Public Property Get eleList() As String
    eleList = "eleList"
End Property
