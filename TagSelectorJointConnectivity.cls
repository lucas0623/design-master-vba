VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TagSelectorJointConnectivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("DataSheet.ModelData")

Option Explicit
Public Property Get jtID() As String
    jtID = "jtID"
End Property

Public Property Get memID() As String
    memID = "memID"
End Property

Public Property Get eleID() As String
    eleID = "eleID"
End Property

Public Property Get sections() As String
    sections = "sections"
End Property

Public Property Get midMem() As String
    midMem = "midMem"
End Property

Public Property Get endMem() As String
    endMem = "endMem"
End Property
Public Property Get isConn() As String
    isConn = "isConn"
End Property
Public Property Get isRestraint() As String
    isRestraint = "isRestraint"
End Property
Public Property Get matchedType() As String
    matchedType = "matchedType"
End Property

