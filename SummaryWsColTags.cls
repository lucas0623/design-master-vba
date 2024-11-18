VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SummaryWsColTags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Operation.SummaryToWS")
' The class module name should be "SummaryWsColTags"
Option Explicit

' Private member variables
Private pName As String
Private pColNumInWs As Long
Private pColNumInDF As Long
Private pIsOutputTag As Boolean
Private pIsTargetWsTag As Boolean

' Property for name
Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(value As String)
    pName = value
End Property

' Property for colNumInWs
Public Property Get colNumInWs() As Long
    colNumInWs = pColNumInWs
End Property

Public Property Let colNumInWs(value As Long)
    pColNumInWs = value
End Property

' Property for colNumInDF
Public Property Get colNumInDF() As Long
    colNumInDF = pColNumInDF
End Property

Public Property Let colNumInDF(value As Long)
    pColNumInDF = value
End Property

' Property for isOutputTag
Public Property Get isOutputTag() As Boolean
    isOutputTag = pIsOutputTag
End Property

Public Property Let isOutputTag(value As Boolean)
    pIsOutputTag = value
End Property

' Property for isTargetWsTag
Public Property Get isTargetWsTag() As Boolean
    isTargetWsTag = pIsTargetWsTag
End Property

Public Property Let isTargetWsTag(value As Boolean)
    pIsTargetWsTag = value
End Property

' Method to Initialize the class properties
Public Sub Initialize(ByVal Name As String, ByVal colNumInWs As Long)
    Me.Name = Name
    Me.colNumInWs = colNumInWs
End Sub
