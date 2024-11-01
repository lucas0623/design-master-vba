VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TagSelectorConnectionType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("DataSheet.ModelData")

Option Explicit
Public Property Get connID() As String
    connID = "connID"
End Property

Public Property Get mapJt() As String
    mapJt = "mapJt"
End Property

Public Property Get mapMethod() As String
    mapMethod = "mapMethod"
End Property

Public Property Get mapSections() As String
    mapSections = "mapSections"
End Property

Public Property Get isRestraint() As String
    isRestraint = "isRestraint"
End Property

Public Property Get extractSections() As String
    extractSections = "extractSections"
End Property
Public Property Get isFullyPopulate() As String
    isFullyPopulate = "isFullyPopulate"
End Property
Public Property Get endOffset() As String
    endOffset = "endOffset"
End Property
Public Property Get matchedJoint() As String
    matchedJoint = "matchedJoint"
End Property


