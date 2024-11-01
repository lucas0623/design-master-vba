VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TagSelectorJointCoor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("DataSheet.ModelData")

'********************************************************
'Arthor: Lucas LEUNG
'Update Log
'09 Sep 2023 - Initial Setup
'*******************************************************

Option Explicit

Property Get ID() As String
    ID = "ID"
End Property

Property Get x() As String
    x = "x"
End Property

Property Get y() As String
    y = "y"
End Property
Property Get z() As String
    z = "z"
End Property

Property Get isRestraint() As String
    isRestraint = "isRestraint"
End Property
