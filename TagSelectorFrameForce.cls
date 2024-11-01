VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TagSelectorFrameForce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("DataSheet.ModelData")
                        
Public Property Get eleID() As String
    eleID = "eleID"
End Property

Public Property Get station() As String
    station = "station"
End Property

Public Property Get loadComb() As String
    loadComb = "loadComb"
End Property

Public Property Get stepType() As String
    stepType = "stepType"
End Property

Public Property Get p() As String
    p = "P"
End Property

Public Property Get V2() As String
    V2 = "V2"
End Property

Public Property Get V3() As String
    V3 = "V3"
End Property

Public Property Get t() As String
    t = "T"
End Property

Public Property Get M2() As String
    M2 = "M2"
End Property

Public Property Get M3() As String
    M3 = "M3"
End Property

Public Property Get section() As String
    section = "section"
End Property

Public Property Get memID() As String
    memID = "memID"
End Property

Public Property Get subEleName() As String
    subEleName = "subEleName"
End Property
