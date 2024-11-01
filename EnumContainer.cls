VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EnumContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum PositionFilterType
    NoFilter = 0
    EndI = 1
    EndJ = 2
    BothEnd = 3
    MidOnly = 4
    MidAndBothEnd = 5
    SpecifiedPosAbs = 6
    SpecifiedPosRel = 7
End Enum


