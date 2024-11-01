VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "lib3DCoorGeom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Library")
Private Const PI As Double = 3.14159265358979

Public Function AngleToHorizontalPlane(x1 As Double, y1 As Double, z1 As Double, x2 As Double, y2 As Double, z2 As Double, _
                                Optional isRadian As Boolean = True) As Double
    ' Calculate the vector components of the line
    Dim dx As Double
    Dim dy As Double
    Dim dz As Double
    Dim dxy As Double
    dx = x2 - x1
    dy = y2 - y1
    dz = z2 - z1
    dxy = (dx ^ 2 + dy ^ 2) ^ (1 / 2)
    
    ' Calculate the angle between the line and the horizontal plane
    Dim angle As Double
    If dxy = 0 Then
        angle = PI / 2
    Else
        angle = Math.Atn(dz / (dx ^ 2 + dy ^ 2) ^ (1 / 2))
    End If
    
    If Not isRadian Then angle = Degree(angle)

    AngleToHorizontalPlane = angle
End Function

Private Function Radian(deg As Double) As Double
    Radian = deg * PI / 180
End Function

Private Function Degree(rad As Double) As Double
    Degree = rad * 180 / PI
End Function
