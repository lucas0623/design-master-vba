VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "lib2DCoorGeom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Library")
Private Const PI As Double = 3.14159265358979

Public Function Dist(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double
    Dist = ((x2 - x1) ^ 2 + (y2 - y1) ^ 2) ^ (1 / 2)
End Function

Public Function Intersection(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double()

End Function

Public Function isParallel(x1 As Double, y1 As Double, x2 As Double, y2 As Double, _
                        x3 As Double, y3 As Double, x4 As Double, y4 As Double, _
                        Optional tolerance As Double = 0.0001) As Boolean
    Dim m12 As Double, m34 As Double
    m12 = Slope(x1, y1, x2, y2)
    m34 = Slope(x3, y3, x3, y3)
    If m12 = m34 Then
        isParallel = True
    ElseIf (m12 + tolerance) > m34 And (m12 - tolerance) < m34 Then
        isParallel = True
    End If
End Function

Public Function AngleToHorizontal(x1 As Double, y1 As Double, x2 As Double, y2 As Double, _
                                Optional isRadian As Boolean = True) As Double
    'return 0 to 90 degree
    Dim dx As Double, dy As Double
    dx = x2 - x1
    dy = y2 - y1
    
    Dim angle As Double
    If dx = 0 Then
        angle = PI / 2
    Else
        angle = Math.Abs(Math.Atn(dy / dx))
    End If
    
    If Not isRadian Then angle = Degree(angle)

    AngleToHorizontal = angle
End Function

Public Function AngleToVertical(x1 As Double, y1 As Double, x2 As Double, y2 As Double, _
                                Optional isRadian As Boolean = True) As Double
    'return 0 to 90 degree
    AngleToVertical = (PI / 2) - Math.Abs(AngleToHorizontal(x1, y1, x2, y2, isRadian))
End Function

Public Function AngleBetweenTwoLines(x1 As Double, y1 As Double, x2 As Double, y2 As Double, _
                        x3 As Double, y3 As Double, x4 As Double, y4 As Double, Optional isRadian As Boolean = True) As Double
    
    Dim m12 As Double, m34 As Double, angle As Double
    m12 = Slope(x1, y1, x2, y2)
    m34 = Slope(x3, y3, x3, y3)
    If m12 * m34 = -1 Then
        angle = PI / 2
    Else
        angle = Math.Atn((m12 - m34) / (1 + m12 * m34))
    End If
    
    If Not isRadian Then angle = Degree(angle)

    AngleBetweenTwoLines = angle
End Function

Public Function Slope(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double
    Dim dx As Double, dy As Double
    dx = x2 - x1
    dy = y2 - y1
    
    Dim ret As Double
    If dx = 0 Then
        ret = 1E+16
    Else
        ret = dy / dx
    End If

    Slope = ret
End Function

Private Function Radian(deg As Double) As Double
    Radian = deg * PI / 180
End Function

Private Function Degree(rad As Double) As Double
    Degree = rad * 180 / PI
End Function


