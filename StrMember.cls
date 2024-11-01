VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StrMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Structural Model"


'********************************************************
'This module represent the Structural Joint/Node Object
'Arthor: Lucas LEUNG
'Update Log

'07 Jun 2023 - Initial
'*******************************************************
Option Explicit
Private pName As String
Private pNickName As String
Private pFrames As New Collection
Private pTotalLen As Double
Private mRestr As New Collection 'collection of Restraint object

Private Type Restraint
    jt As StrJoint
    restrType As String
End Type

Public Enum RestraintType
    MajorOnly
    MinorOnly
    MajorAndMinor
    Undefined
End Enum

Public Sub AddFrame(frame As StrFrame, isBefore As Boolean)
    If isKeyExist(pFrames, frame.name) Then Exit Sub
    
    If isBefore Then
        If pFrames.count = 0 Then
            pFrames.Add frame, frame.name
        Else
            pFrames.Add frame, frame.name, Before:=1
        End If
    Else
        pFrames.Add frame, frame.name
    End If
    
    pTotalLen = pTotalLen + frame.Length
    If isBefore Then pName = frame.name
End Sub


Private Function isKeyExist(coll As Collection, sKey As String) As Boolean
    Dim obj As Object
    On Error GoTo NotExist:
    Set obj = coll(sKey)
    isKeyExist = True
    Exit Function
NotExist:
    isKeyExist = False
    On Error GoTo -1
    'Debug.Print "Err Num: =" & Err.Number
End Function

Public Property Get RelativeLength(frmID As String, isFromI As Boolean) As Double
    Dim rLen As Double
    Dim i As Long
    
    For i = 1 To pFrames.count
        If pFrames.item(i).name = frmID Then
            If isFromI Then
                RelativeLength = rLen
            Else
                RelativeLength = pTotalLen - rLen - pFrames.item(i).Length
            End If
            Exit Property
        End If
        rLen = rLen + pFrames.item(i).Length
    Next i
    
End Property

Public Property Get TotalLength() As Double
    TotalLength = pTotalLen
End Property

Public Property Get frames() As Object
    Set frames = pFrames
End Property
Public Property Get preFrame(frameID As String) As StrFrame
    Dim i As Long
    For i = 1 To pFrames.count
        If pFrames(i).name = frameID Then
            If Not i = 1 Then
                Set preFrame = pFrames(i - 1)
                Exit Property
            Else
                Set preFrame = Nothing
                Exit Property
            End If
        End If
    Next i
End Property
Public Property Get nextFrame(frameID As String) As StrFrame
    Dim i As Long
    For i = 1 To pFrames.count
        If pFrames.item(i).name = frameID Then
            If Not i = pFrames.count Then
                Set nextFrame = pFrames.item(i + 1)
                Exit Property
            Else
               'Set nextFrame = Nothing
                Exit Property
            End If
        End If
    Next i
End Property
Public Property Get iFrame() As StrFrame
    Set iFrame = pFrames(1)
End Property
Public Property Get iFrameName() As String
    iFrameName = pFrames(1).name
End Property
Public Property Get jFrame() As StrFrame
    Set jFrame = pFrames(pFrames.count)
End Property
Public Property Get jFrameName() As String
    jFrameName = pFrames(pFrames.count).name
End Property
Public Property Get name() As String
    If pName = vbNullString Then
        pName = Me.iFrameName
    End If
    name = pName
End Property

Public Property Get eleList(Optional delimiter As String = ",", Optional prefix As String = "'") As String
    Dim i As Long, str As String
    str = prefix & pFrames(1).name
    If Not pFrames.count = 1 Then
        For i = 2 To pFrames.count
            str = str & delimiter & pFrames.item(i).name
        Next i
    End If
    eleList = str
End Property

Public Property Get eleNames() As String()
    Dim str() As String
    str = Split(Me.eleList, ",")
    str(0) = Replace(str(0), "'", "", 1, 1)
    str = eleNames
End Property

Public Sub AddRestriant(jt As StrJoint, Optional restrType As RestraintType = Undefined)
    Dim restr As Restraint
    Set restr.jt = jt
    If Not restrType = Undefined Then restr.restrType = restrType
    mRestr.Add restrType
End Sub

Public Property Get connectedFrames(jt As StrJoint) As Collection
    Dim frms1 As Collection, frms2 As Collection
    Set frms1 = jt.connectedFrames
    If Not frms1.count = 0 Then
        Set frms2 = CloneCollection(frms1)
        Dim i As Long
        For i = 1 To frms2.count
            If frms2(i).memberName = Me.name Then frms2.Remove (i)
        Next i
    End If
    Set connectedFrames = frms2
End Property

Public Function CloneCollection(ByVal Source As Collection) As Collection
    Dim result As New Collection
    Dim item As Variant
    For Each item In Source
        result.Add item
    Next
    Set CloneCollection = result
End Function
'Property Get Force_byLen()
'
'End Property
'
'Property Get Force_by()
'
'End Property
'Private Function DictToArrObj(dict As Object) As Variant
'    Dim arr() As Object
'    Dim key As Variant, count As Long
'    ReDim arr(dict.count - 1)
'    For Each key In dict.keys
'        Set arr(count) = dict(key)
'        count = count + 1
'    Next key
'End Function
