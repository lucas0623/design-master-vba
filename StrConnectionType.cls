VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StrConnectionType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Structural Model"
Option Explicit
Private mName As String
Private mJoints As New Collection 'matched joints

Private mMapJtsName As String
Private mMapMethod As String
Private mMapSections As String
Private mIsRestraint As Boolean

Private mExtractSections As String
Private mIsFullyPopulate As String
Private mEndOffset As Double

Public Sub Init(Name As String)
    mName = Name

End Sub

Public Sub InitForMapping(Name As String, mapJtName As String, mapMethod As String, mapSections As String, isRestraint As Boolean)
    mName = Name
    mMapJtsName = Replace(mapJtName, "'", "", 1, 1)
    mMapMethod = mapMethod
    mMapSections = Replace(mapSections, "'", "", 1, 1)
    mIsRestraint = isRestraint
End Sub

Public Sub InitForExtraction(Name As String, extractSections As String, isFullyPopulate As Variant, endOffset As Variant, matchedJoint As Collection)
    mName = Name
    mExtractSections = Replace(extractSections, "'", "", 1, 1)
    If VarType(isFullyPopulate) = vbBoolean Then
        mIsFullyPopulate = isFullyPopulate
    Else
        mIsFullyPopulate = True
    End If
    
    If VarType(endOffset) = vbDouble Then
        mEndOffset = endOffset
    Else
        mEndOffset = 0
    End If
    
    Set mJoints = matchedJoint
    
End Sub

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(str As String)
    mName = str
End Property

Public Property Get joints() As Collection 'collection of joint
    Set joints = mJoints
End Property

Public Property Let joints(jts As Collection)
    Set mJoints = jts
End Property

Public Property Get jointsName() As String
    Dim jt As Variant, retStr As String
    If mJoints.count > 0 Then
        For Each jt In mJoints
            If Not retStr = vbNullString Then
                retStr = retStr & "," & jt.Name
            Else
                retStr = "'" & jt.Name
            End If
        Next
    End If
    jointsName = retStr
End Property

Public Property Get jointsNameArr() As String()
    Dim retStr() As String, i As Long
    retStr = Split(Me.jointsName, ",")
    For i = LBound(retStr) To UBound(retStr)
        retStr(i) = Trim(retStr(i))
    Next i
    retStr(0) = Replace(retStr(0), "'", "", 1, 1)
    jointsNameArr = retStr
End Property

' mMapJtsName
Public Property Get mapJtsName() As String
    mapJtsName = mMapJtsName
End Property

Public Property Let mapJtsName(ByVal value As String)
    mMapJtsName = value
End Property

' mMapMethod
Public Property Get mapMethod() As String
    mapMethod = mMapMethod
End Property

Public Property Let mapMethod(ByVal value As String)
    mMapMethod = value
End Property

' mMapSections
Public Property Get mapSections() As String
    mapSections = mMapSections
End Property

Public Property Let mapSections(ByVal value As String)
    mMapSections = value
End Property

Property Get numOfFrames(orientationType As EleOrientationType) As Long
    Dim jt As StrJoint
    Set jt = mJoints(1)
    numOfFrames = jt.numOfConnectedFrames(orientationType)
End Property

Property Get numOfVertFrames() As Long
    numOfVertFrames = Me.numOfFrames(Vertical)
End Property
Property Get numOfDiagonalFrames() As Long
    numOfDiagonalFrames = Me.numOfFrames(Diagonal)
End Property
Property Get numOfHorizontalFrames() As Long
    numOfHorizontalFrames = Me.numOfFrames(Horizontal)
End Property
' mIsRestraint
Public Property Get isRestraint() As Boolean
    isRestraint = mIsRestraint
End Property

Public Property Let isRestraint(ByVal value As Boolean)
    mIsRestraint = value
End Property

' mExtractSections
Public Property Get extractSections() As String
    extractSections = mExtractSections
End Property

Public Property Let extractSections(ByVal value As String)
    mExtractSections = value
End Property

' mExtractMethod
Public Property Get extractMethod() As String
    extractMethod = mIsFullyPopulate
End Property

Public Property Let extractMethod(ByVal value As String)
    mIsFullyPopulate = value
End Property

' mEndOffset
Public Property Get endOffset() As Double
    endOffset = mEndOffset
End Property

Public Property Let endOffset(ByVal value As Double)
    mEndOffset = value
End Property

Public Function AddMatchedJoints(jt As StrJoint)
    mJoints.Add jt, jt.Name
End Function

Public Property Get isFullyPopulate() As Boolean
    isFullyPopulate = mIsFullyPopulate
End Property
Private Function SplitAndTrimString(str As String, andOperator As String) As String()
    Dim result() As String

    Dim i As Long

    ' Use the Split function to split the input string using the andOperator delimiter
    result = Split(str, andOperator)
    
    ' Trim and store the values in the result array
    For i = LBound(result) To UBound(result)
        result(i) = Trim(result(i))
    Next i

    ' Set the output as the result array
    SplitAndTrimString = result
End Function
