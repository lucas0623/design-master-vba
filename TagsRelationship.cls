VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TagsRelationship"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Operation.SummaryToWS")
' Class Module: TagsRelationship

' A class to contain the data of design Ws IO Tag <> summary tag conversion.
' Used in DesignWsGroup
Private mName As String
Private mIsPrint As Boolean
Private mTagRelation As Object

' Class Initialization
Private Sub Class_Initialize()
    mIsPrint = True
    Set mTagRelation = CreateObject("Scripting.Dictionary")
End Sub

' Method to add a tag relationship
Public Function Add(designTag As String, summaryTag As String) As Integer
    Dim ret As Integer
    If Not mTagRelation.Exists(designTag) Then
        mTagRelation.Add designTag, summaryTag
    Else
        ret = -1
    End If
ExitFunc:
    Add = ret
End Function

' Method to get a summary tag from a design tag
Public Function GetSummaryTag(designTag As String) As String
    If mTagRelation.Exists(designTag) Then
        GetSummaryTag = mTagRelation(designTag)
    Else
        GetSummaryTag = "" ' Return empty if not found
    End If
End Function

Public Property Get tagRelation() As Object
    Set tagRelation = mTagRelation
End Property
' Method to remove a tag relationship
Public Sub Remove(designTag As String)
    If mTagRelation.Exists(designTag) Then
        mTagRelation.Remove designTag
    End If
End Sub

' Method to print all tag relations (for debugging or display)
Public Sub PrintTagRelations()
    Dim key As Variant
    For Each key In mTagRelation.keys
        Debug.Print "Design Tag: " & key & " => Summary Tag: " & mTagRelation(key)
    Next key
End Sub

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(value As String)
    mName = value
End Property

Public Property Get isPrint() As Boolean
    isPrint = mIsPrint
End Property

Public Property Let isPrint(value As Boolean)
    mIsPrint = value
End Property

Public Sub Initialize(ByVal Name As String)
    Me.Name = Name
End Sub
