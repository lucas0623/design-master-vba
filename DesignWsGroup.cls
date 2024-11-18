VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DesignWsGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Operation.SummaryToWS")

' The class module name "DesignWsIO"
Option Explicit

' Private member variables
Private mName As String
Private mTagsRelationships As Collection ' A collection to contain TagsRelationship of each worksheet

' Class Initialization
Private Sub Class_Initialize()
    Set mTagsRelationships = New Collection
End Sub

' Property to get or set the name
Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(value As String)
    mName = value
End Property

Public Sub Initialize(ByVal Name As String)
    Me.Name = Name
End Sub

Public Property Get TagsRelationships() As Collection
    Set TagsRelationships = mTagsRelationships
End Property
' Method to add a TagsRelationship to the collection
Public Sub AddTagsRelationship(tagsRel As TagsRelationship)
    mTagsRelationships.Add tagsRel
End Sub

' Method to get a TagsRelationship by worksheet name
Public Function GetTagsRelationship(index As Long) As TagsRelationship
    Set GetTagsRelationship = mTagsRelationships.item(index)
End Function

' Method to remove a TagsRelationship by worksheet name
'Public Sub RemoveTagsRelationship(wsName As String)
'    On Error Resume Next
'    mTagsRelationships.Remove wsName
'    On Error GoTo 0
'End Sub

'' Method to print all worksheet relationships
'Public Sub PrintAllRelationships()
'    Dim wsName As Variant
'    For Each wsName In mTagsRelationships
'        Debug.Print "Worksheet: " & wsName
'        mTagsRelationships(wsName).PrintTagRelations
'    Next wsName
'End Sub

'' Method to list all worksheet names with relationships
'Public Function ListWorksheetNames() As Collection
'    Dim wsNames As Collection
'    Dim wsName As Variant
'
'    Set wsNames = New Collection
'    For Each wsName In mTagsRelationships
'        wsNames.Add wsName
'    Next wsName
'
'    Set ListWorksheetNames = wsNames
'End Function

