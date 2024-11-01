VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "oProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("DataSheet")
Option Explicit
Private mText As String
Private mLoc As Long
Private mDesc As String


' Create Get and Set properties for mText
Public Property Get Text() As String
    Text = mText
End Property

Public Property Let Text(ByVal value As String)
    mText = value
End Property

' Create Get and Set properties for mLoc
Public Property Get Loc() As Long
    Loc = mLoc
End Property

Public Property Let Loc(ByVal value As Long)
    mLoc = value
End Property

' Create Get and Set properties for mDesc
Public Property Get Desc() As String
    Desc = mDesc
End Property

Public Property Let Desc(ByVal value As String)
    mDesc = value
End Property
