VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataSheetManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("DataSheet.ModelData")

'********************************************************
'This module control
'Arthor: Lucas LEUNG
'Update Log
'05 Jun 2023 - Initial Setup
'*******************************************************
Option Explicit
'Private wsInteract As clsWorksheetsInteraction
Private mDataSheets As New Collection
'Private wsName As String, mDataSheet As oDataSheet

Private Enum ModelDataSheetType
    JointCoor
    FrameData
    FrameForce
    MemberData
    JointConnectivity
    ConnectionType
End Enum

'Private Sub Class_Initialize()
'
'    Set wsInteract = New clsWorksheetsInteraction
'
'End Sub

Private Sub InitDataSheetJoint()
    If isKeyExist(mDataSheets, DSName(JointCoor)) Then Exit Sub
    Dim dataSheet As oDataSheet
    Set dataSheet = New oDataSheet
    Dim tagSelector As TagSelectorJointCoor
    Set tagSelector = New TagSelectorJointCoor
    dataSheet.Initialize "Data_Joint", tagSelector, "ID", "x", "y", "z", "isRestraint"
    mDataSheets.Add dataSheet, DSName(JointCoor)
End Sub

Private Sub InitDataSheetFrameData()
    If isKeyExist(mDataSheets, DSName(FrameData)) Then Exit Sub
    Dim dataSheet As oDataSheet
    Set dataSheet = New oDataSheet
    Dim tagSelector As TagSelectorFrameData
    Set tagSelector = New TagSelectorFrameData
    dataSheet.Initialize "Data_Frame", tagSelector, "eleID", "section", "jtI", "jtJ", "eleLen", "memID", "eleType", "eleLocalAxis", "eleOrientationType"
    mDataSheets.Add dataSheet, DSName(FrameData)
End Sub

Private Sub InitDataSheetFrameForce()
    If isKeyExist(mDataSheets, DSName(FrameForce)) Then Exit Sub
    Dim dataSheet As oDataSheet
    Set dataSheet = New oDataSheet
    Dim tagSelector As TagSelectorFrameForce
    Set tagSelector = New TagSelectorFrameForce
    dataSheet.Initialize "Data_FrameForce", tagSelector, "eleID", "station", "loadComb", "stepType", _
                        "P", "V2", "V3", "T", "M2", "M3", "section", "memID", "subEleName"
    mDataSheets.Add dataSheet, DSName(FrameForce)
End Sub

Private Sub InitDataSheetMemberData()
    If isKeyExist(mDataSheets, DSName(MemberData)) Then Exit Sub
    Dim dataSheet As oDataSheet
    Set dataSheet = New oDataSheet
    Dim tagSelector As TagSelectorMemberData
    Set tagSelector = New TagSelectorMemberData
    dataSheet.Initialize "Data_MemberList", tagSelector, "MemberID", "fFrame", "lFrame", "totalLen", "eleList"
    mDataSheets.Add dataSheet, DSName(MemberData)
End Sub

Private Sub InitDataSheetJointConnectivity()
    If isKeyExist(mDataSheets, DSName(JointConnectivity)) Then Exit Sub
    Dim dataSheet As oDataSheet
    Set dataSheet = New oDataSheet
    Dim tagSelector As TagSelectorJointConnectivity
    Set tagSelector = New TagSelectorJointConnectivity
    dataSheet.Initialize "Data_JointConn", tagSelector, "jtID", "memID", "eleID", "sections", "isRestraint", "isConn", "matchedType"
    mDataSheets.Add dataSheet, DSName(JointConnectivity)
End Sub

Private Sub InitDataSheetConnectionType()
    If isKeyExist(mDataSheets, DSName(ConnectionType)) Then Exit Sub
    Dim dataSheet As oDataSheet
    Set dataSheet = New oDataSheet
    Dim tagSelector As TagSelectorConnectionType
    Set tagSelector = New TagSelectorConnectionType
    dataSheet.Initialize "Define_ConnectionType", tagSelector, "connID", "mapJt", "mapMethod", "mapSections", "isRestraint", "extractSections", "isFullyPopulate", "endOffset", "matchedJoint"
    mDataSheets.Add dataSheet, DSName(ConnectionType)
End Sub

Private Sub InitAll()
    InitDataSheetJoint
    InitDataSheetFrameData
    InitDataSheetFrameForce
    InitDataSheetMemberData
    InitDataSheetJointConnectivity
End Sub

Public Sub ClearAllData(Optional isClearSavedUserInput As Boolean = True)
    'InitAll
    Me.DSJointCoor.ClearData
    Me.DSFrameData.ClearData
    
    Me.DSFrameForce.ClearData Me.DSFrameForce.refRow + 2
    Me.DSFrameForce.ClearColumns "eleID", "station", "loadComb", "stepType", _
                        "P", "V2", "V3", "T", "M2", "M3"
    'Me.DSFrameForce.ClearColumnData "section", Me.DSFrameForce.refRow + 2
   ' Me.DSFrameForce.ClearColumnData "memID", Me.DSFrameForce.refRow + 2
    
    Me.DSMemberData.ClearData
    Me.DSJointConnectivity.ClearData
    With DSConnectionType.tagSelector
        Me.DSConnectionType.ClearColumnData .matchedJoint
    End With
    
    Dim DsSys As DataSheetSystem
    Set DsSys = New DataSheetSystem
    If isClearSavedUserInput Then DsSys.DeleteAllCustomPropGrp
        
    DsSys.RestoreToDefault

End Sub

Private Function DSName(dsType As ModelDataSheetType) As String
    Dim str As String
    Select Case dsType
        Case JointCoor: str = "JointCoor"
        Case FrameData: str = "FrameData"
        Case FrameForce: str = "FrameForce"
        Case MemberData: str = "MemberData"
        Case JointConnectivity: str = "JointConnectivity"
        Case ConnectionType: str = "ConnectionType"
    End Select
    
    DSName = str
End Function

Private Function DSEnum(str As String) As ModelDataSheetType
    Dim dsType As ModelDataSheetType
    Select Case str
        Case "JointCoor": dsType = JointCoor
        Case "FrameData": dsType = FrameData
        Case "FrameForce": dsType = FrameForce
        Case "MemberData": dsType = MemberData
        Case "JointConnectivity": dsType = JointConnectivity
        Case "ConnectionType": dsType = ConnectionType
    End Select
    DSEnum = dsType
End Function

Public Property Get DSJointCoor() As oDataSheet
    If Not isKeyExist(mDataSheets, DSName(JointCoor)) Then
        InitDataSheetJoint
    End If
    Set DSJointCoor = mDataSheets(DSName(JointCoor))
End Property

Public Property Get DSFrameData() As oDataSheet
    If Not isKeyExist(mDataSheets, DSName(FrameData)) Then
        InitDataSheetFrameData
    End If
    Set DSFrameData = mDataSheets(DSName(FrameData))
End Property

Public Property Get DSMemberData() As oDataSheet
    If Not isKeyExist(mDataSheets, DSName(MemberData)) Then
        InitDataSheetMemberData
    End If
    Set DSMemberData = mDataSheets(DSName(MemberData))
End Property

Public Property Get DSFrameForce() As oDataSheet
    If Not isKeyExist(mDataSheets, DSName(FrameForce)) Then
        InitDataSheetFrameForce
    End If
    Set DSFrameForce = mDataSheets(DSName(FrameForce))
End Property

Public Property Get DSJointConnectivity() As oDataSheet
    If Not isKeyExist(mDataSheets, DSName(JointConnectivity)) Then
        InitDataSheetJointConnectivity
    End If
    Set DSJointConnectivity = mDataSheets(DSName(JointConnectivity))
End Property

Public Property Get DSConnectionType() As oDataSheet
    If Not isKeyExist(mDataSheets, DSName(ConnectionType)) Then
        InitDataSheetConnectionType
    End If
    Set DSConnectionType = mDataSheets(DSName(ConnectionType))
End Property

Private Function isKeyExist(coll As Collection, sKey As String) As Boolean
    'Input shall be colleciton of OBJECT
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


