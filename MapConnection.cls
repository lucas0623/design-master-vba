VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MapConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Operation.ConnAnalysis")

Private mModel As New StrModel
Private mDsSys As New DataSheetSystem
Private mCompleteMessageAddition As String
Private mTerminateMessageMain As String, mTerminateMessageAddition As String

Private MapByJointName As New MapByJointName
Private MapByContainsSecAndRestr As New MapByContainsSecAndRestr
Private MapByMatchAllSecAndRestr As New MapByMatchAllSecAndRestr
Private MapByContainsOnly As New MapByContainsOnly

Public Property Get CompleteMessageAddition() As String
    CompleteMessageAddition = mCompleteMessageAddition
End Property

Public Property Get TerminateMessageMain() As String
    TerminateMessageMain = mTerminateMessageMain
End Property

Public Property Get TerminateMessageAddition() As String
    TerminateMessageAddition = mTerminateMessageAddition
End Property

Public Function Main() As Integer
    Dim ret As Integer
    
    ret = CheckDataBeforeStart
    If Not ret = 0 Then
        mTerminateMessageMain = "no data is found in the workbook! The macro will be terminated."
        GoTo ExitFunction
    End If
    
    g_log.WriteLog "Connection Matching Started."
    ret = FormStrObj
    If Not ret = 0 Then
        GoTo ExitFunction
    End If
    
    ret = MatchConn
    If Not ret = 0 Then
        mTerminateMessageMain = "no Connection Type is defined in the workbook! The macro will be terminated." _
                        & " Plesae go to the worksheet 'Define_ConnectionType' to define the connection type."
        GoTo ExitFunction
    End If
    WriteData
    g_log.WriteLog "Connection Matching Completed."
    mDsSys.prop("isCreated", "isMappedConn") = True
ExitFunction:
    Main = ret
    Application.StatusBar = False
End Function


Private Function CheckDataBeforeStart() As Integer
    If Not mDsSys.prop("WhatOperationCanProcess", "MapConnection") Then
        CheckDataBeforeStart = -1
    End If
End Function

Private Function FormStrObj() As Integer
    Dim ret As Integer
    With mModel.Constructor
        ret = .FormJointObj
        If Not ret = 0 Then GoTo ExitFunction
        
        ret = .FormJointObjForConnMap
        If Not ret = 0 Then GoTo ExitFunction
        
        ret = .FormFrmObj
        If Not ret = 0 Then GoTo ExitFunction
        
        ret = .FormMemberObj
        If Not ret = 0 Then GoTo ExitFunction
        
        ret = .FormConnTypeObjForMatching
        If Not ret = 0 Then GoTo ExitFunction
    End With
ExitFunction:
    FormStrObj = ret
End Function

Private Function MatchConn() As Integer
    
    Dim i As Long
    Dim connType As StrConnectionType, connTypes As Collection
    Dim oMapMethod As IMapConnMethod
    Dim ret As Integer, numOfConnTypes As Long
    Set connTypes = mModel.conns
    
    numOfConnTypes = connTypes.count
    If numOfConnTypes = 0 Then
        ret = -1
        GoTo ExitFunction
    End If
    
    'InitializeMappingMethod
    
    For i = 1 To numOfConnTypes
        Application.StatusBar = "Start Mapping Connection...Progress = " & i & "/" & numOfConnTypes
        Set connType = connTypes(i)
        Set oMapMethod = SelectMappingMethod(connType.mapMethod, mModel, connType)
        If Not oMapMethod Is Nothing Then
            oMapMethod.MapConnection connType
            g_log.WriteLog "  Connection Type '" & connType.Name & "' matching completed."
        End If
    Next
    
ExitFunction:
    MatchConn = ret
End Function

Private Sub WriteData()
    
    Dim dsManager As DataSheetManager
    Dim dsJtConnectivity As oDataSheet, dsConnType As oDataSheet
    Dim df As clsDataFrame
    
    Set dsManager = New DataSheetManager
    Set dsJtConnectivity = dsManager.DSJointConnectivity
    Set dsConnType = dsManager.DSConnectionType
    
    Set df = mModel.GetDataframe(obj_jt, "matchedConnectionStr")
    With dsJtConnectivity.tagSelector
        dsJtConnectivity.ClearColumnData .matchedType
        dsJtConnectivity.WriteDataframe df, True, False, .matchedType
    End With
    
    Set df = mModel.GetDataframe(obj_connection, "jointsName")
    With dsConnType.tagSelector
        dsConnType.ClearColumnData .matchedJoint
        dsConnType.WriteDataframe df, True, False, .matchedJoint
    End With
End Sub
'Private Sub InitializeMappingMethod()
''    MapByJointName.Initialize mModel
''    MapByContainsSecAndRestr.Initialize mModel
''    MapByMatchAllSecAndRestr.Initialize mModel
''    MapByContainsOnly.Initialize mModel
'End Sub
Private Function SelectMappingMethod(sMapMethod As String, model As StrModel, connType As StrConnectionType) As IMapConnMethod
    Dim oMapMethod As IMapConnMethod
    sMapMethod = UCase(sMapMethod)
    
    If Not connType.mapJtsName = vbNullString Then
        Set oMapMethod = MapByJointName
    ElseIf sMapMethod = "CONTAINS" Then
        Set oMapMethod = MapByContainsSecAndRestr
    ElseIf sMapMethod = "MATCH ALL" Then
        Set oMapMethod = MapByMatchAllSecAndRestr
    ElseIf sMapMethod = "CONTAINS ONLY" Then
        Set oMapMethod = MapByContainsOnly
    Else
        g_log.RaiseWarning "Cannot regconize connection mapping method for Connection Type '" & _
                            connType.Name & "'. Record Skip", failToRegconizeConnMatchMethod
        Exit Function
    End If
    
    oMapMethod.Initialize model
    Set SelectMappingMethod = oMapMethod
End Function



Public Property Get completeMsg() As String
    completeMsg = mCompleteMsg
End Property
