VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GetModelConnectionType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Operation.ConnAnalysis")
'The model will read all Identified Connection data,
'and write the result to a new sheet

'This operation is used mainly for the first time creating all connection data in excel.


Private mModel As New StrModel
Private mDsSys As New DataSheetSystem
Private mCompleteMsgAddition As String
Private mTerminateMsgMain As String, mTerminateMsgAdditional As String

Public Function Main() As Integer
    Dim ret As Integer

    ret = CheckDataBeforeStart
    If Not ret = 0 Then
        mTerminateMsgMain = "no data is found in the workbook! The macro will be terminated."
        GoTo ExitFunction
    End If

    g_log.WriteLog "Getting Model Identified Connection Type."
    ret = FormStrObj
    If Not ret = 0 Then
        GoTo ExitFunction
    End If

    ret = CreateData
    If Not ret = 0 Then
        mTerminateMsgMain = "Unexpected Error in GetModelConnectionType.Main!! "
        GoTo ExitFunction
    End If
    WriteData
    g_log.WriteLog "Get Identified Connection Type Completed."

ExitFunction:
    Main = ret
    Application.StatusBar = False
End Function

Public Property Get CompleteMessageAddition() As String
    CompleteMessageAddition = mCompleteMessageAddition
End Property

Public Property Get TerminateMessageMain() As String
    TerminateMessageMain = mTerminateMessageMain
End Property

Public Property Get TerminateMessageAddition() As String
    TerminateMessageAddition = mTerminateMessageAddition
End Property

Private Function CheckDataBeforeStart() As Integer
    If Not mDsSys.prop("WhatOperationCanProcess", "GetIdentifiedConnType") Then
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

        ret = .FormFrmObj(isAllData:=True)
        If Not ret = 0 Then GoTo ExitFunction

        ret = .FormMemberObj
        If Not ret = 0 Then GoTo ExitFunction

    End With
ExitFunction:
    FormStrObj = ret
End Function

Private Function CreateData() As Integer

    Dim count As Long
    Dim jt As StrJoint, jts As Collection
    Dim frameSections As Object 'arraylist to store and sort frameSections
    Dim connType As StrConnectionType, connTypeKey As String
    Dim connTypes As Object 'a dictionary
    Dim ret As Integer

    'Loop Through Each Joint
    'Check if isConn, if false, next joint
    'Get the connected sections name array, sort the array, and condense to a string
    'ConnTypes will be a dictionary. key = sortedConnectedSectionsName.
    'Check if key exist, if yes, find if not, add item

    Set jts = mModel.joints

    If jts.count = 0 Then
        ret = -1
        GoTo ExitFunction
    End If
    count = 1
    Set frameSections = CreateObject("System.Collections.ArrayList")
    Set connTypes = CreateObject("Scripting.Dictionary")
    For Each jt In jts
        Application.StatusBar = "Start Searching Connection...Progress = " & count & "/" & jts.count

        If Not jt.isConn Then GoTo NextIteration
        connTypeKey = jt.SortedConnectedFramesSectionStr

        If connTypes.exists(connTypeKey) Then
            Set connType = mModel.GetStrObject(connTypeKey, obj_connection)
            connType.AddMatchedJoints jt
        Else
            Set connType = New StrConnectionType
            connType.Init connTypeKey
            connType.AddMatchedJoints jt
            connTypes.Add connTypeKey, connType
            ret = mModel.AddStrObjToColl(connType, obj_connection)
        End If
        count = count + 1
NextIteration:
    Next

ExitFunction:
    CreateData = ret
End Function

Private Sub WriteData()

    Dim dsManager As DataSheetManager
    Dim ws As Worksheet
    'Dim dsJtConnectivity As oDataSheet, dsConnType As oDataSheet
    Dim df As clsDataFrame

    Set dsManager = New DataSheetManager
    'Set dsJtConnectivity = dsManager.DSJointConnectivity
    'Set dsConnType = dsManager.DSConnectionType


    Set df = mModel.GetDataframe(obj_connection, "jointsName", "name", "numOfVertFrames", _
                                "numOfDiagonalFrames", "numOfHorizontalFrames")
    Set ws = Sheets.Add
    df.WriteToRange ws, Range("A2"), False, True
'    With dsConnType.tagSelector
'        dsConnType.ClearColumnData .matchedJoint
'        dsConnType.WriteDataframe df, dsConnType.refRow + 1, .matchedJoint
'    End With
End Sub




