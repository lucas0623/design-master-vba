VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MapByJointName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Operation.ConnAnalysis")

Implements IMapConnMethod
Private mModel As StrModel
Private mJts As Collection 'all imported joints data in the dataSheet


Private Sub IMapConnMethod_Initialize(model As StrModel)
    Set mModel = model
    Set mJts = mModel.joints
        
End Sub

Private Function IMapConnMethod_MapConnection(connType As StrConnectionType) As Integer
    g_log.WriteLog "Mapping Connection Type = " & connType.Name & "; Method = 'Map By Specified Joint Name'"
    Dim mapJtsName() As String
    
    Dim dataStrOper As libDataString
    Set dataStrOper = New libDataString
    
    mapJtsName = dataStrOper.SplitAndTrimString(connType.mapJtsName, ",")
    
    Dim i As Long, jt As StrJoint
    For i = LBound(mapJtsName) To UBound(mapJtsName)
        If mModel.IsStrObjectExist(mapJtsName(i), obj_jt) Then
            Set jt = mModel.GetStrObject(mapJtsName(i), obj_jt)
            jt.AddConnectionTypes connType
            connType.AddMatchedJoints jt
            g_log.WriteLog "    Joint '" & mapJtsName(i) & "' mapped."
        Else
            g_log.WriteLog "    Joint '" & mapJtsName(i) & "' cannot be found in the data. Record skipped. Please check your input."
        End If
    Next i
End Function

