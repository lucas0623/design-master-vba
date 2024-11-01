VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtractConnForceMethodCo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Operation.ExtractConnectionForce")

'Loop for each connection type
'depends of 'isByNode', if true, loop for each node. otherwise get all frameforce of all node. need check if node exist in the 'Selected Node list'
'each node
'See if 'isEachLC' -> further breakdown
'For each section if matched -->
'Find Max/Min and correspondence cases

Implements IExtractFrmForceMethod
Private collMethod As New LibCollectionMethod, genFunc As New clsGeneralFunctions
Private mUFControl As VMForceExtraction
Private lc() As String, NodesName() As String, ConnTypeNames() As String 'for use in Userform
Private NodesColl As New Collection 'for checking if node is filtered
'Private PosFilterType As PositionFilterType, posFilterVal As Double
Private IsByNode As Boolean 'true = by section, false = by member
Private IsEachLC As Boolean
Private ExtremeCases() As Boolean
Private MaxMinEnv() As Boolean
Private mModel As StrModel

Private Sub IExtractFrmForceMethod_Initialize(model As StrModel, UFControl As Object)
    Set mUFControl = UFControl
    Set mModel = model
End Sub

Private Function GetUserInput() As Integer
    Dim ret As Integer
    With mUFControl
        'Set ws_sum = .ws_sum
        lc = .lc
        ConnTypeNames = .connTypes
        NodesName = .Nodes
        'PosFilterType = .PosFilterType
        'posFilterVal = .PosFilterValue
        IsByNode = .IsByNode
        'isExtremeCase = True
        IsEachLC = .IsEachLC
        ExtremeCases = .ExtremeCase
        MaxMinEnv = .MaxMinEnv
    End With
    GetNodeDict
End Function

Private Function GetNodeDict()
    Dim i As Long, jt As StrJoint
    For i = LBound(NodesName) To UBound(NodesName)
        Set jt = mModel.GetStrObject(NodesName(i), obj_jt)
        NodesColl.Add jt, jt.name
    Next i
End Function

Private Function ExtractConnForceByCo() As Collection
    GetUserInput
    
    Dim cConnType As StrConnectionType, matchedNodes() As String
    Dim cJt As StrJoint
    Dim frmForces As Collection, reFrmForcesNotConverted As Collection, reFrmForces As New Collection
    Dim i As Long, j As Long, k As Long
    
    g_log.WriteLog "Current Procedure: ExtractConnForce.ExtractConnForceMethodCo"
    g_log.WriteLog "IsByNode = " & CStr(IsByNode)
    g_log.WriteLog "Total number of Connection Types = " & UBound(ConnTypeNames) + 1
    
    For i = LBound(ConnTypeNames) To UBound(ConnTypeNames)
        Set cConnType = mModel.GetStrObject(ConnTypeNames(i), obj_connection)
        
        matchedNodes = cConnType.jointsNameArr

        g_log.WriteLog "Start extracting forces for no. " & i + 1 & "/" & UBound(ConnTypeNames) + 1 & " Connection Types."
        Application.StatusBar = "Start extracting forces for no. " & i + 1 & "/" & UBound(ConnTypeNames) + 1 & " Connection Types."
        
        If IsByNode Then
            For j = LBound(matchedNodes) To UBound(matchedNodes)
                If Not collMethod.isKeyExist(NodesColl, matchedNodes(j)) Then GoTo Nextj
                Set cJt = mModel.GetStrObject(matchedNodes(j), obj_jt)
                Set frmForces = New Collection
                Set reFrmForcesNotConverted = New Collection
                
                Set frmForces = GetFrameForcesCollForOneNode(cConnType, cJt)
                If frmForces.count = 0 Then
                    g_log.RaiseWarning "WARNING: Fail to match any node for joint '" & _
                                    matchedNodes(k) & "' of Connection Type '" & ConnTypeNames(i), failToMatchAnySectionForConnForceExtract
                    GoTo Nextj
                End If
                Set reFrmForcesNotConverted = GetExtremeCases(frmForces, ExtremeCases, MaxMinEnv)
                collMethod.AddColltoColl reFrmForces, ConvertToFrmForceForConn(cConnType, reFrmForcesNotConverted)
Nextj:
            Next j
        Else
            Set frmForces = New Collection
            Set reFrmForcesNotConverted = New Collection
            For k = LBound(matchedNodes) To UBound(matchedNodes)
                If Not collMethod.isKeyExist(NodesColl, matchedNodes(k)) Then GoTo Nextk
                Set cJt = mModel.GetStrObject(matchedNodes(k), obj_jt)
                collMethod.AddColltoColl frmForces, GetFrameForcesCollForOneNode(cConnType, cJt)
Nextk:
            Next k
            If frmForces.count = 0 Then
                g_log.RaiseWarning "WARNING: Fail to match any node for joint '" & _
                                    matchedNodes(k) & "' of Connection Type '" & ConnTypeNames(i), failToMatchAnySectionForConnForceExtract
                GoTo Nexti
            End If
            Set reFrmForcesNotConverted = GetExtremeCases(frmForces, ExtremeCases, MaxMinEnv)
            collMethod.AddColltoColl reFrmForces, ConvertToFrmForceForConn(cConnType, reFrmForcesNotConverted)
            
        End If
Nexti:
    Next i
    
    Set ExtractConnForceByCo = reFrmForces
End Function

Private Function GetFrameForcesCollForOneNode(connType As StrConnectionType, node As StrJoint) As Collection
    
    Dim frmForces As New Collection, cConnectedFrm As StrFrame
    Dim extractSections() As String
    Dim retFrmForces As New Collection
    'Set cJt = mModel.GetStrObject(matchedNodes(j), obj_jt)
    extractSections = Split(connType.extractSections, ",")
    For Each cConnectedFrm In node.connectedFrames
        If genFunc.IsInArr(cConnectedFrm.section, extractSections) Then
            If cConnectedFrm.jtI.name = node.name Then
                Set frmForces = cConnectedFrm.GetFrameForces_EndI(lc)
                collMethod.AddColltoColl retFrmForces, frmForces
            Else
                Set frmForces = cConnectedFrm.GetFrameForces_EndJ(lc)
                collMethod.AddColltoColl retFrmForces, frmForces
            End If
        End If
    Next
    Set GetFrameForcesCollForOneNode = retFrmForces
End Function

Private Function GetExtremeCases(cFrmForces As Collection, ExtremeCase() As Boolean, _
                            isMaxMinEnv() As Boolean)
    Set GetExtremeCases = mModel.frmForceAnalyser.FindExtremeForceWithMaxMinEnv(cFrmForces, ExtremeCase, isMaxMinEnv)
End Function

Private Function ConvertToFrmForceForConn(connType As StrConnectionType, coll As Collection) As Collection

    Dim frmForce As StrFrameForce
    Dim collFrmForceForConn As New Collection, oFrmForceForConn As New StrFrameForceForConn
    Dim count As Integer
    'Dim identificationStr As String, currentStr As String
    Dim jt As StrJoint

    
    For Each frmForce In coll
        Set oFrmForceForConn = New StrFrameForceForConn
        'Find the jt
        If frmForce.pos_fromEleJtI_percent Then
            Set jt = frmForce.frame.jtI
        Else
            Set jt = frmForce.frame.jtJ
        End If
             
        oFrmForceForConn.Initialize connType, jt, frmForce
        
NextIteration:
        collFrmForceForConn.Add oFrmForceForConn
    Next
    
    Set ConvertToFrmForceForConn = collFrmForceForConn
End Function

Private Function IExtractFrmForceMethod_ExtractForce() As Collection
    Set IExtractFrmForceMethod_ExtractForce = ExtractConnForceByCo
End Function


