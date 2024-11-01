VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtractConnForceMethodAllNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Operation.ExtractConnectionForce")
Implements IExtractFrmForceMethod

Private genFunc As clsGeneralFunctions
Private mUFControl As VMForceExtraction
Private mModel As StrModel

'Logic for Connection force extraction
'Loop through each selected connection type
'Loop through each matched nodes in the type. skip excluded node.
'for each node, select option of extract all/ each

'loop each load comb
'loop thorugh each connected frames. if connected frame section = output section, then extract force.
'extract force: get frame force obj from the frame (apply lc & position filter).
'envelope force according to setting
'sort by permutation of the envelope forces for each load comb

Private Function ExtractForce() As Collection
    Dim selectedConnTypes() As String, selectedNodes() As String, selectedLc() As String, isMaxMinEnv() As Boolean 'from Userform
    Dim dictSelectedNodes As Object
    Dim cConnType As StrConnectionType, matchedNodes() As String, extractSections() As String
    Dim cJt As StrJoint
    Dim cConnectedFrm As StrFrame
    Dim frmForces As New Collection, frmForcesOfOneLC As Collection, numFrames As Long, frmForcesBeforeConvert As New Collection
    Dim i As Long, j As Long, k As Long
    
    selectedConnTypes = mUFControl.connTypes
    selectedNodes = mUFControl.Nodes
    selectedLc = mUFControl.lc
    isMaxMinEnv = mUFControl.MaxMinEnv
    
    Set dictSelectedNodes = CreateObject("Scripting.Dictionary")
    For i = LBound(selectedNodes) To UBound(selectedNodes)
        dictSelectedNodes.Add selectedNodes(i), selectedNodes(i)
    Next i
    
    For i = LBound(selectedConnTypes) To UBound(selectedConnTypes)
        Set cConnType = mModel.GetStrObject(selectedConnTypes(i), obj_connection)
        matchedNodes = cConnType.jointsNameArr
        extractSections = Split(cConnType.extractSections, ",")
        For j = LBound(matchedNodes) To UBound(matchedNodes)
            If Not dictSelectedNodes.Exists(matchedNodes(j)) Then GoTo Nextj
            Set cJt = mModel.GetStrObject(matchedNodes(j), obj_jt)
            For k = LBound(selectedLc) To UBound(selectedLc)
                Set frmForcesOfOneLC = New Collection
                numFrames = 0
                For Each cConnectedFrm In cJt.connectedFrames
                    If genFunc.IsInArr(cConnectedFrm.section, extractSections) Then
                        AddColltoColl frmForcesOfOneLC, GetFrameForceAtNode(cJt, cConnectedFrm, selectedLc(k), isMaxMinEnv)
                        numFrames = numFrames + 1
                    End If
                Next
                If frmForcesOfOneLC.count > 0 Then
                    Set frmForcesBeforeConvert = SortFrameForceByMaxMinEnv(frmForcesOfOneLC, numFrames)
                    AddColltoColl frmForces, ConvertToFrmForceForConn(cJt, cConnType, frmForcesBeforeConvert)
                End If
            Next k
Nextj:
        Next j
    Next i
    Set ExtractForce = frmForces
End Function

Private Function GetFrameForceAtNode(node As StrJoint, frm As StrFrame, lc As String, isMaxMinEnv() As Boolean) As Collection
'to get the frame forces collection of certain frame at certain node of certain load combination. apply envlope option
     Dim frmForces As New Collection
     'determine which end of the frame is connected to the node
     If frm.jtI.name = node.name Then
        Set frmForces = frm.GetFrameForces_EndI(lc)
    Else
        Set frmForces = frm.GetFrameForces_EndJ(lc)
    End If
    
    Set GetFrameForceAtNode = mModel.frmForceAnalyser.EnvMaxMin(frmForces, isMaxMinEnv)
End Function

Private Function SortFrameForceByMaxMinEnv(coll As Collection, numFrames As Long) As Collection
    If numFrames = 1 Then
        Set SortFrameForceByMaxMinEnv = coll
    Else
        Dim collSize As Long, numPermutation As Long
        Dim retColl As New Collection
        Dim i As Long, j As Long
        collSize = coll.count
        numPermutation = collSize / numFrames
        For i = 1 To numPermutation
            For j = 1 To numFrames
                retColl.Add coll.item(i + numPermutation * (j - 1))
            Next j
        Next i
        Set SortFrameForceByMaxMinEnv = retColl
    End If
End Function

Private Function ConvertToFrmForceForConn(jt As StrJoint, connType As StrConnectionType, coll As Collection) As Collection

    Dim frmForce As StrFrameForce
    Dim collFrmForceForConn As New Collection, oFrmForceForConn As New StrFrameForceForConn
    Dim count As Integer
    Dim identificationStr As String, currentStr As String
    
    For Each frmForce In coll
        Set oFrmForceForConn = New StrFrameForceForConn
        
        'dealing with isFullyPopulate Option
        If connType.isFullyPopulate Then
            oFrmForceForConn.Initialize connType, jt, frmForce
        Else
            currentStr = connType.name & connType.jointsName & frmForce.loadComb & frmForce.stepType
            
            If count = 0 Then
                identificationStr = currentStr
                oFrmForceForConn.Initialize connType, jt, frmForce
                count = count + 1
                GoTo NextIteration
            End If
            
            If currentStr = identificationStr Then
                oFrmForceForConn.Initialize connType, jt, frmForce, vbNullString, vbNullString, vbNullString
            Else
                identificationStr = currentStr
                oFrmForceForConn.Initialize connType, jt, frmForce
            End If
            
        End If
        
NextIteration:
        collFrmForceForConn.Add oFrmForceForConn
    Next
    
    Set ConvertToFrmForceForConn = collFrmForceForConn
End Function

Private Sub IExtractFrmForceMethod_Initialize(model As StrModel, UFControl As Object)
    Set mUFControl = UFControl
    Set mModel = model
    Set genFunc = New clsGeneralFunctions
End Sub

Private Function IExtractFrmForceMethod_ExtractForce() As Collection
    Set IExtractFrmForceMethod_ExtractForce = ExtractForce
End Function

Private Sub AddColltoColl(coll As Collection, coll2 As Collection)
    Dim i As Long
    For i = 1 To coll2.count
        coll.Add coll2(i)
    Next i
End Sub
