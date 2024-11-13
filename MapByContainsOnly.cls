VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MapByContainsOnly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Operation.ConnAnalysis")

Implements IMapConnMethod
Private mModel As StrModel
Private dataStrOper As New libDataString, genFunc As New clsGeneralFunctions
Private DictForMap As Object

Private Sub IMapConnMethod_Initialize(model As StrModel)
    Set mModel = model
    'Set DictForMap = SetDictForMap
End Sub

Private Function IMapConnMethod_MapConnection(connType As StrConnectionType) As Integer
    
    If DictForMap Is Nothing Then Set DictForMap = SetDictForMap
    Dim mapSectionsName As String 'Sections Name in the connection Type
    'Dim jtSectionsName As String 'sections name of the connected frames in a jt
    If Not genFunc.isInitialised(connType.mapSections) Then
        g_log.WriteLogInDetailMode "  No Mapping Section Input for Connection Type '" & connType.Name
        GoTo ExitFunction:
    End If
    mapSectionsName = dataStrOper.SortAndRemoveDuplicateStr(connType.mapSections, ",")
    mapSectionsName = mapSectionsName & CStr(connType.isRestraint)
    
    g_log.WriteLog "Mapping Connection Type = " & connType.Name & "; Method = 'Map By Contains Only'"
    g_log.WriteLog "  Sections = " & connType.mapSections
    
    Dim i As Long, j As Long, jt As StrJoint
    
    
'    If ConnType.isRestraint = jt.isRestraint And jt.isConn Then
        If Not DictForMap.Exists(mapSectionsName) Then
            g_log.WriteLogInDetailMode "  No Joint is mapped for  '" & connType.Name & "'"
            GoTo ExitFunction
        End If
'    Else
'        g_log.WriteLogInDetailMode "  Joint '" & jt.name & "': Restraint Condition not matched or this joint is not connection."
'        GoTo ExitFunction
'    End If
    

    For i = 1 To DictForMap(mapSectionsName).count
        Set jt = DictForMap(mapSectionsName)(i)
        jt.AddConnectionTypes connType
        connType.AddMatchedJoints jt
        g_log.WriteLog "  Mapped Joint '" & jt.Name & "'"
    Next i
ExitFunction:

End Function

Private Function SetDictForMap() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    'loop for each jt
    'get the sections of all connected frames in array, removed duplicate, and sort, and convert to a string (=key)
    'item in dict shall be collection of jt
    'create new key if key not exist
    'add jt to the dict
    
    Dim key As String
    Dim jt As StrJoint, jts As Object
    Set jts = mModel.joints
    For Each jt In jts
    If Not jt.isConn Then GoTo NextIteration
        key = jt.SortedConnectedFramesSectionStr(vbNullString, True)
        key = key & CStr(jt.isRestraint)
        If dict.Exists(key) Then
            dict(key).Add jt
        Else
            Dim coll As Collection
            Set coll = New Collection

            coll.Add jt
            dict.Add key, coll
        End If
NextIteration:
    Next
    
    Set SetDictForMap = dict
End Function


