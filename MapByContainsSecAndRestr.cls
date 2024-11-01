VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MapByContainsSecAndRestr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Operation.ConnAnalysis")

Implements IMapConnMethod
Private mModel As StrModel
'Private mJts As Collection 'all imported joints data in the dataSheet
Private dataStrOper As New libDataString, genFunc As New clsGeneralFunctions
Private DictForMap As Object
Private DictFunc As New libDictionaryFunctions
    
Private Sub IMapConnMethod_Initialize(model As StrModel)
    Set mModel = model
    Set mJts = mModel.joints
End Sub

Private Function IMapConnMethod_MapConnection(connType As StrConnectionType) As Integer

    If DictForMap Is Nothing Then Set DictForMap = SetDictForMap
    
    Dim mapSectionsName() As String 'Sections Name in the connection Type
    Dim i As Long, j As Long, jt As StrJoint, count As Long
    Dim lLimit As Long, uLimit As Long
    'Dim jtSectionsName() As String 'sections name of the connected frames in a jt
    
    mapSectionsName = dataStrOper.SplitAndTrimString(connType.mapSections, ",")
    If Not genFunc.isInitialised(mapSectionsName) Then
        g_log.WriteLogInDetailMode "  No Mapping Section Input for Connection Type '" & connType.name
        GoTo ExitFunction:
    End If
    
    lLimit = LBound(mapSectionsName)
    uLimit = UBound(mapSectionsName)
    For i = lLimit To uLimit
        mapSectionsName(i) = mapSectionsName(i) & CStr(connType.isRestraint)
    Next i
    
    g_log.WriteLog "Mapping Connection Type = " & connType.name & "; Method = 'Map By contains sections and restraints'"
    g_log.WriteLog "  Sections = " & connType.mapSections
    
    
    Dim DictOfMatched As Object, dict2 As Object
    Set DictOfMatched = CreateObject("Scripting.Dictionary")
    Set dict2 = CreateObject("Scripting.Dictionary")
    count = uLimit - lLimit + 1
    If count = 1 Then
        If DictForMap.Exists(mapSectionsName(lLimit)) Then Set DictOfMatched = DictForMap(mapSectionsName(lLimit))
    ElseIf count = 2 Then
        If DictForMap.Exists(mapSectionsName(lLimit)) Then Set DictOfMatched = DictForMap(mapSectionsName(lLimit))
        If DictForMap.Exists(mapSectionsName(uLimit)) Then Set dict2 = DictForMap(mapSectionsName(uLimit))
        Set DictOfMatched = DictFunc.GetDuplicateItemsInTwoDicts(DictOfMatched, dict2)
    Else
        If DictForMap.Exists(mapSectionsName(lLimit)) Then Set DictOfMatched = DictForMap(mapSectionsName(lLimit))
        If DictForMap.Exists(mapSectionsName(lLimit + 1)) Then Set dict2 = DictForMap(mapSectionsName(uLimit))
        Set DictOfMatched = DictFunc.GetDuplicateItemsInTwoDicts(DictOfMatched, dict2)
        For i = lLimit + 2 To uLimit
            If DictForMap.Exists(mapSectionsName(i)) Then
                Set dict2 = DictForMap(mapSectionsName(i))
                Set DictOfMatched = DictFunc.GetDuplicateItemsInTwoDicts(DictOfMatched, dict2)
            Else
                Set DictOfMatched = Nothing
                g_log.WriteLogInDetailMode "  Cannot find key  '" & mapSectionsName(i) & "' in the data!!"
                GoTo ExitFunction
            End If
        Next i
    End If
    
    If DictOfMatched.count = 0 Then
       g_log.WriteLogInDetailMode "  No Joint is mapped for  '" & connType.name & "'"
        GoTo ExitFunction
    End If
    
    Dim key As Variant
    For Each key In DictOfMatched.keys
        Set jt = DictOfMatched(key)
        jt.AddConnectionTypes connType
        connType.AddMatchedJoints jt
        g_log.WriteLog "  Mapped Joint '" & jt.name & "'"
    Next
            
ExitFunction:

End Function

Private Function SetDictForMap() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    'loop for each jt

    
    Dim key As String
    Dim jt As StrJoint, jts As Object
    Dim frm As StrFrame, frms As Object
    Set jts = mModel.joints
    For Each jt In jts
        If Not jt.isConn Then GoTo NextIteration
        Set frms = jt.connectedFrames
        For Each frm In frms
            key = frm.section & CStr(jt.isRestraint)
            If dict.Exists(key) Then
                If Not dict(key).Exists(jt.name) Then dict(key).Add jt.name, jt
            Else
                Dim dict2 As Object
                Set dict2 = CreateObject("Scripting.Dictionary")
                dict2.Add jt.name, jt
                dict.Add key, dict2
            End If
        Next
NextIteration:
    Next
    
    Set SetDictForMap = dict
End Function
