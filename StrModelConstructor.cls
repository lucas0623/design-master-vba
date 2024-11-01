VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StrModelConstructor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 '@Folder "Structural Model"

'********************************************************
'This module controls the formation of model objects
'Arthor: Lucas LEUNG
'Update Log

'07 Jun 2023 - Initial
'*******************************************************

Option Explicit
Private wsInteract As clsWorksheetsInteraction
Private genFunc As clsGeneralFunctions
Private mModel As StrModel, mDsManager As DataSheetManager
Private mDsJt As oDataSheet, mDsFrm As oDataSheet
Private mDsFrmForce As oDataSheet, mDsMember As oDataSheet
Private mDsConnType As oDataSheet, mDsJtConn As oDataSheet

Private mDfJt As clsDataFrame, mDfFrm As clsDataFrame
Private mDfFrmForce As clsDataFrame, mDfMember As clsDataFrame
Private mDfConnType As clsDataFrame

Public Sub Initialize(model As StrModel)
    Set mModel = model
    Set wsInteract = New clsWorksheetsInteraction
    Set genFunc = New clsGeneralFunctions
    Set mDsManager = New DataSheetManager
    Set mDsJt = mDsManager.DSJointCoor
    Set mDsFrm = mDsManager.DSFrameData
    Set mDsFrmForce = mDsManager.DSFrameForce
    Set mDsMember = mDsManager.DSMemberData
    Set mDsConnType = mDsManager.DSConnectionType
    Set mDsJtConn = mDsManager.DSJointConnectivity
End Sub


Public Function FormJointObj() As Integer
    'return 0 if formed joint obj.
    'return -1 if not formed
    If Not mModel.joints.count = 0 Then GoTo ExitFunction
    
    Dim dsJt As oDataSheet
    Dim df As clsDataFrame, jt As StrJoint
    Dim col_name As Long, col_x As Long, col_y As Long, col_z As Long, col_isRestraint As Long 'column number in the dataframe
    Dim i As Long
    Dim ret As Integer
    Set dsJt = mDsJt
    Set df = Me.DfJt
    
    If df.isInitialized Then
        With dsJt.tagSelector
            col_name = df.columnNum(.ID)
            col_x = df.columnNum(.x)
            col_y = df.columnNum(.y)
            col_z = df.columnNum(.z)
            col_isRestraint = df.columnNum(.isRestraint)
        End With
        
        For i = 1 To df.CountRows
            Set jt = New StrJoint
            jt.Init df.idata(i, col_name), df.idata(i, col_x), df.idata(i, col_y), df.idata(i, col_z)
            jt.isRestraint = df.idata(i, col_isRestraint)
            ret = mModel.AddStrObjToColl(jt, obj_jt)
        Next i
    Else
        ret = -1
    End If
    'g_log.WriteLog "Number of Joint Objects Formed = " & mModel.joints.count
ExitFunction:
    FormJointObj = ret
End Function

Public Function FormJointObjForConnMap() As Integer
    FormJointObj
    'return 0 if formed joint obj.
    'return -1 if not formed
    Dim dsJtConn As oDataSheet
    Dim df As clsDataFrame, jt As StrJoint
    Dim col_name As Long, col_isConn As Long
    Dim i As Long
    Dim ret As Integer
    Set dsJtConn = mDsJtConn
    
    With dsJtConn.tagSelector
        Set df = mDsJtConn.GetDataframe(.jtID, .isConn)
    End With
    
    If df.isInitialized Then
        With dsJtConn.tagSelector
            col_name = df.columnNum(.jtID)
            col_isConn = df.columnNum(.isConn)
        End With
        
        For i = 1 To df.CountRows
            Set jt = mModel.GetStrObject(df.idata(i, col_name), obj_jt)
            jt.isConn = df.idata(i, col_isConn)
        Next i
    Else
        FormJointObjForConnMap = -1
    End If
    g_log.WriteLog "Number of Joint Objects Formed = " & mModel.joints.count
End Function


Public Function FormFrmObj(Optional isSetJtObj As Boolean = True, _
                                Optional sectionFilter As Variant, Optional memberFilter As Variant, _
                                Optional isAllData As Boolean = False) As Integer
    'Forming frame object for connectivity analysis
    '0 = frame object formed.
    '-1 = no frame object formed.
    If Not mModel.frames.count = 0 Then GoTo ExitFunction
    Dim model As StrModel
    Dim dsFrm As oDataSheet
    Dim df As clsDataFrame, frm As StrFrame
    Dim jtI As StrJoint, jtJ As StrJoint
    Dim col_eleID As Long, col_section As Long 'column number in the dataframe
    Dim col_jtI As Long, col_jtJ As Long, col_eleLen As Long 'column number in the dataframe
    Dim col_MemID As Long, col_eleType As Long, col_localAxis As Long, col_orientationType As Long
    Dim col_subFrameName As Long
    Dim i As Long
    Dim ret As Integer
    
    Set model = mModel
    Set dsFrm = mDsFrm
    Set df = Me.DfFrm
    'g_log.WriteLog "Number of Frame Objects Data Readed to DataFrame = " & df.CountRows
    If df.isInitialized Then
        With dsFrm.tagSelector
            col_eleID = df.columnNum(.eleID)
            col_section = df.columnNum(.section)
            col_jtI = df.columnNum(.jtI)
            col_jtJ = df.columnNum(.jtJ)
            col_eleLen = df.columnNum(.eleLen)
            col_MemID = df.columnNum(.memID)
            col_eleType = df.columnNum(.EleType)
            col_localAxis = df.columnNum(.eleLocalAxis)
            col_orientationType = df.columnNum(.EleOrientationType)
            
        End With
        'Filter if needed
        If genFunc.isInitialised(sectionFilter) Then Set df = df.filter(col_section, sectionFilter)
        If genFunc.isInitialised(memberFilter) Then Set df = df.filter(col_MemID, memberFilter)
        
        For i = 1 To df.CountRows
            Set frm = New StrFrame
            
            
            If isSetJtObj Then
            'form jtI if jt I is not exist
                Set jtI = SetJoint(mModel, df.idata(i, col_jtI))
                Set jtJ = SetJoint(mModel, df.idata(i, col_jtJ))
                frm.Init df.idata(i, col_eleID), df.idata(i, col_section), jtI, jtJ, Length:=df.idata(i, col_eleLen)
                
            Else
                frm.Init df.idata(i, col_eleID), df.idata(i, col_section), Length:=df.idata(i, col_eleLen)
            End If
            
            If isAllData Then
                frm.localAxis = df.idata(i, col_localAxis)
                frm.OrientationTypeStr = df.idata(i, col_orientationType)

            End If
            'frm.df.idata(i, col_eleLen)
            ret = model.AddStrObjToColl(frm, obj_frm)
        Next i
    Else
        ret = -1
    End If
    
    'g_log.WriteLog "Number of Frame Objects Formed = " & mModel.frames.count
    
ExitFunction:
    FormFrmObj = ret
End Function

Public Function FormFrmForceObj(Optional sectionFilter As Variant, Optional memberFilter As Variant, _
                                Optional lcFilter As Variant) As Integer
    'Frame Object shall be formed in prior
    '0 = frame object formed.
    '-1 = no frame object formed.
    If Not mModel.frmForces.count = 0 Then GoTo ExitFunction
    
    Dim model As StrModel
    Dim dsFrmForce As oDataSheet
    Dim df As clsDataFrame, frm As StrFrame
    Dim frmForce As StrFrameForce
    Dim col_eleID As Long, col_station As Long 'column number in the dataframe
    Dim col_loadcomb As Long, col_stepType As Long 'column number in the dataframe
    Dim col_P As Long, col_V2 As Long, col_V3 As Long
    Dim col_T As Long, col_M2 As Long, col_M3 As Long
    Dim col_section As Long, col_MemID As Long
    Dim col_subFrameName As Long
    Dim i As Long
    Dim ret As Integer
    
    Set model = mModel
    Set dsFrmForce = mDsFrmForce
    
    Set df = Me.DfFrmForce
    g_log.WriteLog "Number of Frame Force Objects Data Readed to DataFrame = " & df.CountRows
      If df.isInitialized Then
        With dsFrmForce.tagSelector
          col_eleID = df.columnNum(.eleID)
          col_station = df.columnNum(.station)
          col_loadcomb = df.columnNum(.loadComb)
          col_stepType = df.columnNum(.stepType)
          col_P = df.columnNum(.p)
          col_V2 = df.columnNum(.V2)
          col_V3 = df.columnNum(.V3)
          col_T = df.columnNum(.t)
          col_M2 = df.columnNum(.M2)
          col_M3 = df.columnNum(.M3)
          col_section = df.columnNum(.section)
          col_MemID = df.columnNum(.memID)
          col_subFrameName = df.columnNum(.subEleName)
        End With
          'Filter if needed
          If genFunc.isInitialised(sectionFilter) Then Set df = df.filter(col_section, sectionFilter)
          If genFunc.isInitialised(memberFilter) Then Set df = df.filter(col_MemID, memberFilter)
          If genFunc.isInitialised(lcFilter) Then Set df = df.filter(col_loadcomb, lcFilter)
          
          With df
            For i = 1 To .CountRows
                Set frmForce = New StrFrameForce
                
                If frm Is Nothing Then
                    Set frm = SetFrame(model, CStr(.idata(i, col_eleID)))
                ElseIf Not frm.name = CStr(.idata(i, col_eleID)) Then
                    Set frm = SetFrame(model, CStr(.idata(i, col_eleID)))
                End If
                
                If frm Is Nothing Then
                    FormFrmForceObj = -1
                    Exit Function
                End If
                
                frmForce.Init2 frm, CDbl(.idata(i, col_station)), CStr(.idata(i, col_loadcomb)), _
                                    CStr(.idata(i, col_stepType)), _
                                    .idata(i, col_P), .idata(i, col_V2), .idata(i, col_V3), _
                                     .idata(i, col_T), .idata(i, col_M2), .idata(i, col_M3), CStr(.idata(i, col_subFrameName))
                frmForce.frame.AddFrameForceToColl frmForce
                ret = model.AddStrObjToColl(frmForce, obj_frmForce)
            Next i
          End With
      Else
          ret = -1
      End If
      g_log.WriteLog "Number of Frame Force Objects Formed = " & mModel.frmForces.count
ExitFunction:
    FormFrmForceObj = ret
End Function

Public Function FormMemberObj(Optional memberFilter As Variant) As Integer
    'return 0 if formed member obj.
    'return -1 if not formed/ problem encountered
    
    If Not mModel.members.count = 0 Then GoTo ExitFunction
    Dim model As StrModel
    Dim dsMember As oDataSheet
    Dim df As clsDataFrame, mem As StrMember, frm As StrFrame
    Dim col_MemID As Long, col_eleList As Long 'column number in the dataframe
    Dim i As Long, j As Long
    Dim eleNames() As String
    Dim isFormMember As Boolean
    Dim ret As Integer
    
    Set model = mModel
    Set dsMember = mDsMember

    Set df = Me.DfMember
    g_log.WriteLog "Number of Member Objects Data Readed to DataFrame = " & df.CountRows
    If df.isInitialized Then
        With dsMember.tagSelector
            col_MemID = df.columnNum(.memID)
            col_eleList = df.columnNum(.eleList)
        End With
        If genFunc.isInitialised(memberFilter) Then Set df = df.filter(col_MemID, memberFilter)
        With df
            For i = 1 To .CountRows
                eleNames = Split(.idata(i, col_eleList), ",")
                eleNames(0) = Replace(eleNames(0), "'", "", 1, 1)
                isFormMember = False
                For j = 0 To UBound(eleNames)
                    If model.IsStrObjectExist(eleNames(j), obj_frm) Then
                        isFormMember = True
                        Exit For
                    End If
                Next j
                
                If isFormMember Then
                    Set mem = New StrMember
                    For j = 0 To UBound(eleNames)
                        Set frm = SetFrame(model, eleNames(j))
                        mem.AddFrame frm, False
                        Set frm.member = mem
                    Next j
                    ret = model.AddStrObjToColl(mem, obj_member)
                End If
                
            Next i
        End With
    Else
        ret = -1
    End If
    g_log.WriteLog "Number of Member Objects Formed = " & mModel.members.count
ExitFunction:
    FormMemberObj = ret
End Function

Public Function FormConnTypeObjForMatching() As Integer
    'return 0 if formed joint obj.
    'return -1 if not formed
    Dim dsConnType As oDataSheet
    Dim df As clsDataFrame, connType As StrConnectionType
    Dim col_name As Long, col_mapJt As Long, col_mapMethod As Long, col_mapSections As Long
    Dim col_isRestraint As Long 'column number in the dataframe
    Dim i As Long
    Dim ret As Integer
    
    Set dsConnType = mDsConnType
    Set df = Me.DfConnType
    
    If df.isInitialized Then
        With dsConnType.tagSelector
            col_name = df.columnNum(.connID)
            col_mapJt = df.columnNum(.mapJt)
            col_mapMethod = df.columnNum(.mapMethod)
            col_mapSections = df.columnNum(.mapSections)
            col_isRestraint = df.columnNum(.isRestraint)
        End With
        
        For i = 1 To df.CountRows
            Set connType = New StrConnectionType
            connType.InitForMapping df.idata(i, col_name), df.idata(i, col_mapJt), df.idata(i, col_mapMethod), _
                                    df.idata(i, col_mapSections), df.idata(i, col_isRestraint)
            ret = mModel.AddStrObjToColl(connType, obj_connection)
            If Not ret = 0 Then
                g_log.RaiseWarning "Duplicate Connection type Name '" & df.idata(i, col_name) & "'.", DuplicateObjName
                MsgBox "Duplicate Connection type Name '" & df.idata(i, col_name) & "'. Macro will be terminated."
                GoTo ExitFunction
            End If
        Next i
    Else
        ret = -1
    End If
    g_log.WriteLog "Number of Connection Objects Formed = " & mModel.conns.count
ExitFunction:
    FormConnTypeObjForMatching = ret
End Function

Public Function FormConnTypeObjForForceExtract() As Integer
    'return 0 if formed joint obj.
    'return -1 if not formed
    Dim dsConnType As oDataSheet
    Dim df As clsDataFrame, connType As StrConnectionType
    Dim col_name As Long, col_extractSections As Long, col_isFullyPopulate As Long, col_endOffset As Long
    Dim col_matchedJoint As Long 'column number in the dataframe
    Dim i As Long
    Dim ret As Integer
    
    Set dsConnType = mDsConnType
    Set df = Me.DfConnType
    
    If df.isInitialized Then
        With dsConnType.tagSelector
            col_name = df.columnNum(.connID)
            col_extractSections = df.columnNum(.extractSections)
            col_isFullyPopulate = df.columnNum(.isFullyPopulate)
            col_endOffset = df.columnNum(.endOffset)
            col_matchedJoint = df.columnNum(.matchedJoint)
        End With
        
        Dim matchedJoint As String
        Dim jtNames() As String, j As Long
        Dim coll As Collection
        For i = 1 To df.CountRows
            Set connType = New StrConnectionType
            Set coll = New Collection
            matchedJoint = df.idata(i, col_matchedJoint)
            If Not matchedJoint = vbNullString Then
                jtNames = SplitAndTrimString(matchedJoint, ",")
                For j = LBound(jtNames) To UBound(jtNames)
                    coll.Add SetJoint(mModel, jtNames(j))
                Next j
            End If
            connType.InitForExtraction df.idata(i, col_name), df.idata(i, col_extractSections), df.idata(i, col_isFullyPopulate), _
                                    df.idata(i, col_endOffset), coll
            ret = mModel.AddStrObjToColl(connType, obj_connection)
            If Not ret = 0 Then
                g_log.RaiseWarning "Duplicate Connection type Name '" & df.idata(i, col_name) & "'.", DuplicateObjName
                MsgBox "Duplicate Connection type Name '" & df.idata(i, col_name) & "'. Macro will be terminated."
            End If
        Next i
    Else
        ret = -1
    End If
    g_log.WriteLog "Number of Connection Objects Formed = " & mModel.conns.count
ExitFunction:
    FormConnTypeObjForForceExtract = ret
End Function
Private Function FormSingleFrmObjByName(DfFrm As clsDataFrame, eleName As String) As StrFrame
    'May be used when forming member. if any frame in a member is missing due to applied section filter
    Dim dsFrm As oDataSheet
    Dim col_eleID As Long, col_section As Long 'column number in the dataframe
    Dim col_jtI As Long, col_jtJ As Long, col_eleLen As Long 'column number in the dataframe
    Dim col_MemID As Long
    
    Dim frm As StrFrame
    Dim jtI As StrJoint, jtJ As StrJoint
    Dim rRow As Long
    Dim eleID As String, section As String 'column number in the dataframe
    Dim jtIName As String, jtJName As String, eleLen As Double 'column number in the dataframe

    Set dsFrm = mDsFrm
    Dim df As clsDataFrame
    Set df = Me.DfFrm
    Dim ret As Integer
    
    With dsFrm.tagSelector
        col_eleID = df.columnNum(.eleID)
        col_section = df.columnNum(.section)
        col_jtI = df.columnNum(.jtI)
        col_jtJ = df.columnNum(.jtJ)
        col_eleLen = df.columnNum(.eleLen)
        col_MemID = df.columnNum(.memID)
    End With
        
    rRow = df.LookupRowNum(eleName, col_eleID)
    With df
        eleID = .idata(rRow, col_eleID)
        section = .idata(rRow, col_section)
        jtIName = .idata(rRow, col_jtI)
        jtJName = .idata(rRow, col_jtJ)
        eleLen = .idata(rRow, col_eleLen)
    End With
    Set frm = New StrFrame

    Set jtI = SetJoint(mModel, jtIName)
    Set jtJ = SetJoint(mModel, jtJName)
    frm.Init eleID, section, jtI, jtJ, Length:=eleLen

    ret = mModel.AddStrObjToColl(frm, obj_frm)

    Set FormSingleFrmObjByName = frm
End Function

Private Function SetJoint(model As StrModel, jtName As String) As StrJoint
    Dim jt As StrJoint
    Dim ret As Integer
    
    If model.IsStrObjectExist(jtName, obj_jt) Then
        Set jt = model.GetStrObject(jtName, obj_jt)
    Else
        Set jt = New StrJoint
        jt.Init jtName
        ret = model.AddStrObjToColl(jt, obj_jt)
    End If
    Set SetJoint = jt
End Function

Private Function SetFrame(model As StrModel, frmName As String) As StrFrame
    Dim frm As StrFrame
    
    If model.IsStrObjectExist(frmName, obj_frm) Then
        Set frm = model.GetStrObject(frmName, obj_frm)
    Else
        'this is expect to be happen rarely. Warning message if this happen
        Set frm = FormSingleFrmObjByName(mDfFrm, frmName)
    End If
    Set SetFrame = frm
End Function


Public Property Get DfJt() As clsDataFrame
    If mDfJt Is Nothing Then
        Set mDfJt = New clsDataFrame
        With mDsJt.tagSelector
            Set mDfJt = mDsJt.GetDataframe(.ID, .x, .y, .z, .isRestraint)
        End With
    End If
    Set DfJt = mDfJt
End Property

Public Property Get DfFrm() As clsDataFrame
    If mDfFrm Is Nothing Then
        'Set mDfFrm = New clsDataFrame
        With mDsFrm.tagSelector
            Set mDfFrm = mDsFrm.GetDataframe(.eleID, .section, .jtI, .jtJ, .eleLen, .memID, .EleType, .eleLocalAxis, .EleOrientationType)
        End With
    End If
    Set DfFrm = mDfFrm
End Property

Public Property Get DfFrmForce() As clsDataFrame
    If mDfFrmForce Is Nothing Then
        'Set mDfFrmForce = New clsDataFrame
        With mDsFrmForce.tagSelector
            Set mDfFrmForce = mDsFrmForce.GetDataframe(.eleID, .station, .loadComb, .stepType, .p, .V2, .V3, .t, .M2, .M3, .section, .memID, .subEleName)
        End With
    End If
    Set DfFrmForce = mDfFrmForce
End Property

Public Property Get DfMember() As clsDataFrame
    If mDfMember Is Nothing Then
        'Set mDfMember = New clsDataFrame
        With mDsMember.tagSelector
            Set mDfMember = mDsMember.GetDataframe(.memID, .eleList)
        End With
    End If
    Set DfMember = mDfMember
End Property

Public Property Get DfConnType() As clsDataFrame
    If mDfConnType Is Nothing Then
        'Set mDfMember = New clsDataFrame
        With mDsConnType.tagSelector
            Set mDfConnType = mDsConnType.GetDataframe(.connID, .mapJt, .mapMethod, .mapSections, .isRestraint, _
                                                    .extractSections, .isFullyPopulate, .endOffset, .matchedJoint)
        End With
    End If
    Set DfConnType = mDfConnType
End Property

Private Function SplitAndTrimString(str As String, andOperator As String) As String()
    Dim result() As String

    Dim i As Long

    ' Use the Split function to split the input string using the andOperator delimiter
    result = Split(str, andOperator)
    
    ' Trim and store the values in the result array
    For i = LBound(result) To UBound(result)
        result(i) = Trim(result(i))
    Next i
    result(0) = Replace(result(0), "'", "", 1, 1)
    ' Set the output as the result array
    SplitAndTrimString = result
End Function


