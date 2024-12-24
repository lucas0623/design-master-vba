VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VMForceExtraction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Userform.ForceExtraction"
Option Explicit
Implements INotifyPropertyChanged
Private Notification As New PropertyChangeNotification

Private genFunc As clsGeneralFunctions
Private ds_sys As DataSheetSystem
Private dsManager As DataSheetManager
Private console As New Debugger

'Path for the Properties

'Properties
Private msLcAll() As String, msMemberNamesAll() As String, msSectionsAll() As String 'for use in Userform
Private msLcSelect() As String, msMemberNamesSelect() As String, msSectionsSelect() As String 'for use in Userform
Private msLcExcl() As String, msMemberNamesExcl() As String, msSectionsExcl() As String
Private mGroupBy As String ', mbIsBySec As Boolean 'true = by section, false = by member
Private mGroupByNodeOrConnType As String
Private mPosFilterType As PositionFilterType, mdblPosFilterVal As Double, mPosFilterTypeStr As String
Private mbIsEnvMaxMin As Boolean, mbIsEachLC As Boolean, mbIsEachPos As Boolean
Private mbExtremeCase() As Boolean
Private mbMaxMinEnv() As Boolean
Private mWsSum As Worksheet, msWsList() As String
Private msConnTypeAll() As String, msConnTypeSelect() As String, msConnTypeExcl() As String
Private msNodeAll() As String, msNodeSelect() As String, msNodeExcl() As String

'Command
Private Commands As New Collection
'Private mSwitchUFValueCommand As New SwitchUFValueCommand

'Information to be obtained from the Userform
Private mbIsShownUF As Boolean, mbIsLoadLastSaveVerified As Boolean
Private mbIsAbleLoadLastSave As Boolean, msWarningMsg As String

'Default Value of some variables
Private mOperName As String
Private defaultIsBySec As Boolean

Private Sub OnPropertyChanged(ByVal Name As String)
    INotifyPropertyChanged_OnPropertyChanged Me, Name
End Sub
 
Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal Name As String)
    Notification.Notify Source, Name
End Sub
 
Private Sub INotifyPropertyChanged_RegisterHandler(ByVal handler As IHandlePropertyChanged)
    Notification.AddHandler handler
End Sub

Public Sub RegisterCommand(ByVal cmd As ICommand)
    Commands.Add cmd
End Sub

Public Function CountCommands() As Integer
    CountCommands = Commands.count
End Function
Public Sub Initialize(operName As String)
    mOperName = operName
    Dim ret As Integer
    ret = LoadLastSaveVal
    If Not ret = 0 Then LoadDefaultVal
End Sub
Private Sub Class_Initialize()
    Dim ds_frm As oDataSheet
    Dim ds_frmForce As oDataSheet
    Dim ds_member As oDataSheet
    Dim dsJt As oDataSheet, dsConnType As oDataSheet
    
    Set genFunc = New clsGeneralFunctions
    'Set pModel = New clsStrModel
    Set dsManager = New DataSheetManager
    Set ds_frm = dsManager.DSFrameData
    Set ds_frmForce = dsManager.DSFrameForce
    Set ds_member = dsManager.DSMemberData
    Set ds_sys = New DataSheetSystem
    Set dsJt = dsManager.DSJointCoor
    Set dsConnType = dsManager.DSConnectionType
    
    Dim df As clsDataFrame
    Set df = ds_frm.GetDataframe(ds_frm.tagSelector.section)
    msSectionsAll = genFunc.CStr_arr(genFunc.DeDupeOneDimArray(df.iColumn(1)))
    
    Set df = ds_member.GetDataframe(ds_member.tagSelector.memID)
    msMemberNamesAll = genFunc.CStr_arr(genFunc.ChangeArrBase(df.iColumn(1), 0))

    
    Set df = ds_frmForce.GetDataframe(ds_frmForce.tagSelector.loadComb)
    msLcAll = genFunc.CStr_arr(genFunc.DeDupeOneDimArray(df.iColumn(1)))
    
    Set df = dsJt.GetDataframe(dsJt.tagSelector.ID)
    msNodeAll = genFunc.CStr_arr(genFunc.ChangeArrBase(df.iColumn(1), 0))
    
    Set df = dsConnType.GetDataframe(dsConnType.tagSelector.connID)
    If df.isInitialized Then msConnTypeAll = genFunc.CStr_arr(genFunc.ChangeArrBase(df.iColumn(1), 0))
    
    ReDim mbExtremeCase(11)
    ReDim mbMaxMinEnv(5)
    msWsList = GetSummmarySheetsName
    'Set mWsSum = ActiveSheet
    
'    Set mSwitchUFValueCommand = New SwitchUFValueCommand
'    Set mSwitchUFValueCommand = mSwitchUFValueCommand.Create(Me)
    
    
End Sub

Public Sub LoadDefaultVal()
    g_log.WriteLogInDetailMode "Loading Default Value of the Userform..."
    Dim i As Integer, str() As String
    If mWsSum Is Nothing Then Set Me.wsSum = ActiveSheet
    
    Me.lc = msLcAll
    Me.MemberNames = genFunc.CStr_arr(genFunc.SortArray(msMemberNamesAll))
    Me.sections = genFunc.CStr_arr(genFunc.SortArray(msSectionsAll))
    Me.connTypes = genFunc.CStr_arr(genFunc.SortArray(msConnTypeAll))
    Me.Nodes = genFunc.CStr_arr(genFunc.SortArray(msNodeAll))
    'Erase Me.lcExcl
    'Erase Me.MemberNamesExcl
    'Erase Me.SectionsExcl
    Me.lcExcl = str
    Me.MemberNamesExcl = str
    Me.SectionsExcl = str
    Me.ConnTypesExcl = str
    Me.NodesExcl = str
    'set me.lcExcl=
    mPosFilterType = NoFilter
    mdblPosFilterVal = -1
    Me.IsBySec = True
    Me.IsByNode = False
    Me.isEnvMaxMin = True
    
    Me.IsEachLC = False
    Me.isEachPos = False
    Me.PosFilterType = NoFilter

    Me.ExtremeCase0 = True
    Me.ExtremeCase1 = True
    Me.ExtremeCase2 = True
    Me.ExtremeCase3 = True
    Me.ExtremeCase4 = True
    Me.ExtremeCase5 = True
    Me.ExtremeCase6 = True
    Me.ExtremeCase7 = True
    Me.ExtremeCase8 = True
    Me.ExtremeCase9 = True
    Me.ExtremeCase10 = True
    Me.ExtremeCase11 = True

    Me.MaxMinEnv0 = False
    Me.MaxMinEnv1 = True
    Me.MaxMinEnv2 = True
    Me.MaxMinEnv3 = True
    Me.MaxMinEnv4 = True
    Me.MaxMinEnv5 = True
    g_log.WriteLogInDetailMode "Default Value of the Loaded to the UserForm."
End Sub

Public Function LoadLastSaveVal()
        
    'Load Save Input if possible
    Dim cWS As Worksheet, propGrpName As String
    Dim ret As Integer
    
    If Me.wsSum Is Nothing Then
        Set cWS = ActiveSheet
    Else
        Set cWS = Me.wsSum
    End If
    
    propGrpName = cWS.Name & mOperName
    
    g_log.WriteLogInDetailMode "Loading Last Save for the Userform..."
    g_log.WriteLogInDetailMode "propGrpName = " & propGrpName
    
    '1.Check if there is existing save. If yes, try loading the data
    If Not ds_sys.isPropGrpExist(propGrpName) Then
        g_log.WriteLogInDetailMode "Cannot find the propGrpName in Data_System Sheet"
        ret = -1
        GoTo ExitFunc
    End If
    
    '2 Check if the saved data fit the default data (check ele, member, section and load comb)
    If Not mbIsLoadLastSaveVerified Then
        Dim isIdentical As Boolean
        
        
'        temp1 = ds_sys.propCondense(propGrpName, "Load Comb Selected")
'        temp2 = ds_sys.propCondense(propGrpName, "Load Comb Excluded")
'        temp = genFunc.CombineArrays(temp1, temp2)
'        temp = genFunc.SortArray(temp)
'        console.Log temp, "temp"
'        If genFunc.isInitialised(temp) Then temp = ggenFunc.CStr_arr(temp)
        isIdentical = VerifySavedData(propGrpName, "Load Comb Selected", "Load Comb Excluded", msLcAll)
        If Not isIdentical Then GoTo CannotLoad

        isIdentical = VerifySavedData(propGrpName, "Section Selected", "Section Excluded", msSectionsAll)
        If Not isIdentical Then GoTo CannotLoad

        isIdentical = VerifySavedData(propGrpName, "Member Selected", "Member Excluded", msMemberNamesAll)
        If Not isIdentical Then GoTo CannotLoad
        
        isIdentical = VerifySavedData(propGrpName, "Connection Type Selected", "Connection Type Excluded", msConnTypeAll)
        If Not isIdentical Then GoTo CannotLoad
        
        isIdentical = VerifySavedData(propGrpName, "Node Selected", "Node Excluded", msNodeAll)
        If Not isIdentical Then GoTo CannotLoad
      
        mbIsLoadLastSaveVerified = True
        mbIsAbleLoadLastSave = True
    ElseIf Not mbIsAbleLoadLastSave Then
        MsgBox msWarningMsg
        GoTo ExitFunc
    End If
    
    '3. Load Data
    Set Me.wsSum = Worksheets(ds_sys.prop(propGrpName, "wsName"))
    Me.lc = genFunc.CStr_arr(ds_sys.propCondense(propGrpName, "Load Comb Selected"))
    Me.lcExcl = genFunc.CStr_arr(ds_sys.propCondense(propGrpName, "Load Comb Excluded"))
    Me.MemberNames = genFunc.CStr_arr(ds_sys.propCondense(propGrpName, "Member Selected"))
    Me.MemberNamesExcl = genFunc.CStr_arr(ds_sys.propCondense(propGrpName, "Member Excluded"))
    Me.sections = genFunc.CStr_arr(ds_sys.propCondense(propGrpName, "Section Selected"))
    Me.SectionsExcl = genFunc.CStr_arr(ds_sys.propCondense(propGrpName, "Section Excluded"))
    Me.connTypes = genFunc.CStr_arr(ds_sys.propCondense(propGrpName, "Connection Type Selected"))
    Me.ConnTypesExcl = genFunc.CStr_arr(ds_sys.propCondense(propGrpName, "Connection Type Excluded"))
    Me.Nodes = genFunc.CStr_arr(ds_sys.propCondense(propGrpName, "Node Selected"))
    Me.NodesExcl = genFunc.CStr_arr(ds_sys.propCondense(propGrpName, "Node Excluded"))
    Me.PosFilterTypeStr = ds_sys.prop(propGrpName, "Position Filter Type")
    'Me.PosFilterTypeStr = ds_sys.prop(propGrpName, "Position Filter Value")
    'Me.IsBySec = CBool(ds_sys.prop(propGrpName, "isBySec"))
    Me.GroupBy = ds_sys.prop(propGrpName, "Group By")
    Me.GroupByNodeOrConnType = ds_sys.prop(propGrpName, "Group By Node Or Conn Type")
    Me.isEnvMaxMin = ds_sys.prop(propGrpName, "isEnvMaxMin")
    Me.ExtremeCase = genFunc.CBoolArr(ds_sys.prop(propGrpName, "Extreme Cases"))
    Me.IsEachLC = ds_sys.prop(propGrpName, "isEachLC")
    Me.isEachPos = ds_sys.prop(propGrpName, "isEachPos")
    Me.MaxMinEnv = genFunc.CBoolArr(ds_sys.prop(propGrpName, "Max Min Env"))
    g_log.WriteLogInDetailMode "Last Save Loaded."
    Exit Function
    
CannotLoad:
    g_log.WriteLogInDetailMode "Last saved data not identical to existing data."
    ret = -1

ExitFunc:
    LoadLastSaveVal = ret
    g_log.WriteLogInDetailMode "Last Save Cannot be Loaded."
End Function

Private Function VerifySavedData(propGrpName As String, prop1 As String, prop2 As String, compareData As Variant) As Boolean
    Dim ret As Boolean, i As Long
    
    Dim temp As Variant, temp1 As Variant, temp2 As Variant, count As Long
    temp1 = ds_sys.propCondense(propGrpName, prop1)
    temp2 = ds_sys.propCondense(propGrpName, prop2)
    temp = genFunc.CombineArrays(temp1, temp2)
    If genFunc.isInitialised(temp) Then
        temp = genFunc.CStr_arr(temp)
        temp = genFunc.SortArray(temp)
    Else
        If Not genFunc.isInitialised(compareData) Then
            ret = True
            GoTo ExitFunction
            
        Else
            GoTo ExitFunction
        End If
    End If
    
    
    If genFunc.isArraysIdentical(temp, compareData) Then
        ret = True
    Else
        msWarningMsg = "Failed to Load Last Saved Data!"
        If mbIsShownUF Then MsgBox msWarningMsg
    End If
'    For i = LBound(savedData) To UBound(savedData)
'        If Not genFunc.IsInArr(savedData(i), compareData) Then
'            'msWarningMsg = "Loading last saved data failed. Load Combination '" & tempLC(i) & "' cannot be found in the last saved data."
'            msWarningMsg = "Failed to Load Last Saved Data!"
'            If mbIsShownUF Then MsgBox msWarningMsg
'            ret = False
'            GoTo ExitFunction
'        End If
'    Next i
ExitFunction:
    VerifySavedData = ret
End Function
Public Sub SaveInput()
    'Save User Input to the
    Dim propGrpName As String
    
    propGrpName = mWsSum.Name & mOperName
    g_log.WriteLogInDetailMode "Saving User Input..."
    '1.Check if there is existing save. If yes, clear existing data. If no, create a new save
    If ds_sys.isPropGrpExist(propGrpName) Then
        'Clear Existing Data
        ds_sys.ClearPropGrpVal propGrpName
    Else
        'Create new save
        ds_sys.CreateCustomPropGrp propGrpName, "wsName", "Load Comb Selected", "Load Comb Excluded", _
                                                "Member Selected", "Member Excluded", _
                                                "Section Selected", "Section Excluded", _
                                                "Node Selected", "Node Excluded", _
                                                "Connection Type Selected", "Connection Type Excluded", _
                                                "isEachPos", "Position Filter Type", "Position Filter Value", _
                                                "Group By", "Group By Node Or Conn Type", _
                                                "isExtremeCase", "Extreme Cases", _
                                                "isEachLC", "isEnvMaxMin", "Max Min Env"
    End If

    '2. Write Data
    ds_sys.prop(propGrpName, "wsName") = mWsSum.Name
    'ds_sys.prop(propGrpName, "isBySec") = mbIsBySec
    ds_sys.propCondense(propGrpName, "Load Comb Selected") = msLcSelect
    ds_sys.propCondense(propGrpName, "Load Comb Excluded") = msLcExcl
    ds_sys.propCondense(propGrpName, "Member Selected") = msMemberNamesSelect
    ds_sys.propCondense(propGrpName, "Member Excluded") = msMemberNamesExcl
    ds_sys.propCondense(propGrpName, "Section Selected") = msSectionsSelect
    ds_sys.propCondense(propGrpName, "Section Excluded") = msSectionsExcl
    ds_sys.propCondense(propGrpName, "Node Selected") = msNodeSelect
    ds_sys.propCondense(propGrpName, "Node Excluded") = msNodeExcl
    ds_sys.propCondense(propGrpName, "Connection Type Selected") = msConnTypeSelect
    ds_sys.propCondense(propGrpName, "Connection Type Excluded") = msConnTypeExcl
    ds_sys.prop(propGrpName, "Position Filter Type") = PositionFilterTypeToString(mPosFilterType)
    ds_sys.prop(propGrpName, "Position Filter Value") = mdblPosFilterVal
    'ds_sys.prop(propGrpName, "isBySec") = mbIsBySec
    ds_sys.prop(propGrpName, "Group By") = mGroupBy
    ds_sys.prop(propGrpName, "Group By Node Or Conn Type") = mGroupByNodeOrConnType
    ds_sys.prop(propGrpName, "isEnvMaxMin") = mbIsEnvMaxMin
    ds_sys.prop(propGrpName, "Extreme Cases") = mbExtremeCase
    ds_sys.prop(propGrpName, "isEachLC") = mbIsEachLC
    ds_sys.prop(propGrpName, "isEachPos") = mbIsEachPos
    ds_sys.prop(propGrpName, "Max Min Env") = mbMaxMinEnv
End Sub

Private Function StringToPositionFilterType(ByVal filterTypeString As String) As PositionFilterType
    filterTypeString = UCase(filterTypeString)
    Select Case filterTypeString
        Case "NO FILTER"
            StringToPositionFilterType = NoFilter
        Case "END I"
            StringToPositionFilterType = EndI
        Case "END J"
            StringToPositionFilterType = EndJ
        Case "BOTH END"
            StringToPositionFilterType = BothEnd
        Case "SPECIFIED POSITION (ABSOLUTE)"
            StringToPositionFilterType = SpecifiedPosAbs
        Case Else
            StringToPositionFilterType = NoFilter ' Default value
    End Select
End Function
Private Function PositionFilterTypeToString(ByVal filterType As PositionFilterType) As String

    Select Case filterType
        Case NoFilter
            PositionFilterTypeToString = "NO FILTER"
        Case EndI
            PositionFilterTypeToString = "END I"
        Case EndJ
            PositionFilterTypeToString = "END J"
        Case BothEnd
            PositionFilterTypeToString = "BOTH END"
        Case SpecifiedPosAbs
            PositionFilterTypeToString = "SPECIFIED POSITION (ABSOLUTE)"
    End Select
End Function

Private Function GetSummmarySheetsName() As String()
    Dim ws_count As Long, arrL As Object
    Dim ws As Worksheet, i As Long
    Set arrL = CreateObject("System.Collections.ArrayList")
    
    'Set ws = ActiveSheet
    'With ActiveSheet
        'If InStr(1, ActiveSheet.name, "Sum") > 0 Then
        arrL.Add ActiveSheet.Name
        'End If
    'End With
    With ActiveWorkbook
        ws_count = .Worksheets.count
        'Set the activesheet as the first item
        
        For i = 1 To ws_count
            If InStr(1, .Worksheets(i).Name, "Sum") > 0 And Not .Worksheets(i).Name = ActiveSheet.Name Then
                arrL.Add .Worksheets(i).Name
            End If
        Next i
    End With
    
'    'Add last saved worksheets as well
'    Dim lastSaveSheet As String
'    lastSaveSheet = ds_sys.prop(propGrpName, "wsName")
'    If Not arrL.contains(lastSaveSheet) Then arrL.Add lastSaveSheet
    
    GetSummmarySheetsName = genFunc.CStr_arr(arrL.ToArray)
    
End Function


Public Property Get wsSum() As Worksheet
    Set wsSum = mWsSum
End Property

Public Property Set wsSum(ws As Worksheet)
    Set mWsSum = ws
    Me.wsSumName = ws.Name
    'OnPropertyChanged "wsSum"
End Property

Public Property Get wsSumName() As String
    wsSumName = mWsSum.Name
End Property

Public Property Let wsSumName(wsName As String)
    Set mWsSum = Worksheets(wsName)
    OnPropertyChanged "wsSumName"
End Property

Public Property Get wsList() As String()
    wsList = msWsList
End Property

Public Property Get GroupBy() As String
    GroupBy = mGroupBy
End Property

Public Property Let GroupBy(val As String)
    mGroupBy = val
    OnPropertyChanged "GroupBy"
End Property

Public Property Get IsBySec() As Boolean
    If UCase(Me.GroupBy) = "BY SECTION" Then
        IsBySec = True
    Else
        IsBySec = False
    End If
End Property

Public Property Let IsBySec(ByVal bSec As Boolean)
    'mbIsBySec = bSec
    If bSec Then
        Me.GroupBy = "By Section"
    Else
        Me.GroupBy = "By Member"
    End If
End Property

Public Property Get IsByNode() As Boolean
    If UCase(Me.GroupBy) = "BY NODE" Then
        IsByNode = True
    Else
        IsByNode = False
    End If
End Property

Public Property Let IsByNode(ByVal bNode As Boolean)
    'mbIsBySec = bSec
    If bNode Then
        Me.GroupByNodeOrConnType = "By Node"
    Else
        Me.GroupByNodeOrConnType = "By Connection Type"
    End If
End Property

Public Property Get GroupByNodeOrConnType() As String
    GroupByNodeOrConnType = mGroupByNodeOrConnType
End Property

Public Property Let GroupByNodeOrConnType(val As String)
    mGroupByNodeOrConnType = val
    OnPropertyChanged "GroupByNodeOrConnType"
End Property
Public Property Get lc() As String()
    lc = msLcSelect
End Property

Public Property Let lc(val() As String)
    msLcSelect = val
    OnPropertyChanged "lc"
End Property

Public Property Get lcExcl() As String()
    lcExcl = msLcExcl
End Property

Public Property Let lcExcl(val() As String)
    msLcExcl = val
    OnPropertyChanged "lcExcl"
End Property
Public Property Get MemberNames() As String()
    MemberNames = msMemberNamesSelect
End Property

Public Property Let MemberNames(memNames() As String)
    msMemberNamesSelect = memNames
    OnPropertyChanged "MemberNames"
End Property

Public Property Get MemberNamesExcl() As String()
    MemberNamesExcl = msMemberNamesExcl
End Property

Public Property Let MemberNamesExcl(memNames() As String)
    msMemberNamesExcl = memNames
    OnPropertyChanged "MemberNamesExcl"
End Property

Public Property Get sections() As String()
    sections = msSectionsSelect
End Property

Public Property Let sections(secArray() As String)
    msSectionsSelect = secArray
    OnPropertyChanged "Sections"
End Property

Public Property Get SectionsExcl() As String()
    SectionsExcl = msSectionsExcl
End Property

Public Property Let SectionsExcl(secArray() As String)
    msSectionsExcl = secArray
    OnPropertyChanged "SectionsExcl"
End Property

Public Property Get connTypes() As String()
    connTypes = msConnTypeSelect
End Property

Public Property Let connTypes(arr() As String)
    msConnTypeSelect = arr
    OnPropertyChanged "ConnTypes"
End Property

Public Property Get ConnTypesExcl() As String()
    ConnTypesExcl = msConnTypeExcl
End Property

Public Property Let ConnTypesExcl(arr() As String)
    msConnTypeExcl = arr
    OnPropertyChanged "ConnTypesExcl"
End Property
Public Property Get Nodes() As String()
    Nodes = msNodeSelect
End Property

Public Property Let Nodes(arr() As String)
    msNodeSelect = arr
    OnPropertyChanged "Nodes"
End Property

Public Property Get NodesExcl() As String()
    NodesExcl = msNodeExcl
End Property

Public Property Let NodesExcl(arr() As String)
    msNodeExcl = arr
    OnPropertyChanged "NodesExcl"
End Property

Public Property Get PosFilterTypeStr() As String
    PosFilterTypeStr = mPosFilterTypeStr
End Property

Public Property Let PosFilterTypeStr(str As String)
    mPosFilterTypeStr = str
    mPosFilterType = StringToPositionFilterType(str)
    OnPropertyChanged "PosFilterTypeStr"
End Property

Public Property Get PosFilterType() As PositionFilterType
    PosFilterType = mPosFilterType
End Property

Public Property Let PosFilterType(ByVal pft As PositionFilterType)
    mPosFilterType = pft
    mPosFilterTypeStr = PositionFilterTypeToString(pft)
End Property

Public Property Get PosFilterValue() As Double
    PosFilterValue = mdblPosFilterVal
End Property

Public Property Let PosFilterValue(ByVal dblVal As Double)
    mdblPosFilterVal = dblVal
End Property

Public Property Get IsEachLC() As Boolean
    IsEachLC = mbIsEachLC
End Property

Public Property Let IsEachLC(ByVal bEach As Boolean)
    mbIsEachLC = bEach
    OnPropertyChanged "IsEachLC"
End Property

Public Property Get isEachPos() As Boolean
    isEachPos = mbIsEachPos
End Property

Public Property Let isEachPos(ByVal bEachPos As Boolean)
    mbIsEachPos = bEachPos
    OnPropertyChanged "IsEachPos"
End Property

Public Property Get isEnvMaxMin() As Boolean
    isEnvMaxMin = mbIsEnvMaxMin
End Property

Public Property Let isEnvMaxMin(ByVal extCase As Boolean)
    mbIsEnvMaxMin = extCase
    OnPropertyChanged "IsEnvMaxMin"
End Property

Public Property Get ExtremeCase() As Boolean()
    ExtremeCase = mbExtremeCase
End Property

Public Property Let ExtremeCase(bExtCase() As Boolean)
    Me.ExtremeCase0 = bExtCase(0)
    Me.ExtremeCase1 = bExtCase(1)
    Me.ExtremeCase2 = bExtCase(2)
    Me.ExtremeCase3 = bExtCase(3)
    Me.ExtremeCase4 = bExtCase(4)
    Me.ExtremeCase5 = bExtCase(5)
    Me.ExtremeCase6 = bExtCase(6)
    Me.ExtremeCase7 = bExtCase(7)
    Me.ExtremeCase8 = bExtCase(8)
    Me.ExtremeCase9 = bExtCase(9)
    Me.ExtremeCase10 = bExtCase(10)
    Me.ExtremeCase11 = bExtCase(11)
End Property

Public Property Get ExtremeCase0() As Boolean
    ExtremeCase0 = mbExtremeCase(0)
End Property

Public Property Let ExtremeCase0(ByVal val As Boolean)
    mbExtremeCase(0) = val
    OnPropertyChanged "ExtremeCase0"
End Property

Public Property Get ExtremeCase1() As Boolean
    ExtremeCase1 = mbExtremeCase(1)
End Property

Public Property Let ExtremeCase1(ByVal val As Boolean)
    mbExtremeCase(1) = val
    OnPropertyChanged "ExtremeCase1"
End Property

Public Property Get ExtremeCase2() As Boolean
    ExtremeCase2 = mbExtremeCase(2)
End Property

Public Property Let ExtremeCase2(ByVal val As Boolean)
    mbExtremeCase(2) = val
    OnPropertyChanged "ExtremeCase2"
End Property

Public Property Get ExtremeCase3() As Boolean
    ExtremeCase3 = mbExtremeCase(3)
End Property

Public Property Let ExtremeCase3(ByVal val As Boolean)
    mbExtremeCase(3) = val
    OnPropertyChanged "ExtremeCase3"
End Property

Public Property Get ExtremeCase4() As Boolean
    ExtremeCase4 = mbExtremeCase(4)
End Property

Public Property Let ExtremeCase4(ByVal val As Boolean)
    mbExtremeCase(4) = val
    OnPropertyChanged "ExtremeCase4"
End Property

Public Property Get ExtremeCase5() As Boolean
    ExtremeCase5 = mbExtremeCase(5)
End Property

Public Property Let ExtremeCase5(ByVal val As Boolean)
    mbExtremeCase(5) = val
    OnPropertyChanged "ExtremeCase5"
End Property

Public Property Get ExtremeCase6() As Boolean
    ExtremeCase6 = mbExtremeCase(6)
End Property

Public Property Let ExtremeCase6(ByVal val As Boolean)
    mbExtremeCase(6) = val
    OnPropertyChanged "ExtremeCase6"
End Property

Public Property Get ExtremeCase7() As Boolean
    ExtremeCase7 = mbExtremeCase(7)
End Property

Public Property Let ExtremeCase7(ByVal val As Boolean)
    mbExtremeCase(7) = val
    OnPropertyChanged "ExtremeCase7"
End Property

Public Property Get ExtremeCase8() As Boolean
    ExtremeCase8 = mbExtremeCase(8)
End Property

Public Property Let ExtremeCase8(ByVal val As Boolean)
    mbExtremeCase(8) = val
    OnPropertyChanged "ExtremeCase8"
End Property

Public Property Get ExtremeCase9() As Boolean
    ExtremeCase9 = mbExtremeCase(9)
End Property

Public Property Let ExtremeCase9(ByVal val As Boolean)
    mbExtremeCase(9) = val
    OnPropertyChanged "ExtremeCase9"
End Property

Public Property Get ExtremeCase10() As Boolean
    ExtremeCase10 = mbExtremeCase(10)
End Property

Public Property Let ExtremeCase10(ByVal val As Boolean)
    mbExtremeCase(10) = val
    OnPropertyChanged "ExtremeCase10"
End Property

Public Property Get ExtremeCase11() As Boolean
    ExtremeCase11 = mbExtremeCase(11)
End Property

Public Property Let ExtremeCase11(ByVal val As Boolean)
    mbExtremeCase(11) = val
    OnPropertyChanged "ExtremeCase11"
End Property

Public Property Get MaxMinEnv() As Boolean()
    MaxMinEnv = mbMaxMinEnv
End Property

Public Property Let MaxMinEnv(bMaxMin() As Boolean)
    Me.MaxMinEnv0 = bMaxMin(0)
    Me.MaxMinEnv1 = bMaxMin(1)
    Me.MaxMinEnv2 = bMaxMin(2)
    Me.MaxMinEnv3 = bMaxMin(3)
    Me.MaxMinEnv4 = bMaxMin(4)
    Me.MaxMinEnv5 = bMaxMin(5)
End Property

Public Property Get MaxMinEnv0() As Boolean
    MaxMinEnv0 = mbMaxMinEnv(0)
End Property

Public Property Let MaxMinEnv0(ByVal val As Boolean)
    mbMaxMinEnv(0) = val
    OnPropertyChanged "MaxMinEnv0"
End Property

Public Property Get MaxMinEnv1() As Boolean
    MaxMinEnv1 = mbMaxMinEnv(1)
End Property

Public Property Let MaxMinEnv1(ByVal val As Boolean)
    mbMaxMinEnv(1) = val
    OnPropertyChanged "MaxMinEnv1"
End Property

Public Property Get MaxMinEnv2() As Boolean
    MaxMinEnv2 = mbMaxMinEnv(2)
End Property

Public Property Let MaxMinEnv2(ByVal val As Boolean)
    mbMaxMinEnv(2) = val
    OnPropertyChanged "MaxMinEnv2"
End Property

Public Property Get MaxMinEnv3() As Boolean
    MaxMinEnv3 = mbMaxMinEnv(3)
End Property

Public Property Let MaxMinEnv3(ByVal val As Boolean)
    mbMaxMinEnv(3) = val
    OnPropertyChanged "MaxMinEnv3"
End Property

Public Property Get MaxMinEnv4() As Boolean
    MaxMinEnv4 = mbMaxMinEnv(4)
End Property

Public Property Let MaxMinEnv4(ByVal val As Boolean)
    mbMaxMinEnv(4) = val
    OnPropertyChanged "MaxMinEnv4"
End Property

Public Property Get MaxMinEnv5() As Boolean
    MaxMinEnv5 = mbMaxMinEnv(5)
End Property

Public Property Let MaxMinEnv5(ByVal val As Boolean)
    mbMaxMinEnv(5) = val
    OnPropertyChanged "MaxMinEnv5"
End Property

'Public Property Get SwitchUFValueCommand()
'    Set SwitchUFValueCommand = mSwitchUFValueCommand
'End Property


