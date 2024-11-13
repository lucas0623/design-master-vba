VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataSheetSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("DataSheet")


'********************************************************
'The sheet that contains all the design data, including extracted force and input data, as a connection between the force and design input data
'Arthor: Lucas LEUNG
'Update Log
'28 Aug 2023 - Initial Setup
'*******************************************************
Option Explicit
Private wsInteract As clsWorksheetsInteraction

Private pWsName As String, pWS As Worksheet
Private lRow As Long 'last row of data

'These are the default properties/tags that will be used in the system
Private iRow As Long, sRow As String 'the reference row representing the heading of the table
Private mStartRowWritten As Long, mEndRowWritten As Long
'Private Type TProp
'    text As String
'    loc As Long
'    desc As String
'End Type

Private mCollProps As New Collection


Private Sub Class_Initialize()
    
    Set wsInteract = New clsWorksheetsInteraction
    
    sRow = "rRow"
    CreateProp "nodeName", "Node Name"
    CreateProp "connType", "Connection Type"
    CreateProp "section", "Section Name"
    CreateProp "eleName", "Element Name"
    CreateProp "subEleName", "The Sub-element Name"
    CreateProp "eleLen", "Element Length"
    CreateProp "jtI", "Joint I"
    CreateProp "jtJ", "Joint J"
    CreateProp "memName", "Member Name"
    CreateProp "fFrm", "First Element of the Member"
    CreateProp "lFrm", "Last Element of the Member"
    CreateProp "pFrm", "Previous Element in the Member"
    CreateProp "nFrm", "Next Element in the Member"
    CreateProp "memTotalLen", "Total Length of Member"
    CreateProp "pos_fromMemJtI_percent", "Position in % from I-End of the Member"
    CreateProp "pos_fromMemJtJ_percent", "Position in % from J-End of the Member"
    CreateProp "pos_fromMemJtI", "Absolute Position from I-End of the Member"
    CreateProp "pos_fromMemJtJ", "Absolute Position from J-End of the Member"
    CreateProp "pos_fromEleJtI", "Absolute Position from I-End of the Element"
    CreateProp "pos_fromEleJtJ", "Absolute Position from J-End of the Element"
    CreateProp "pos_fromEleJtI_percent", "Position in % from I-End of the Element"
    CreateProp "pos_fromEleJtJ_percent", "Position in % from J-End of the Element"
    CreateProp "loadComb", "Load Combination"
    CreateProp "stepType", "Element Section Name"
    CreateProp "caseName", "Extreme Case"
    CreateProp "P", "Axial Force"
    CreateProp "V2", "Shear Force Along Minor Axis"
    CreateProp "V3", "Shear Force Along Major Axis"
    CreateProp "T", "Torsion"
    CreateProp "M2", "Moment About Minor Axis"
    CreateProp "M3", "Moment About Major Axis"
    CreateProp "targetWS", "The Design Worksheet to be used"
    
End Sub

Public Sub Initialize(wsName As String)
    On Error GoTo Err:
    pWsName = wsName
    Set pWS = wsInteract.setWorksheet(wsName) 'if fail to set the worksheet, error msg will be popped up.

    'if fail to set any of these variable, error msg will be popped up.
    iRow = wsInteract.getLocVar(pWS, sRow, False)

    Dim rCol As Long
    rCol = wsInteract.getLocVar(pWS, Me.section, True, isMustMatch:=False)
    If rCol = -1 Then rCol = wsInteract.getLocVar(pWS, Me.memName, True, isMustMatch:=False)
    If rCol = -1 Then rCol = wsInteract.getLocVar(pWS, Me.p, True, isMustMatch:=False)
    If rCol = -1 Then rCol = wsInteract.getLocVar(pWS, Me.targetWs, True, isMustMatch:=False)
    
    lRow = wsInteract.FindLastRow(iRow + 1, rCol, pWS)
    If lRow <= iRow Then lRow = iRow
    Exit Sub

Err:
    'Err occur if the worksheet/ any tag cannot be found
    g_log.HandleError Err.Source, CStr(Err.Number), Err.Description
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    End
End Sub

Public Function GetDataframe(ParamArray propName() As Variant) As clsDataFrame
    Dim df As clsDataFrame, heads() As String, i As Long
    Set df = New clsDataFrame
    ReDim heads(0 To UBound(propName))
    For i = 0 To UBound(propName)
        heads(i) = CStr(propName(i))
    Next i
    df.Init_ReadWorksheet pWS, rRow:=iRow, headTags:=heads
    Set GetDataframe = df
End Function

Public Sub WriteDataframe(df As clsDataFrame, ParamArray propName() As Variant)
    'the df shall be in the same order of the 'true' column
    Dim count As Long, i As Long, col As Long
    count = 1
    
    For i = LBound(propName) To UBound(propName)
        col = wsInteract.getLocVar(pWS, CStr(propName(i)), True, isMustMatch:=False)
        If Not col = -1 Then wsInteract.WriteArrToColumn df.iColumn(count), lRow + 1, col, pWS
        count = count + 1
    Next i
    mStartRowWritten = lRow + 1
    lRow = lRow + df.CountRows
    mEndRowWritten = lRow
End Sub

Public Function GetTagColumns(ws As Worksheet) As Long()
'Return the columns of all tags
    Dim Loc() As Long, i As Long
    ReDim Loc(0 To mCollProps.count - 1)
    For i = 1 To mCollProps.count
        mCollProps(i).Loc = wsInteract.getLocVar(ws, mCollProps(i).Text, True, isMustMatch:=False)
        Loc(i - 1) = mCollProps(i).Loc
    Next i
    GetTagColumns = Loc
End Function

Public Function GetTagsText(ParamArray propName() As Variant) As String()
'Return the string of all tags
    Dim str() As String, i As Long
    ReDim str(0 To mCollProps.count - 1)
    For i = 1 To mCollProps.count
        str(i - 1) = mCollProps(i).Text
    Next i
    GetTagsText = str
End Function

Public Function GetTagsDesc(ParamArray propName() As Variant) As String()
'Return the description of all tags
    Dim str() As String, i As Long
    ReDim str(0 To mCollProps.count - 1)
    For i = 1 To mCollProps.count
        str(i - 1) = mCollProps(i).Desc
    Next i
    GetTagsDesc = str
End Function

Private Sub CreateProp(Text As String, Desc As String)
    Dim prop As oProp
    Set prop = New oProp
    prop.Text = Text
    prop.Desc = Desc
    mCollProps.Add prop, Text
End Sub

Public Property Get Name() As String
    Name = pWS.Name
End Property

Public Property Get nodeName() As String
    nodeName = mCollProps("nodeName").Text
End Property

Public Property Get connType() As String
    connType = mCollProps("connType").Text
End Property

Public Property Get section() As String
    section = mCollProps("section").Text
End Property

Public Property Get eleName() As String
    eleName = mCollProps("eleName").Text
End Property

Public Property Get eleLen() As String
    eleLen = mCollProps("eleLen").Text
End Property

Public Property Get jtI() As String
    jtI = mCollProps("jtI").Text
End Property

Public Property Get jtJ() As String
    jtJ = mCollProps("jtJ").Text
End Property

Public Property Get memName() As String
    memName = mCollProps("memName").Text
End Property

Public Property Get fFrm() As String
    fFrm = mCollProps("fFrm").Text
End Property

Public Property Get lFrm() As String
    lFrm = mCollProps("lFrm").Text
End Property

Public Property Get pFrm() As String
    pFrm = mCollProps("pFrm").Text
End Property

Public Property Get nFrm() As String
    nFrm = mCollProps("nFrm").Text
End Property

Public Property Get memTotalLen() As String
    memTotalLen = mCollProps("memTotalLen").Text
End Property

Public Property Get pos_fromMemJtI_percent() As String
    pos_fromMemJtI_percent = mCollProps("pos_fromMemJtI_percent").Text
End Property

Public Property Get pos_fromMemJtJ_percent() As String
    pos_fromMemJtJ_percent = mCollProps("pos_fromMemJtJ_percent").Text
End Property

Public Property Get pos_fromMemJtI() As String
    pos_fromMemJtI = mCollProps("pos_fromMemJtI").Text
End Property

Public Property Get pos_fromMemJtJ() As String
    pos_fromMemJtJ = mCollProps("pos_fromMemJtJ").Text
End Property

Public Property Get pos_fromEleJtI() As String
    pos_fromEleJtI = mCollProps("pos_fromEleJtI").Text
End Property

Public Property Get pos_fromEleJtJ() As String
    pos_fromEleJtJ = mCollProps("pos_fromEleJtJ").Text
End Property

Public Property Get pos_fromEleJtI_percent() As String
    pos_fromEleJtI_percent = mCollProps("pos_fromEleJtI_percent").Text
End Property

Public Property Get pos_fromEleJtJ_percent() As String
    pos_fromEleJtJ_percent = mCollProps("pos_fromEleJtJ_percent").Text
End Property

Public Property Get loadComb() As String
    loadComb = mCollProps("loadcomb").Text
End Property

Public Property Get stepType() As String
    stepType = mCollProps("stepType").Text
End Property

Public Property Get caseName() As String
    caseName = mCollProps("caseName").Text
End Property

Public Property Get p() As String
    p = mCollProps("P").Text
End Property

Public Property Get V2() As String
    V2 = mCollProps("V2").Text
End Property

Public Property Get V3() As String
    V3 = mCollProps("V3").Text
End Property

Public Property Get t() As String
    t = mCollProps("T").Text
End Property

Public Property Get M2() As String
    M2 = mCollProps("M2").Text
End Property

Public Property Get M3() As String
    M3 = mCollProps("M3").Text
End Property

Public Property Get targetWs() As String
    targetWs = mCollProps("targetWS").Text
End Property

Public Property Get startRowWritten() As Long
    startRowWritten = mStartRowWritten
End Property

Public Property Get endRowWritten() As Long
    endRowWritten = mEndRowWritten
End Property

Public Property Get subEleName() As String
    subEleName = mCollProps("subEleName").Text
End Property

