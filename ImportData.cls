VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImportData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Operation.ImportData")

'********************************************************
'This module read GSA .csv file and write the data to the worksheet
'Arthor: Lucas LEUNG
'Update Log
'05 Jun 2023 - Initial Setup
'*******************************************************
Option Explicit
Private mCompleteMessageAddition As String
Private mTerminateMessageMain As String, mTerminateMessageAddition As String

Private mThisWorkbook As Workbook
Private DataConvertor As IDataFormatConvertor
Private mFilePath As String
Private mLog As clsLog
Private wsInteract As clsWorksheetsInteraction
Private genFunc As clsGeneralFunctions
Private UI As clsUIManager
Private DS As DataSheetManager
Private ds_jt As oDataSheet
Private ds_frm As oDataSheet
Private ds_frmForce As oDataSheet
Private ds_member As oDataSheet
Private ds_sys As DataSheetSystem

Private df_ele As clsDataFrame, df_force As clsDataFrame
Private df_joint As clsDataFrame

Private isTerminate As Boolean

Public Enum SourceDataFormat
    gsa = 1
    SAPv17 = 2
End Enum

Public Property Get CompleteMessageAddition() As String
    CompleteMessageAddition = mCompleteMessageAddition
End Property

Public Property Get TerminateMessageMain() As String
    TerminateMessageMain = mTerminateMessageMain
End Property

Public Property Get TerminateMessageAddition() As String
    TerminateMessageAddition = mTerminateMessageAddition
End Property

Private Sub Class_Initialize()
    Set mThisWorkbook = ActiveWorkbook
    Set wsInteract = New clsWorksheetsInteraction
    Set genFunc = New clsGeneralFunctions
    Set UI = New clsUIManager
    Set DS = New DataSheetManager
    Set ds_jt = DS.DSJointCoor
    Set ds_frm = DS.DSFrameData
    Set ds_frmForce = DS.DSFrameForce
    Set ds_member = DS.DSMemberData
    Set ds_sys = New DataSheetSystem
End Sub

Public Sub Initialize(Optional logFile As clsLog)
    Set mLog = logFile
End Sub

Public Function Main(dataformat As SourceDataFormat, Optional isDelete As Boolean = True, Optional isNoticeComplete As Boolean = True) As Integer
    Dim ret As Integer
    Set DataConvertor = SelectDataConvertor(dataformat)
    mFilePath = DataConvertor.GetUserInput
    ret = DataConvertor.ReadData(mFilePath)
    If ret = -1 Then GoTo TerminateSub
    
    mThisWorkbook.Activate
    Set df_ele = DataConvertor.DfEle
    Set df_force = DataConvertor.DfForce
    Set df_joint = DataConvertor.DfJoint
    
    If isDelete Then DS.ClearAllData False
    WriteData
    
    If Not mLog Is Nothing Then mLog.WriteLog "Data are Sucessfully Imported."
    If Not mLog Is Nothing Then mLog.WriteLog "Number of Joint Imported = " & df_joint.CountRows
    If Not mLog Is Nothing Then mLog.WriteLog "Number of Frame Imported = " & df_ele.CountRows
    If Not mLog Is Nothing Then mLog.WriteLog "Number of Frame Force Imported = " & df_force.CountRows
    mCompleteMessageAddition = "Number of Joint Imported = " & df_joint.CountRows & Chr(10) _
                            & "Number of Frame Imported = " & df_ele.CountRows & Chr(10) _
                            & "Number of Frame Force Imported = " & df_force.CountRows
'    If isNoticeComplete Then MsgBox "Data are Sucessfully Imported."
    
TerminateSub:
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Main = ret
End Function


Private Sub WriteData()
    Dim rRow As Long
    If Not df_ele Is Nothing Then
        With ds_frm.tagSelector
            'rRow = ds_frm.lastRow + 1
            ds_frm.WriteDataframe df_ele, False, True, .eleID, .section, .jtI, .jtJ, .eleLen, .EleType, .eleLocalAxis
            ds_frm.WriteColumn df_ele.iColumn(1), False, .memID
        End With
        If Not mLog Is Nothing Then mLog.WriteLog "'Element Table' is sucessfully read and saved to the worksheet."
    End If
    
    If Not df_force Is Nothing Then
        With ds_frmForce.tagSelector
            'rRow = ds_frmForce.lastRow + 1
            'If rRow = ds_frmForce.refRow + 2 Then rRow = rRow - 1 'as the first row of data will not be deleted for frame force sheet
            ds_frmForce.WriteDataframe df_force, False, True, .eleID, .station, .loadComb, .stepType, .p, .V2, .V3, .t, .M2, .M3, .subFrameID
            ds_frmForce.FillColumns .section, .memID
        End With
        If Not mLog Is Nothing Then mLog.WriteLog "'Element Force and Moment Table' is sucessfully read and saved to the worksheet."
    End If
    
    If Not df_joint Is Nothing Then
        With ds_jt.tagSelector
            'rRow = ds_jt.lastRow + 1
            ds_jt.WriteDataframe df_joint, False, False, .ID, .x, .y, .z, .isRestraint
            If Not mLog Is Nothing Then mLog.WriteLog "'Joint Definition' Table is sucessfully read and saved to the worksheet."
        End With
    End If
    
    If Not df_ele Is Nothing Then
        rRow = ds_member.lastRow + 1
        With ds_member.tagSelector
            ds_member.WriteColumn df_ele.iColumn(1), False, .memID
            ds_member.WriteColumn df_ele.iColumn(1), False, .fFrame
            ds_member.WriteColumn df_ele.iColumn(1), False, .lFrame
            ds_member.WriteColumn df_ele.iColumn(5), False, .totalLen
            ds_member.WriteColumn df_ele.iColumn(1), False, .eleList
        End With
        If Not mLog Is Nothing Then mLog.WriteLog "'Member Table' is sucessfully read and saved to the worksheet."
    End If
End Sub

Private Function SelectDataConvertor(dataformat As SourceDataFormat) As IDataFormatConvertor
    Dim convertor As IDataFormatConvertor
    If dataformat = gsa Then
        Set convertor = New DataConvertorGSA
    ElseIf dataformat = SAPv17 Then
        Set convertor = New DataConvertorSAP
    End If
    Set SelectDataConvertor = convertor
End Function


