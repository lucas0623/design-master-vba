VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtractFrmForce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Operation.ExtractFrameForce"

'********************************************************
'This module Control the Frame Force Extraction Behaviour
'Arthor: Lucas LEUNG
'Update Log

'12 Jun 2023 - Initial
'*******************************************************

Option Explicit
Private mCompleteMessageAddition As String
Private mTerminateMessageMain As String, mTerminateMessageAddition As String

Private genFunc As New clsGeneralFunctions
Private UFControl As New VMForceExtraction, uf As IView
Private ds_sys As New DataSheetSystem
Private ds_designData As New DataSheetSummary
Private mModel As New StrModel
Private coll_frmForces As Collection 'extracted frmForces
Private ws_sum As Worksheet

'Setting for Force Extraction. Obtained from Userform
Private ExtractFrmForceMethod As IExtractFrmForceMethod
Private lc() As String, MemberNames() As String, sections() As String 'for use in Userform

Public Enum FrameForceExtractMethod
    ByExtremeCases = 1
    ByMemberPos = 2
    ByMemberForGraph = 3
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

Public Function Main(method As FrameForceExtractMethod) As Integer
    Dim ret As Integer
    
    ret = CheckDataBeforeStart
    If Not ret = 0 Then
        mTerminateMessageMain = "no data is found in the workbook! The macro will be terminated."
        GoTo ExitFunc
    End If
    g_log.WriteLog "Force Extraction (Correspondence Cases) Started."
    
    '1. Get Data From Userform
    ret = ShowUFAndGetUserInput(method)
    If Not ret = 0 Then GoTo ExitFunc
    ds_designData.Initialize ws_sum.Name
    
    '2. Form StrModel Objects
    Application.StatusBar = "Start forming Structural Model Objects according to user preference...."
    g_log.WriteLogWithTime "Start forming Structural Model Objects according to user preference...."
    ret = FormStrObj 'StrObj is formed after the userform due to performance issue
    If Not ret = 0 Then GoTo ExitFunc
    g_log.WriteLogWithTime "Structural Model Objects are formed successfully.", True
    
    '3. Extract The forces according to user preference. Return as a collection of frame force objects (1 obj = 1 row result)
    g_log.WriteLogWithTime "Start extracting frame forces according to user perference from the Structrual Model Objects..."
    Set coll_frmForces = SelectExtractionMethod(method).ExtractForce
    g_log.WriteLogWithTime "Successfully formed the frame force collection. Total number of frame force object = " & coll_frmForces.count, True
    
    '4. Transform the frame force object to data frame object
    Application.StatusBar = "Transforming data for output...."
    g_log.WriteLog "Transforming data for output...."
    Dim df As clsDataFrame
    Set df = mModel.GetDataframe_fromColl(coll_frmForces, "frameSection", "frameName", "frameLength", _
                "frameJtIName", "frameJtJName", "memberName", "memberIFrameName", "memberJFrameName", _
                "preFrameName", "nextFrameName", "memberTotalLength", "pos_fromMemJtI_percent", "pos_fromMemJtJ_percent", _
                "pos_fromMemJtI", "pos_fromMemJtJ", "pos_fromEleJtI", _
                "pos_fromEleJtJ", "pos_fromEleJtI_percent", "pos_fromEleJtJ_percent", "loadcomb", "extremeCaseType", "stepType", _
                "P", "V2", "V3", "T", "M2", "M3", "subFrameName")
    
    '5. Write the df to the sheet
    g_log.WriteLogWithTime "Writing extracted data to the worksheet '" & ds_designData.Name & "'. Total number of data = " & df.CountRows
    With ds_designData
        .WriteDataframe df, .section, .eleName, .eleLen, .jtI, .jtJ, .memName, .fFrm, .lFrm, _
                .pFrm, .nFrm, .memTotalLen, .pos_fromMemJtI_percent, .pos_fromMemJtJ_percent, _
                .pos_fromMemJtI, .pos_fromMemJtJ, .pos_fromEleJtI, .pos_fromEleJtJ, _
                .pos_fromEleJtI_percent, .pos_fromEleJtJ_percent, .loadComb, .caseName, .stepType, .p, .V2, _
                .V3, .t, .M2, .M3, .subEleName
    End With
    g_log.WriteLogWithTime "All data written to the worksheet."
    Application.StatusBar = "Start forming Structural Model Objects according to user preference...."

    mCompleteMessageAddition = "Total number of data extracted = " & df.CountRows & "." & Chr(10) & _
                                 "Data is Written on Rows " & ds_designData.startRowWritten & " to " & ds_designData.endRowWritten
    Application.StatusBar = False
    Exit Function
ExitFunc:
    Main = -1
End Function

Private Function CheckDataBeforeStart() As Integer
    If Not ds_sys.prop("isWSImported", "ws_frame") Or Not ds_sys.prop("isWSImported", "ws_frameForce") Then
        CheckDataBeforeStart = -1
    End If
End Function

Private Function GetUserInput() As Integer
    Dim ret As Integer
    With UFControl
        Set ws_sum = .wsSum
        MemberNames = .MemberNames
        sections = .sections
    End With
End Function

Private Function FormStrObj() As Integer
    'Form Str Obj with sections, members and load comb filter applied
    Dim ret_formFrm As Integer, ret_formFrmForce As Integer, ret_formMem As Integer
    With mModel.Constructor
        ret_formFrm = .FormFrmObj
        ret_formFrmForce = .FormFrmForceObj
        ret_formMem = .FormMemberObj
        
    End With
End Function

Private Function SelectUF(method As FrameForceExtractMethod) As IView
    Dim uf As IView
    Select Case method
        Case ByExtremeCases
            Set uf = New UFExtractForceFrame
            uf.Initialize UFControl
            UFControl.Initialize "ExtractFrmForceMethodCo"
        
    End Select
    Set SelectUF = uf
End Function

Private Function ShowUFAndGetUserInput(method As FrameForceExtractMethod) As Integer
    Dim uf As IView, ret As Integer
    
    Set uf = SelectUF(method)
    uf.Show
    
    If uf.CloseState = 0 Then
        ret = GetUserInput
        g_log.WriteLogWithTime "The Userform is closed and selected preferences are saved.", True
    Else
        ret = -1
        g_log.WriteLogWithTime "The Userform is closed and the procedure will be terminated."
        GoTo ExitFunction
    End If
    
ExitFunction:
    ShowUFAndGetUserInput = ret
End Function

Private Function SelectExtractionMethod(method As FrameForceExtractMethod) As IExtractFrmForceMethod
    Select Case method
    Case ByExtremeCases
        Set ExtractFrmForceMethod = New ExtractFrmForceMethodCo
    End Select
    
    ExtractFrmForceMethod.Initialize mModel, UFControl
    Set SelectExtractionMethod = ExtractFrmForceMethod
End Function


