VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IdentifyFrameOrientationType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Operation.ConnAnalysis")
Private mCompleteMsgAddition As String
Private mTerminateMsgMain As String, mTerminateMsgAdditional As String

Private model As New StrModel
Private dsManager As New DataSheetManager
Private mDsSys As New DataSheetSystem
Private frmTypeClassifier As New FrameTypeClassifier

Public Property Get CompleteMessageAddition() As String
    CompleteMessageAddition = mCompleteMessageAddition
End Property

Public Property Get TerminateMessageMain() As String
    TerminateMessageMain = mTerminateMessageMain
End Property

Public Property Get TerminateMessageAddition() As String
    TerminateMessageAddition = mTerminateMessageAddition
End Property

Function Main() As Integer
    Dim ret As Integer
    
    ret = CheckDataBeforeStart
    If Not ret = 0 Then
        GoTo ExitFunction
    End If
    
    g_log.WriteLog "Classifing Elements Orientation Type."
    FormStrObj
    ClassifyFrameType
    WriteData
    mDsSys.prop("isCreated", "frameOrientationType") = True
    g_log.WriteLog "Classification of Frames Orientation Type Completed."
    g_log.WriteLog ""
    
ExitFunction:
    Main = ret
    
End Function


Private Function CheckDataBeforeStart() As Integer
    Application.Calculate
    If Not mDsSys.prop("isWSImported", "ws_frame") Or Not mDsSys.prop("isWSImported", "ws_joint") Then
        mTerminateMsgMain = "no data is found in the workbook! The macro will be terminated."
        CheckDataBeforeStart = -1
    End If
End Function

Private Sub FormStrObj()
    Dim ret As Integer
    ret = model.Constructor.FormJointObj
    ret = model.Constructor.FormFrmObj
End Sub

Private Sub ClassifyFrameType()

    Dim i As Long
    Dim frm As StrFrame, frms As Collection
    Dim frmType As EleOrientationType
    Set frms = model.frames
    For Each frm In frms
        frmType = frmTypeClassifier.Classify(frm)
        frm.orientationType = frmType
        g_log.WriteLog "Frame '" & frm.Name & "' is '" & frm.OrientationTypeStr & "' 1D element."
    Next
    
End Sub

Private Sub WriteData()
    Dim dsFrm As oDataSheet
    Set dsFrm = dsManager.DSFrameData
    Dim df As clsDataFrame
    
    Set df = model.GetDataframe(obj_frm, "OrientationTypeStr")
    g_log.WriteLog "Writing result to worksheets..."
    With dsFrm.tagSelector
        dsFrm.WriteDataframe df, True, False, .EleOrientationType
    End With
    
End Sub

