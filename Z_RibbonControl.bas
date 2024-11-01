Attribute VB_Name = "Z_RibbonControl"
'@Folder("Module")

'********************************************************
'Arthor: Lucas LEUNG
'Update Log
'Aug 2023 - Initial
'*******************************************************

Public g_log As New clsLog, isDetailMode As Boolean
Private ds_sys As New DataSheetSystem

Public Sub ProcessRibbon_DM(Control As IRibbonControl)

    'On Error GoTo Err:

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    isDetailMode = ds_sys.prop("ImportLog", "isDetailMode")
    
    Select Case Control.ID
        'call different macro based on button name pressed

        Case "btn_ImportDesignWS"
            ImportDesignWorkbook

        Case "btn_CreateSummaryEmpty": CreateNewSummary 1
        Case "btn_CreateSummaryGeneral": CreateNewSummary 2
        Case "btn_CreateSummarySteelMember": CreateNewSummary 3
        Case "btn_CreateSummarySteelConnection": CreateNewSummary 4
        Case "btn_CreateSummaryRC": CreateNewSummary 5
        'Import SAP
        
        Case "btn_ImportSAPDataAndDelete"
            ImportAndDelete SAPv17
        Case "btn_ImportSAPDataOnly"
            ImportOnly SAPv17

        'Import GSA
        Case "btn_ImportGSADataAndDelete"
            ImportAndDelete gsa
        Case "btn_ImportGSADataOnly"
            ImportOnly gsa

        'Case "btn_ShowMember"

        Case "btn_ExtractFrame_ExtremeCases"
            ExtractFrameForce_Correspondence

        Case "btn_ExtractFrame_byMemberPosition"
            ExtractFrameForce_WholeMember

        Case "btn_ExtractFrame_WholeMemberForGraph"
            ExtractFrameForce_WholeMemberForGraph

        Case "btn_ExtractConnAllNodeAndLc": ExtractConnForceAllNodesAndLC
        Case "btn_ExtractConnByCorrespondence": ExtractConnForceCorrespondence
        
        'Plot Chart
        Case "btn_PlotForceDiagrams"


        'Process Data
        Case "btn_ProcessData"
            ProcessData

        Case "btn_GetIdentifiedConnectionData": GetIdentifiedConnectionType
        Case "btn_MapConn": MatchConnection
        
        
        Case "btn_SummaryToDS"
            TransferDataFromSummary
        
        Case "btn_BoldMax": BoldMaxRow
        Case "btn_AddBotBorder": AddBottomBorderToSameGrp
        Case "btn_AddBotBorderAtPageBreak": AddBottomBorderAtPageBreak
        Case "btn_SetupFooter": SetupFooter
        Case "btn_ClrBotBorder": ClearBottomBorder
        Case "btn_ClrRightBorder": ClearRightBorder
        
        Case "btn_ViewTag_DataExtraction": ViewExtractionTag
        Case "btn_ViewTag_DesignWS": ViewDesignWorksheetTag
        Case "btn_ViewWorkbookStatus": ViewWorkbookStatus
            

        Case "btn_ClearWsData"
            ClearAllDataSheets

        Case "btn_ClearAllCharts"


        'Info
        Case "btn_viewLog"
            DisplayLog
        Case "btn_version"
            ShowVersion
    End Select


    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Exit Sub
Err:
    g_log.HandleError Err.Source, Err.Number, Err.Description
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
End Sub

Private Sub CreateNewSummary(index As Integer)
    Dim oper As New CreateNewSummary
    Select Case index
        Case 1: oper.CreateEmptySummary
        Case 2: oper.CreateForGeneralPurpose
        Case 3: oper.CreateForSteelMember
        Case 4: oper.CreateForSteelConnection
        Case 5: oper.CreateForRC
    End Select
End Sub


Private Sub SetupFooter()
    Dim oper As New FormattingTools
    oper.SetFooter
End Sub

Private Sub AddBottomBorderAtPageBreak()
    Dim oper As New FormattingTools
    oper.AddThickBottomLineAtPageBreak
End Sub

Private Sub BoldMaxRow()
    Dim oper As New FormattingTools
    oper.BoldMaxOfSameGroup
End Sub

Private Sub AddBottomBorderToSameGrp()
    Dim oper As New FormattingTools
    oper.AddBorder_Hori
End Sub

Private Sub ClearRightBorder()
    Dim oper As New FormattingTools
    oper.ClearBorder_Right
End Sub

Private Sub ClearBottomBorder()
    Dim oper As New FormattingTools
    oper.ClearBorder_Hori
End Sub

Private Sub ImportDesignWorkbook()
    
    Dim oper As ImportDesignWorksheet
    Set oper = New ImportDesignWorksheet
    'oper.Initialize g_log
    
    oper.Main
End Sub

Private Sub ImportAndDelete(dataformat As SourceDataFormat)
    g_log.CreateNewFile isDetailMode
    Dim oper As ImportData, ret As Integer
    Set oper = New ImportData
    oper.Initialize g_log
    ret = oper.Main(dataformat, True)
    
    If ret = 0 Then
        AnnounceComplete additionalStr:=oper.CompleteMessageAddition
    Else
        AnnounceTerminate defaultStr:=oper.TerminateMessageMain
    End If
    g_log.CloseFile
End Sub

Private Sub ImportOnly(dataformat As SourceDataFormat)
    g_log.CreateNewFile isDetailMode
    Dim oper As ImportData, ret As Integer
    Set oper = New ImportData
    oper.Initialize g_log
    ret = oper.Main(dataformat, False)
    
    If ret = 0 Then
        AnnounceComplete additionalStr:=oper.CompleteMessageAddition
    Else
        AnnounceTerminate defaultStr:=oper.TerminateMessageMain
    End If
    g_log.CloseFile
End Sub

Private Sub ExtractFrameForce_Correspondence()
    g_log.CreateNewFile isDetailMode
    Dim oper As ExtractFrmForce, ret As Integer
    Set oper = New ExtractFrmForce
    ret = oper.Main(ByExtremeCases)
    
    If ret = 0 Then
        AnnounceComplete additionalStr:=oper.CompleteMessageAddition
    Else
        AnnounceTerminate defaultStr:=oper.TerminateMessageMain
    End If
    g_log.CloseFile True
End Sub

Private Sub ExtractFrameForce_WholeMember()
    g_log.CreateNewFile isDetailMode
    Dim oper As ExtractFrmForce, ret As Integer
    Set oper = New ExtractFrmForce
    ret = oper.Main(ByMemberPos)
    
    If ret = 0 Then
        AnnounceComplete additionalStr:=oper.CompleteMessageAddition
    Else
        AnnounceTerminate defaultStr:=oper.TerminateMessageMain
    End If
    g_log.CloseFile True
End Sub

Private Sub ExtractFrameForce_WholeMemberForGraph()
    g_log.CreateNewFile isDetailMode
    Dim oper As ExtractFrmForce, ret As Integer
    Set oper = New ExtractFrmForce
    ret = oper.Main(ByMemberForGraph)
    
    If ret = 0 Then
        AnnounceComplete additionalStr:=oper.CompleteMessageAddition
    Else
        AnnounceTerminate defaultStr:=oper.TerminateMessageMain
    End If
    g_log.CloseFile True
End Sub


Public Sub ProcessData()
    g_log.CreateNewFile isDetailMode
    Dim oper As Object, ret As Integer

    Set oper = New OperProcessData
    ret = oper.Main
    If Not ret = 0 Then GoTo ExitSub

ExitSub:
    If ret = 0 Then
        AnnounceComplete additionalStr:=oper.CompleteMessageAddition
    Else
        AnnounceTerminate defaultStr:=oper.TerminateMessageMain
    End If
    g_log.CloseFile True
End Sub

Public Sub JointIdentification()
    'Set g_log = New clsLog
    'Set ds_sys = New DataSheetSystem
    g_log.CreateNewFile isDetailMode
    Dim oper As IdentifyJointInModel, ret As Integer
    Set oper = New IdentifyJointInModel
    ret = oper.Main
    
    If ret = 0 Then
        AnnounceComplete additionalStr:=oper.CompleteMessageAddition
    Else
        AnnounceTerminate defaultStr:=oper.TerminateMessageMain
    End If
    g_log.CloseFile True
    
End Sub

Private Sub GetIdentifiedConnectionType()
    'Set g_log = New clsLog
    'Set ds_sys = New DataSheetSystem
    g_log.CreateNewFile isDetailMode
    Dim oper As GetModelConnectionType, ret As Integer
    Set oper = New GetModelConnectionType
    ret = oper.Main
    
    If ret = 0 Then
        AnnounceComplete additionalStr:=oper.CompleteMessageAddition
    Else
        AnnounceTerminate defaultStr:=oper.TerminateMessageMain
    End If
    g_log.CloseFile True
    
End Sub

Private Sub MatchConnection()
    'Set g_log = New clsLog
    'Set ds_sys = New DataSheetSystem
    g_log.CreateNewFile isDetailMode
    Dim oper As MapConnection, ret As Integer
    Set oper = New MapConnection
    ret = oper.Main
    If ret = 0 Then
        AnnounceComplete additionalStr:=oper.CompleteMessageAddition
    Else
        AnnounceTerminate defaultStr:=oper.TerminateMessageMain
    End If
    g_log.CloseFile True
End Sub

Private Sub ExtractConnForceAllNodesAndLC()
    g_log.CreateNewFile isDetailMode
    Dim oper As ExtractConnForce, ret As Integer
    Set oper = New ExtractConnForce
    ret = oper.Main(EachNodeAllForces)
    If ret = 0 Then
        AnnounceComplete additionalStr:=oper.CompleteMessageAddition
    Else
        AnnounceTerminate defaultStr:=oper.TerminateMessageMain
    End If
    g_log.CloseFile
End Sub

Private Sub ExtractConnForceCorrespondence()
    g_log.CreateNewFile isDetailMode
    Dim oper As ExtractConnForce, ret As Integer
    Set oper = New ExtractConnForce
    ret = oper.Main(CorrespondeceCases)
    If ret = 0 Then
        AnnounceComplete additionalStr:=oper.CompleteMessageAddition
    Else
        AnnounceTerminate defaultStr:=oper.TerminateMessageMain
    End If
    g_log.CloseFile
End Sub

Private Sub TransferDataFromSummary()
    
    g_log.CreateNewFile isDetailMode
    Dim oper As SummaryToWS, ret As Integer
    Set oper = New SummaryToWS
    ret = oper.Main
    
    If ret = 0 Then
        AnnounceComplete additionalStr:=oper.CompleteMessageAddition
    Else
        AnnounceTerminate defaultStr:=oper.TerminateMessageMain
    End If
    g_log.CloseFile
End Sub

Private Sub ClearAllDataSheets()
    g_log.CreateNewFile isDetailMode
    Dim oper As DataSheetManager
    Set oper = New DataSheetManager
    oper.ClearAllData
    MsgBox "All imported data cleared."
    g_log.CloseFile
End Sub

Private Sub ViewWorkbookStatus()
    Dim oper As New ViewWorkbookStatus
    oper.ViewWorkbookStatus
End Sub

Private Sub ViewExtractionTag()
    Dim oper As ViewTag
    Set oper = New ViewTag
    oper.ViewExtractionTag
End Sub

Private Sub ViewDesignWorksheetTag()
    Dim oper As ViewTag
    Set oper = New ViewTag
    oper.ViewDesignWSTag
End Sub

Private Sub DisplayLog()
    g_log.DisplayLog
End Sub

Private Sub ShowVersion()
    Dim uf As New UFInfo
    UFInfo.Show
End Sub

Public Sub TerminateProcedure()
    g_log.CloseFile
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    End
End Sub

Private Sub AnnounceComplete(Optional defaultStr As String, Optional additionalStr As String)
    'By default. first row = "Operation Completed without warning." or "Operation Completed With xxx warning(s). Please view log file."
    Dim outputMsg As String
    If defaultStr = vbNullString Then
        If g_log.numOfWarning = 0 Then
            outputMsg = "Operation Completed without warning."
        Else
            outputMsg = "Operation Completed With " & g_log.numOfWarning & " warning(s). Please view log file."
        End If
    End If
    
    If additionalStr = vbNullString Then
        MsgBox outputMsg
    Else
        MsgBox outputMsg & vbCrLf & additionalStr
    End If
End Sub

Private Sub AnnounceTerminate(Optional defaultStr As String = "Operation Terminated.")
    'By default. first row = "Operation Completed without warning." or "Operation Completed With xxx warning(s). Please view log file."
    If defaultStr = vbNullString Then Exit Sub
    MsgBox defaultStr
End Sub

