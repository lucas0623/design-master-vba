VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Library"

Private fso As Object, oFile As Object
Private pPath As String
Private p_iWarn() As Long '0 = total number of warning
Private mIsDetailMode As Boolean
Private mDataSheetSystem As New DataSheetSystem
Private UI As New clsUIManager

Public Enum WarningType
    
    'DataSheetFormatCannotRead = 1
    frmForceExtract_objNotFound = 1
    duplicateLabelInSummary = 2
    missingLabelInSummary = 3
    failToCreatePDF = 4
    DuplicateObjName = 5
    failToRegconizeConnMatchMethod = 6
    failToMatchAnySectionForConnForceExtract = 7
    CannotFindWSinWSData = 8
    EmptyWsNameInSummary = 9
    NoWorksheetFoundInWorkbook = 10
    missingJointCoor = 11
    unexpectedResult = 12
    duplicateDesignWorksheet = 13
'    objRepeatedname = 3
'    objNotFound = 4 'jt cannot be found for joint assignment data. New joint is created.
'    objAssignment_WrongDataType = 5 'Wrong Data Type. Record Skip.
'    jtRestraintInputWrong = 6 'Joint restraint input is not any one of the recgonized input data. Record skipped. User shall check the input.
'    failToConvertNameToID = 7
'
'    strFamObjNotFound = 8
'    strFamObjMasterTableNotFound = 9
'    strFamObjPropNameNotFoundInMasterTable = 10
''    frmRepeatedname = 6
''    frmNotFound = 7 'jt cannot be found for joint assignment data. New joint is created.
''    frmAssignment_WrongDataType = 8 'Wrong Data Type. Record Skip.
''    linkRepeatedname = 9
''    linkNotFound = 10 'Wrong Data Type. Record Skip.
''    linkAssignment_WrongDataType = 11 'Wrong Data Type. Record Skip.
''    loadPatRepeatedname = 12
End Enum

Public Enum ErrorType
    DataSheetFormatCannotRead = 1
End Enum
Public Sub Class_Initialize()

    Dim lib_fileTools As clsLibFileTools
    Set lib_fileTools = New clsLibFileTools
    Set fso = CreateObject("Scripting.FileSystemObject")
    'pPath = fso.GetParentFolderName(ActiveWorkbook.path) & "/Log.txt"
    On Error GoTo Err:
    pPath = lib_fileTools.GetLocalPath(ActiveWorkbook.path) & "\Log.txt"
    Set oFile = fso.CreateTextFile(pPath)
    Exit Sub
Err:
    'Cannot Find the path. Get user input of the input path
    pPath = mDataSheetSystem.prop("logFilePath", "path")
    If pPath = vbNullString Then
        MsgBox "Failed to save log file. Please select a folder for saving log file."
        pPath = UI.GetFolderPath(title:="Please choose a folder for saving log file")
        If pPath = vbNullString Then
            End
        End If
        pPath = pPath & "\Log.txt"
    End If
    
    mDataSheetSystem.prop("logFilePath", "path") = pPath
End Sub

Property Let path(str As String)
    pPath = str
End Property

Public Sub CreateNewFile(isDetailMode As Boolean)
    
    'Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFile = fso.CreateTextFile(pPath)
    ReDim p_iWarn(12)
    oFile.WriteLine "Start at " & Format(Date, "dd MMM yyyy") & "-" & Time() & vbCrLf
    mIsDetailMode = isDetailMode
End Sub

Public Sub WriteLogInDetailMode(sText As String)
    If mIsDetailMode Then oFile.WriteLine sText
End Sub

Public Sub WriteLog(sText As String)
    oFile.WriteLine sText
End Sub

Public Sub WriteLogWithTime(sText As String, Optional isAddBlankLine As Boolean = False)
    oFile.WriteLine sText & " (Timestamp - " & Time() & ")"
    If isAddBlankLine Then oFile.WriteLine ""
End Sub

Public Sub RaiseWarning(sText As String, warningname As WarningType)
    oFile.WriteLine "WARNING(" & warningname & "): " & sText
    p_iWarn(0) = p_iWarn(0) + 1
    p_iWarn(warningname) = p_iWarn(warningname) + 1

End Sub

'Public Sub RaiseError(sText As String, errName As ErrorType)
'    oFile.WriteLine "Error(" & errName & "): " & sText
'    MsgBox
'End Sub

Public Sub DisplayLog()
    
    Dim tempData As Variant
    
    tempData = Shell("notepad.exe " & pPath, vbNormalFocus)

End Sub
Public Sub CloseFile(Optional isShowMsgBox As Boolean = False)
    
    oFile.WriteLine vbCrLf & "Completed at " & Format(Date, "dd MMM yyyy") & "-" & Time() & vbCrLf
    If Not p_iWarn(0) = 0 Then oFile.WriteLine "WARNING: There are " & CStr(p_iWarn(0)) & " number of warnings recorded. Please Check!"
    oFile.Close
    Set oFile = Nothing
    
'    If isShowMsgBox Then
'        If p_iWarn(0) = 0 Then
'            MsgBox "Completed."
'        Else
'            MsgBox "Completed with " & CStr(p_iWarn(0)) & " number of warnings recorded. Please Check!"
'        End If
'    End If
End Sub
Public Sub HandleError(ByVal sRoutineName As String, _
                        ByVal sErrorNo As String, _
                        ByVal sErrorDescription As String)
    Dim sMessage As String
    sMessage = "Err " & sErrorNo & " - " & sErrorDescription
    Call MsgBox(sMessage, vbCritical, sRoutineName & " - Error")
    'Resume Next
    Call LogFile_WriteError(sRoutineName, sMessage)
    
   'End
End Sub

Private Function LogFile_WriteError(ByVal sRoutineName As String, _
                                   ByVal sMessage As String)
    Dim sText As String
        
   On Error GoTo ErrorHandler

'   path = "C:\Users\lucasleung\Desktop\TEST GSA\Debug Log.txt"
'   If (fso Is Nothing) Then
'      Set fso = New FileSystemObject
'   End If
'   If (oFile Is Nothing) Then
'      If (fso.FileExists(path) = False) Then
'         Set oFile = fso.OpenTextFile(path, IOMode.ForWriting, True)
'      Else
'         Set oFile = fso.OpenTextFile(path, IOMode.ForAppending)
'      End If
'   End If
   sText = sText & vbNullString & vbCrLf
   sText = sText & Format(Date, "dd MMM yyyy") & "-" & Time() & vbCrLf
   sText = sText & " " & sRoutineName & vbCrLf
   sText = sText & " " & sMessage & vbCrLf
   oFile.WriteLine sText
   oFile.Close
   Set oFile = Nothing
   Exit Function
ErrorHandler:
   Set oFile = Nothing
   Call MsgBox("Unable to write to log file", vbCritical, "LogFile_WriteError")
End Function

Public Property Get numOfWarning(Optional warningname As WarningType = 0) As Long
    numOfWarning = p_iWarn(warningname)
End Property




