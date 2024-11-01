VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUIManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Library"

'This class contains Sub and Funcitons for Creating Userforms/ other UIs, Getting User Input and Loading save input
'Arthor: Lucas LEUNG
'Update Log:
'12 Mar 2023: Minor Bug Fix
'06 Feb 2023: Updated getfile path function
'21 Nov 2022 - ver1.0

'***************************************************************
'Get User Input Through Userform
Function AddInputBox_rng(prompt As String, Optional title As String, Optional default As Range) As Range
    If default Is Nothing Then
        Set default = Range("D1")
    End If
    On Error GoTo ErrHandler:
        Set AddInputBox_rng = Application.inputBox(prompt:=prompt, title:=title, default:=default.address, Type:=8)
    On Error GoTo 0
    
    Exit Function
ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    End

End Function

Function GetFilePath(fileType As String, Optional myTitle As String = "File Path", Optional defaultPath As String) As String
    
    Dim path As String
    Dim checkPath As String
    On Error Resume Next
    checkPath = Dir(defaultPath, vbDirectory)
    On Error GoTo 0
    If checkPath = "" Then
        path = ActiveWorkbook.path & "\"
    Else
        path = defaultPath
    End If
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .InitialFileName = path
        .title = myTitle
        .AllowMultiSelect = False
        .filters.Clear
        .filters.Add fileType, "*" & fileType
        If .Show = -1 Then
            GetFilePath = .SelectedItems(1) '=-1 if action button is clicked, =0 if cancel button is clicked
        'Else:
            'End
        End If
    End With
End Function
Function GetListBoxArr(listBox As msForms.listBox) As String()
    Dim i As Long
    Dim tempStr() As String

    If listBox.ListCount > 0 Then
        ReDim tempStr(listBox.ListCount - 1)
    End If

    For i = 0 To listBox.ListCount - 1
      tempStr(i) = listBox.List(i)
    Next i

    GetListBoxArr = tempStr
End Function
'Function GetFilePath(fileType As String, Optional myTitle As String = "File Path", Optional defaultPath As String) As String
'
'    Dim strFileToOpen As String
'    If defaultPath = "" Then defaultPath = ThisWorkbook.path & "\"
'    'ChDrive defaultPath
'
'
'    On Error Resume Next
'    ChDir defaultPath
'    On Error GoTo 0
'    If checkPath = "" Then
'        path = ActiveWorkbook.path & "\"
'    Else
'        path = defaultPath
'    End If
'
'    strFileToOpen = Application.GetOpenFilename _
'    (title:=myTitle, _
'    FileFilter:="File Type *" & fileType & "* (*" & fileType & "*),")
'
'    If strFileToOpen = "False" Then
'        'MsgBox "No file selected.", vbExclamation, "Sorry!"
'        Application.ScreenUpdating = True
'        Application.Calculation = xlAutomatic
'        End
'    Else
'        GetFilePath = strFileToOpen
'    End If
'
'End Function
Public Function GetFolderPath(Optional defaultPath As String, Optional title As String) As String
    
    Dim path As String
    Dim checkPath As String
    On Error Resume Next
    checkPath = Dir(defaultPath, vbDirectory)
    On Error GoTo 0
    If checkPath = vbNullString Then
        path = ActiveWorkbook.path & "\"
    Else
        path = defaultPath
    End If
    
    If title = vbNullString Then title = "Please choose a folder"
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = path
        .title = title
        .AllowMultiSelect = False
    
        If .Show = -1 Then
            GetFolderPath = .SelectedItems(1) '=-1 if action button is clicked, =0 if cancel button is clicked
        Else:
            GetFolderPath = vbNullString
        End If
    End With

'    Dim strFileToOpen As String
'    strFileToOpen = Application.GetOpenFilename _
'    (Title:="Please choose the file path for excel workbooks containing SAP output")
'    'FileFilter:="Excel Workbooks *.xls* (*.xls*),")
'
'    If strFileToOpen = "False" Then
'        'MsgBox "No file selected.", vbExclamation, "Sorry!"
'        Exit Function
'    Else
'        GetFilePath = strFileToOpen
'    End If
    
End Function



