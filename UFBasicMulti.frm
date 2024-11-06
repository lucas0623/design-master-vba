VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFBasicMulti 
   Caption         =   "Information"
   ClientHeight    =   1.96035e5
   ClientLeft      =   5430
   ClientTop       =   5475
   ClientWidth     =   1.96380e5
   OleObjectBlob   =   "UFBasicMulti.frx":0000
End
Attribute VB_Name = "UFBasicMulti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Folder("Userform.GeneralForm")

'********************************************************
'This is the Generalized Userform for the used in different modules

'Arthor: Lucas LEUNG
'Update Log
'Aug 2023 - Initial
'*******************************************************

Option Explicit

'For Hiding Default Tilte Bar and using own title bar
Private Const WM_NCLBUTTONDOWN = &HA1&
Private Const HTCAPTION = 2&
'Private Const WS_BORDER = &H800000
Private Declare PtrSafe Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal HWND As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLongPtr Lib "User32" Alias "SetWindowLongA" (ByVal HWND As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function DrawMenuBar Lib "User32" (ByVal HWND As Long) As Long
Private Declare PtrSafe Sub ReleaseCapture Lib "User32" ()
Private Declare PtrSafe Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal HWND As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private gHWND As Long
Private Const GWL_STYLE As Long = (-16)


'For
Private mmupMain As msForms.MultiPage, mfra() As msForms.frame
Private mbIsButtonAtCenter As Boolean
Private mCurrentPage As Integer
Private eventHandlerCollection As New Collection
Private pHoriSpace As Long, pVertSpace As Long
Private cTopPos() As Double, pBotMargin As Double
Private pCloseMode As Integer '0 when OK button is pressed, -1 when cancel button is pressed

'******************************************************************************************
'***************************For Customized Title Bar***************************************
'******************************************************************************************
Private Sub HandleDragMove(HWND As Long)
    Call ReleaseCapture
    Call SendMessage(HWND, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub LabelX_Click()
    pCloseMode = -1
    Me.Hide
End Sub

Private Sub MyTitleBar_Caption_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Button = 1 Then HandleDragMove gHWND
End Sub


'*****************************************************************************************************************
'***************************Button Hover Control/ OK Cancel Button Control****************************************
'*****************************************************************************************************************
Private Sub OKButtonInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'PURPOSE: Make OK Button appear Green when hovered on
  CancelButtonInactive.Visible = True
  OKButtonInactive.Visible = False
End Sub

Private Sub CancelButtonInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'PURPOSE: Make Cancel Button appear Green when hovered on
    CancelButtonInactive.Visible = False
    OKButtonInactive.Visible = True
End Sub


Private Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'PURPOSE: Reset Userform buttons to Inactive Status
    CancelButtonInactive.Visible = True
    OKButtonInactive.Visible = True
End Sub

Private Sub CancelButton_Click()
    pCloseMode = -1
    Me.Hide
End Sub

Private Sub OKButton_Click()
    pCloseMode = 0
    Me.Hide
End Sub

'*****************************************************************************************************************
'***************************Initialization & Close Behaviour******************************************************
'*****************************************************************************************************************
Private Sub UserForm_Initialize()

    'Remove default title bar and create own window
    Dim frm As Long
    Dim wHandle As Long
    wHandle = FindWindow(vbNullString, Me.caption)
    frm = GetWindowLong(wHandle, GWL_STYLE)
    'frm = frm And WS_BORDER

    SetWindowLongPtr wHandle, -16, 0
    DrawMenuBar wHandle
    gHWND = wHandle
    
    'Other Initializetion
    Me.height = 300
    pHoriSpace = 10
    pVertSpace = 10
    pBotMargin = 10
    
    'cTopPos(mCurrentPage) = 5
    
End Sub

Private Sub Userform_Activate()
    MyTitleBar_Caption.width = Me.width
    MyTitleBar_Border.width = Me.width
    
    'Position t
    Reposition_labelX
    If mbIsButtonAtCenter Then
        RepositionOkAndCancelButtonsToCenter
    Else
        RepositionOkAndCancelButtonsToRight
    End If
    Reposition_MultiPage
    RepositionUF
End Sub

Public Sub Initialize(numPages As Integer, PageNames() As String, Optional width As Double = 300, _
                        Optional isButtonAtCenter As Boolean = True)
    Dim i As Integer
    CreateMultiPageObject numPages, PageNames
    For i = 0 To numPages - 1
        cTopPos(i) = 5
    Next i
    Me.width = width
    mbIsButtonAtCenter = isButtonAtCenter
End Sub
Private Sub CreateMultiPageObject(numPages As Integer, PageNames() As String)
    Dim i As Integer
    
    Set mmupMain = Me.Controls.Add("Forms.MultiPage.1")
'    With mmupMain
'        .left = 0
'        .top = 24
'        .width = Me.width
'        .height = Me.height - pBotMargin - 24 - pVertSpace - OKButton.height
'    End With
    
    ReDim mfra(0 To numPages - 1)
    ReDim cTopPos(0 To numPages - 1)
    If numPages > 2 Then
        For i = 3 To numPages
            mmupMain.Pages.Add , "Page " & i
        Next i
    End If
    
    For i = 0 To numPages - 1
        mmupMain.Pages(i).caption = PageNames(i)
        mmupMain.value = i
        Set mfra(i) = mmupMain.Pages.item(i).Controls.Add("Forms.Frame.1")
        With mfra(i)
            .BackColor = RGB(255, 255, 255) ' White background color
'            .width = mmupMain.width
'            .height = mmupMain.height
'            .left = 0
'            .top = 0
            .BorderStyle = fmBorderStyleNone
        End With
        
    Next i
    mmupMain.value = 0
End Sub
'*************************************************************************************************************************
'*************************************Basic Methods***********************************************************************
'*************************************************************************************************************************
Private Sub Reposition_labelX()
    With LabelX
        .top = 2
        .left = Me.width - 20
    End With
End Sub

Public Sub RepositionOkAndCancelButtonsToCenter()
    With OKButton
        .top = Me.height - pBotMargin - .height
        .left = (Me.width / 2 - .width) / 2
    End With
    With OKButtonInactive
        .top = Me.height - pBotMargin - .height
        .left = (Me.width / 2 - .width) / 2
    End With
    With CancelButton
        .top = Me.height - pBotMargin - .height
        .left = Me.width / 2 + (Me.width / 2 - .width) / 2
    End With
    With CancelButtonInactive
        .top = Me.height - pBotMargin - .height
        .left = Me.width / 2 + (Me.width / 2 - .width) / 2
    End With
End Sub

Public Sub RepositionOkAndCancelButtonsToRight()
    With OKButton
        .top = Me.height - pBotMargin - .height
        .left = Me.width - pHoriSpace * 2 - .width * 2
    End With
    With OKButtonInactive
        .top = Me.height - pBotMargin - .height
        .left = Me.width - pHoriSpace * 2 - .width * 2
    End With
    With CancelButton
        .top = Me.height - pBotMargin - .height
        .left = Me.width - pHoriSpace - .width
    End With
    With CancelButtonInactive
        .top = Me.height - pBotMargin - .height
        .left = Me.width - pHoriSpace - .width
    End With
    
    'Debug.Print OKButton.top
End Sub
Private Sub Reposition_MultiPage()
    Dim i As Integer
    With mmupMain
        .left = 0
        .top = 24
        .width = Me.width - 5
        .height = Me.height - pBotMargin - 24 - pVertSpace - OKButton.height
    End With
    
    For i = 0 To UBound(mfra)
        With mfra(i)
            .width = mmupMain.width
            .height = mmupMain.height
            .left = 0
            .top = 0
        End With
    Next i
End Sub
Private Sub RepositionUF()
   Me.top = Application.top + (Application.UsableHeight / 2) - (Me.height / 2)
   Me.left = Application.left + (Application.UsableWidth / 2) - (Me.width / 2)
End Sub
Property Get CurrentPage() As Integer
    CurrentPage = mCurrentPage + 1
End Property

Property Let CurrentPage(num As Integer)
    mCurrentPage = num - 1
End Property

Property Get CloseState() As Integer
    CloseState = pCloseMode
End Property

Property Let TitleBarCaption(str As String)
    MyTitleBar_Caption.caption = str
End Property

Property Get OKButtonActive() As Object
    Set OKButtonActive = Me.OKButton
End Property
'*************************************************************************************************************************
'******************************Building UI Elements***********************************************************************
'*************************************************************************************************************************
'Public Function AddButton(ByRef reButton As MSForms.CommandButton, Optional caption As String = "Button", _
'                    Optional width As Double = 50, Optional height As Double = 15, _
'                    Optional top As Double = 10, Optional left As Double = 10, Optional pageNum As Integer = -1) As MSForms.CommandButton
'    Dim btn As MSForms.CommandButton
'
'    If pageNum = -1 Then
'        Set btn = Me.Controls.Add("Forms.CommandButton.1")
'        'btn.zposition = front
'    ElseIf pageNum = 0 Then 'ie: title bar of multipage userform
'        Set btn = mmupMain.Pages.Item(0).Controls.Add("Forms.CommandButton.1")
'    Else
'        Set btn = mfra(pageNum).Controls.Add("Forms.CommandButton.1")
'    End If
'
'    With btn
'        .height = height 'Application.Max(label.Height, 15)
'        .width = width
'        .top = top
'        .left = left
'        .caption = caption
'    End With
'
'    Set reButton = btn
'
'End Function
Public Function AddButton(Optional caption As String = "Button", _
                    Optional width As Double = 50, Optional height As Double = 15, _
                    Optional top As Double = 10, Optional left As Double = 10, Optional pageNum As Integer = -1) As msForms.CommandButton
    Dim btn As msForms.CommandButton

    If pageNum = -1 Then
        Set btn = Me.Controls.Add("Forms.CommandButton.1")
        'btn.zposition = front
    ElseIf pageNum = 0 Then 'ie: title bar of multipage userform
        Set btn = mmupMain.Pages.item(0).Controls.Add("Forms.CommandButton.1")
    Else
        Set btn = mfra(pageNum).Controls.Add("Forms.CommandButton.1")
    End If
    
    With btn
        .height = height 'Application.Max(label.Height, 15)
        .width = width
        .top = top
        .left = left
        .caption = caption
    End With
        
    Set AddButton = btn
    
End Function
Public Sub AddCheckBox(reCheckbox As msForms.checkbox, Optional title As String = "Please Select", _
                Optional Description As String, Optional isCheck As Boolean = False, Optional tipText As String, _
                Optional labelWidth As Double)
    Dim label As msForms.label
    Dim checkbox As msForms.checkbox
    
    If Not title = vbNullString Then

        Set label = mfra(mCurrentPage).Controls.Add("Forms.Label.1")
        With label
            .top = cTopPos(mCurrentPage)
            .left = pHoriSpace
            .caption = title
            If labelWidth = 0 Then
                .width = (Me.width - pHoriSpace) / 2 - pHoriSpace
            Else
                .width = labelWidth
            End If
            .AutoSize = True
            .ControlTipText = tipText
        End With
    End If
    
    Set checkbox = mfra(mCurrentPage).Controls.Add("Forms.Checkbox.1")
    With checkbox
        .height = 15 'Application.Max(label.Height, 15)
        .width = (Me.width - pHoriSpace) / 4
        .top = cTopPos(mCurrentPage)
        If labelWidth = 0 Then
            .left = (Me.width - pHoriSpace) / 2 + pHoriSpace
        Else
            .left = labelWidth + pHoriSpace
        End If
        .value = isCheck
        .caption = Description
        
    End With
        
    Set reCheckbox = checkbox
    cTopPos(mCurrentPage) = cTopPos(mCurrentPage) + checkbox.height + 10
    
End Sub

Sub AddCheckBox_double(reCheckbox As msForms.checkbox, reCheckbox2 As msForms.checkbox, _
                        Optional title As String = "Please Select", Optional description1 As String, _
                        Optional description2 As String, Optional isCheck1 As Boolean = False, Optional isCheck2 As Boolean = False, _
                        Optional tipText As String)
    
    Dim label As msForms.label, Label2 As msForms.label
    Dim checkbox As msForms.checkbox, CheckBox2 As msForms.checkbox
    
    AddCheckBox checkbox, title, description1, isCheck1, tipText
    
    cTopPos(mCurrentPage) = cTopPos(mCurrentPage) - checkbox.height - 10
    
    Set CheckBox2 = mfra(mCurrentPage).Controls.Add("Forms.Checkbox.1")
    With CheckBox2
        .height = 15 'Application.Max(label.Height, 15)
        .width = (Me.width - pHoriSpace) / 4
        .top = cTopPos(mCurrentPage)
        .left = 3 * (Me.width - pHoriSpace) / 4 '(Me.Width - pHoriSpace) / 2 + pHoriSpace
        .value = isCheck2
        .caption = description2
    End With
    

    Set reCheckbox = checkbox
    Set reCheckbox2 = CheckBox2
    
    cTopPos(mCurrentPage) = cTopPos(mCurrentPage) + checkbox.height + 10
    
End Sub

Public Sub AddComboBox_Empty(reComboBox As msForms.ComboBox, Optional title As String = "Please Select", Optional tipText As String)
    Dim label As msForms.label
    Dim ComboBox As msForms.ComboBox
    
    Set label = mfra(mCurrentPage).Controls.Add("Forms.Label.1")
    'Set label = Me.Controls.Add("Forms.Label.1")
    With label
        .top = cTopPos(mCurrentPage)
        '.top = 10
        .left = pHoriSpace
        .caption = title
        .width = (Me.width - pHoriSpace) / 2 - pHoriSpace
        .AutoSize = True
        .ControlTipText = tipText
        'AdjustLabelHeight label
        'Debug.Print "Label Height = " & .Height
    End With
    
    Set ComboBox = mfra(mCurrentPage).Controls.Add("Forms.ComboBox.1")
    With ComboBox
        .height = Application.Max(label.height, 18)
        .width = (Me.width - pHoriSpace) / 2 - pHoriSpace * 2
        .top = cTopPos(mCurrentPage)
        .left = (Me.width - pHoriSpace) / 2 + pHoriSpace
    End With
    
    If ComboBox.height > label.height Then
        label.top = label.top + (ComboBox.height - label.height) / 2
    End If
    
    Set reComboBox = ComboBox
    cTopPos(mCurrentPage) = cTopPos(mCurrentPage) + ComboBox.height + 10
End Sub

Sub AddComboBox(arr As Variant, reComboBox As msForms.ComboBox, Optional title As String = "Please Select", _
                Optional defaultVal As Variant = vbNullString, Optional tipText As String)
    AddComboBox_Empty reComboBox, title, tipText:=tipText
    reComboBox.List = arr
    
    If defaultVal = vbNullString Then
        reComboBox.value = arr(0)
    Else: reComboBox.value = defaultVal
    End If
   
End Sub

Sub AddInputBox(reTextBox As msForms.Textbox, Optional title As String = "Please Input", Optional def_value As Variant)
    Dim label As msForms.label
    Dim Textbox As msForms.Textbox
    
    Set label = mfra(mCurrentPage).Controls.Add("Forms.Label.1")
    With label
        .top = cTopPos(mCurrentPage)
        .left = pHoriSpace
        .caption = title
        .width = (Me.width - pHoriSpace) / 2 - pHoriSpace
        .AutoSize = True
    End With
    
    Set Textbox = mfra(mCurrentPage).Controls.Add("Forms.TextBox.1")
    With Textbox
        .height = Application.Max(label.height, 15)
        .width = (Me.width - pHoriSpace) / 2 - pHoriSpace * 2
        .top = cTopPos(mCurrentPage)
        .left = (Me.width - pHoriSpace) / 2 + pHoriSpace
        .value = def_value
    End With
    
    Set reTextBox = Textbox
    cTopPos(mCurrentPage) = cTopPos(mCurrentPage) + Textbox.height + 10
End Sub

'Sub AddRefEdit(reRefEdit As refEdit.refEdit, Optional Title As String = "Please Input", Optional def_value As Variant)
'
'    Dim label As MSForms.label
'    Dim refEdit As refEdit.refEdit
'
'    Set label = mfra(mCurrentPage).Controls.Add("Forms.Label.1")
'    With label
'        .top = cTopPos(mCurrentPage)
'        .left = pHoriSpace
'        .caption = Title
'        .width = (Me.width - pHoriSpace) / 2 - pHoriSpace
'        .AutoSize = True
'    End With
'
'    Set refEdit = mfra(mCurrentPage).Controls.Add("RefEdit.Ctrl")
'    With refEdit
'        .height = Application.Max(label.height, 30)
'        .width = (Me.width - pHoriSpace) / 2 - pHoriSpace * 2
'        .top = cTopPos(mCurrentPage)
'        .left = (Me.width - pHoriSpace) / 2 + pHoriSpace
'        .MultiLine = True
'        .value = def_value
'    End With
'
'    Set reRefEdit = refEdit
'    cTopPos(mCurrentPage) = cTopPos(mCurrentPage) + refEdit.height + 10
'End Sub

Public Sub AddSelectionBoxMulti_Empty(reListBox As msForms.listBox, Optional title As String = "listbox 1", _
                        Optional height_LB As Long = 140, Optional width_LB As Long, _
                        Optional width_cmdBtm As Long = 50, _
                        Optional is_reListBox2 As Boolean = False, Optional reListBox2 As msForms.listBox, _
                        Optional title2 As String = "listbox 2", _
                        Optional isCreateFrame As Boolean = False, Optional frameTitle As String = "", _
                        Optional reCmdBtn1 As msForms.CommandButton, Optional reCmdBtn2 As msForms.CommandButton)
        
    'add the items
    Dim newFrame As msForms.frame
    Dim ListBox1 As msForms.listBox, ListBox2 As msForms.listBox
    Dim cmdB1 As msForms.CommandButton, cmdB2 As msForms.CommandButton
    Dim label As msForms.label, Label2 As msForms.label
    
    Dim btnEvent As EventSelectionBoxMulti
    Dim framePosX As Double, framePosY As Double
    
    If width_LB = 0 Then
        width_LB = Me.width / 2 - pHoriSpace * 2.5 - width_cmdBtm / 2
    End If
    
    If isCreateFrame Then
        Set newFrame = mfra(mCurrentPage).Controls.Add("forms.frame.1", "TEST", True)
        framePosX = 10
        framePosY = 10
        
        With newFrame
            .caption = frameTitle
            .top = cTopPos(mCurrentPage)
            .left = pHoriSpace
            .width = 2 * width_LB + 4 * pHoriSpace + width_cmdBtm + framePosX
            .height = height_LB + framePosY + 15
            .Font.Bold = True
        End With
    End If
    
    Set label = mfra(mCurrentPage).Controls.Add("Forms.Label.1")
    With label
        .top = cTopPos(mCurrentPage) + framePosY
        .left = pHoriSpace + framePosX
        .caption = title
        .width = width_LB + 5
        .AutoSize = True
    End With
    
    Set ListBox1 = mfra(mCurrentPage).Controls.Add("Forms.ListBox.1")
    With ListBox1
        .height = height_LB
        .width = width_LB
        .top = cTopPos(mCurrentPage) + label.height + framePosY
        .left = pHoriSpace + framePosX
        .MultiSelect = fmMultiSelectExtended
    End With
    
    Set cmdB1 = mfra(mCurrentPage).Controls.Add("Forms.CommandButton.1")
    With cmdB1
        .caption = "->"
        .height = height_LB / 4
        .width = width_cmdBtm
        .top = ListBox1.top + height_LB / 2 - 10 - .height
        '.Top = cTopPos(mCurrentPage) + label.Height + height_LB / 2 - 10 - .Height
        .left = ListBox1.left + ListBox1.width + pHoriSpace
    End With
    Set reCmdBtn1 = cmdB1
    
    Set cmdB2 = mfra(mCurrentPage).Controls.Add("Forms.CommandButton.1")
    With cmdB2
        .caption = "<-"
        .height = height_LB / 4
        .width = width_cmdBtm
        .top = ListBox1.top + height_LB / 2 + 10
        .left = ListBox1.left + ListBox1.width + pHoriSpace
    End With
    Set reCmdBtn2 = cmdB2
    
    Set Label2 = mfra(mCurrentPage).Controls.Add("Forms.Label.1")
    With Label2
        .top = cTopPos(mCurrentPage) + framePosY
        .left = cmdB1.left + cmdB1.width + pHoriSpace + framePosX
        .caption = title2
        .width = width_LB
        .AutoSize = True
    End With
    
    Set ListBox2 = mfra(mCurrentPage).Controls.Add("Forms.ListBox.1")
    With ListBox2
        .height = height_LB
        .width = width_LB
        .top = cTopPos(mCurrentPage) + label.height + framePosY
        .left = cmdB1.left + cmdB1.width + pHoriSpace + framePosX
        .MultiSelect = fmMultiSelectExtended
    End With
    
    'Set btnEvent = New EventSelectionBoxMulti
    'btnEvent.Init Me, ListBox1, ListBox2, cmdB1, cmdB2
    'eventHandlerCollection.Add btnEvent
    'mColButtons.Add btnEvent
    
    Set reListBox = ListBox1
    If is_reListBox2 Then
        Set reListBox2 = ListBox2
    End If
    
    'Resize windows
    'Me.Width = Application.Max(Me.Width, width_cmdBtm + width_LB * 2 + pHoriSpace * 5 + framePosX * 3)
'    Me.height = cTopPos + def_botMargin + height_LB + label.height + 15
    
    cTopPos(mCurrentPage) = cTopPos(mCurrentPage) + height_LB + label.height + 10 + framePosY * 2
    
    If isCreateFrame Then
        newFrame.Visible = True
    End If
    
End Sub

Sub AddSelectionBoxMulti(arr As Variant, reListBox As msForms.listBox, Optional title As String = "listbox 1", _
                        Optional title2 As String = "listbox 2", Optional height_LB As Long = 140, _
                        Optional width_LB As Long, Optional width_cmdBtm As Long = 50, _
                        Optional is_reListBox2 As Boolean = False, Optional reListBox2 As msForms.listBox, _
                        Optional isCreateFrame As Boolean = False, Optional frameTitle As String = "")
        
    'add the items
    Dim newFrame As msForms.frame
    Dim ListBox1 As msForms.listBox, ListBox2 As msForms.listBox
    Dim cmdB1 As msForms.CommandButton, cmdB2 As msForms.CommandButton
    Dim label As msForms.label, Label2 As msForms.label
    
    Dim btnEvent As EventSelectionBoxMulti
    Dim framePosX As Double, framePosY As Double
    

    
    If isCreateFrame Then
        Set newFrame = mfra(mCurrentPage).Controls.Add("forms.frame.1", "TEST", True)
        framePosX = 10
        framePosY = 10
        
        With newFrame
            .caption = frameTitle
            .top = cTopPos(mCurrentPage)
            .left = pHoriSpace
            .width = 2 * width_LB + 4 * pHoriSpace + width_cmdBtm + framePosX
            .height = height_LB + framePosY + 15
            .Font.Bold = True
        End With
    End If
    
    Set label = mfra(mCurrentPage).Controls.Add("Forms.Label.1")
    With label
        .top = cTopPos(mCurrentPage) + framePosY
        .left = pHoriSpace + framePosX
        .caption = title
        .width = width_LB + 5
        .AutoSize = True
    End With
    
    Set ListBox1 = mfra(mCurrentPage).Controls.Add("Forms.ListBox.1")
    With ListBox1
        .height = height_LB
        .width = width_LB
        .top = cTopPos(mCurrentPage) + label.height + framePosY
        .left = pHoriSpace + framePosX
        .MultiSelect = fmMultiSelectExtended
        .List = arr
    End With
    
    Set cmdB1 = mfra(mCurrentPage).Controls.Add("Forms.CommandButton.1")
    With cmdB1
        .caption = "->"
        .height = height_LB / 4
        .width = width_cmdBtm
        .top = ListBox1.top + height_LB / 2 - 10 - .height
        '.Top = cTopPos + label.Height + height_LB / 2 - 10 - .Height
        .left = ListBox1.left + ListBox1.width + pHoriSpace
    End With
    
    Set cmdB2 = mfra(mCurrentPage).Controls.Add("Forms.CommandButton.1")
    With cmdB2
        .caption = "<-"
        .height = height_LB / 4
        .width = width_cmdBtm
        .top = ListBox1.top + height_LB / 2 + 10
        .left = ListBox1.left + ListBox1.width + pHoriSpace
    End With
    
    
    Set Label2 = mfra(mCurrentPage).Controls.Add("Forms.Label.1")
    With Label2
        .top = cTopPos(mCurrentPage) + framePosY
        .left = cmdB1.left + cmdB1.width + pHoriSpace + framePosX
        .caption = title2
        .width = width_LB
        .AutoSize = True
    End With
    
    Set ListBox2 = mfra(mCurrentPage).Controls.Add("Forms.ListBox.1")
    With ListBox2
        .height = height_LB
        .width = width_LB
        .top = cTopPos(mCurrentPage) + label.height + framePosY
        .left = cmdB1.left + cmdB1.width + pHoriSpace + framePosX
        .MultiSelect = fmMultiSelectExtended
    End With
    
'    Set btnEvent = New EventSelectionBoxMulti
'    btnEvent.Init Me, ListBox1, ListBox2, cmdB1, cmdB2
'    eventHandlerCollection.Add btnEvent
    'mColButtons.Add btnEvent
    
    Set reListBox = ListBox1
    If is_reListBox2 Then
        Set reListBox2 = ListBox2
    End If
    
    'Resize windows
    Me.width = Application.Max(Me.width, width_cmdBtm + width_LB * 2 + pHoriSpace * 5 + framePosX * 3)
'    Me.height = cTopPos + def_botMargin + height_LB + label.height + 15
    
    cTopPos(mCurrentPage) = cTopPos(mCurrentPage) + height_LB + label.height + 10 + framePosY * 2
    
    If isCreateFrame Then
        newFrame.Visible = True
    End If
    
End Sub

Sub AddComboBoxVisibilityControl(cb_master As msForms.ComboBox, _
                        cb_slave() As Object, val_visible As String)
    Dim btnEvent As EventComboBoxControl
    Set btnEvent = New EventComboBoxControl
    btnEvent.Init Me, cb_master, cb_slave, val_visible
    eventHandlerCollection.Add btnEvent
End Sub

Public Sub AdjustHeight()
    Dim maxTopPos As Double, i As Integer
    For i = 0 To UBound(cTopPos)
        If cTopPos(i) > maxTopPos Then maxTopPos = cTopPos(i)
    Next i
    Me.height = maxTopPos + pBotMargin + OKButton.height + pVertSpace * 2 + 30
End Sub

Public Sub AddEvent(customEvent As Object)
    eventHandlerCollection.Add customEvent
End Sub
'Private Sub AdjustLabelHeight(lb As MSForms.label)
'    Dim minHeight As Double
'    minHeight = 18
'    If lb.Height < minHeight Then
'        lb.AutoSize = False
'        lb.Height = minHeight
'    End If
'End Sub


