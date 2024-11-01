VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SaveInputCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder Userform.Command
'@ModuleDescription "A command that switch the userform values between 'default' and 'last saved input'."
Option Explicit
Implements ICommand
Private mViewModel As Object

'@Description "Creates a new instance of this command."
Public Function Create(ByVal ViewModel As Object) As ICommand
    Dim result As SaveInputCommand
    Set result = New SaveInputCommand
    Set result.ViewModel = ViewModel

    Set Create = result
    result.ViewModel.RegisterCommand result
End Function

Public Property Get ViewModel() As Object
    Set ViewModel = mViewModel
End Property

Public Property Set ViewModel(ByVal rhs As Object)
    Set mViewModel = rhs
End Property

'Public Property Get Button() As Object
'    Set Button = mButton
'End Property
'
'Public Property Set Button(ByVal RHS As Object)
'    Set mButton = RHS
'End Property
Private Function ICommand_CanExecute(ByVal Context As Object) As Boolean
    'If mViewModel.isLoadDefault Then
    ICommand_CanExecute = True
    'End If
End Function

Private Property Get ICommand_Description() As String
    ICommand_Description = "SwitchUFValue."
End Property

Private Sub ICommand_Execute(ByVal Context As Object)
    mViewModel.SaveInput
End Sub

