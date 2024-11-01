VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AcceptCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Exposed
'@Folder Userform.Command
'@ModuleDescription "A command that closes (hides) a View."
'@PredeclaredId
Option Explicit
Implements ICommand
Private mView As IView

'@Description "Creates a new instance of this command."
Public Function Create(ByVal View As IView) As ICommand
    Dim result As AcceptCommand
    Set result = New AcceptCommand
    Set result.View = View
    Set Create = result
End Function

Public Property Get View() As IView
    Set View = mView
End Property

Public Property Set View(ByVal rhs As IView)
    Set mView = rhs
End Property

Private Function ICommand_CanExecute(ByVal Context As Object) As Boolean
    ICommand_CanExecute = True
End Function

Private Property Get ICommand_Description() As String
    ICommand_Description = "Accept changes and close."
End Property

Private Sub ICommand_Execute(ByVal Context As Object)
    MsgBox "ACCEPTED"
End Sub
