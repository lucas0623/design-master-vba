VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImageCommandBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder MVVM.Infrastructure.Binding

Option Explicit
Private WithEvents UI As msForms.Image
Attribute UI.VB_VarHelpID = -1

Private Type TBinding
    ViewModel As Object
    Command As ICommand
'    Context As Object
    CanExecuteContext As Object
    ExecuteContext As Object
End Type

Private this As TBinding

Public Sub Initialize(ByVal ViewModel As Object, ByVal Control As msForms.Image, ByVal Command As ICommand, _
                       Optional ByVal CanExecuteContext As Object, Optional ByVal ExecuteContext As Object)
    Set UI = Control
    Set this.Command = Command
'    Set this.Context = Context
    Set this.CanExecuteContext = CanExecuteContext
    Set this.ExecuteContext = ExecuteContext
End Sub
 
'Private Sub RegisterPropertyChanged(ByVal Source As INotifyPropertyChanged)
'    Source.RegisterHandler Me
'End Sub

Private Sub UI_Click()
    If this.Command.CanExecute(this.CanExecuteContext) Then
        this.Command.Execute this.ExecuteContext
    End If
End Sub



