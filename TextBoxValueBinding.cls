VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextBoxValueBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder MVVM.Infrastructure.Binding

Option Explicit
Implements IHandlePropertyChanged
Private WithEvents UI As msForms.Textbox
Attribute UI.VB_VarHelpID = -1
 
Private Type TBinding
    Source As Object
    SourceProperty As String
End Type
 
Private this As TBinding
 
Public Sub Initialize(ByVal Control As msForms.Textbox, ByVal Source As Object, ByVal SourceProperty As String)
    Set UI = Control
    Set this.Source = Source
    this.SourceProperty = SourceProperty
    If TypeOf Source Is INotifyPropertyChanged Then RegisterPropertyChanges Source
End Sub
 
Private Sub RegisterPropertyChanges(ByVal Source As INotifyPropertyChanged)
    Source.RegisterHandler Me
End Sub
 
Private Sub IHandlePropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal name As String)
    If Source Is this.Source And name = this.SourceProperty Then
        UI.Text = VBA.Interaction.CallByName(this.Source, this.SourceProperty, VbGet)
    End If
End Sub
 
Private Sub UI_Change()
    VBA.Interaction.CallByName this.Source, this.SourceProperty, VbLet, UI.value
End Sub
