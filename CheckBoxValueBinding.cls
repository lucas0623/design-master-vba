VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CheckBoxValueBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder MVVM.Infrastructure.Binding

Option Explicit
Implements IHandlePropertyChanged
Private WithEvents UI As msForms.checkbox
Attribute UI.VB_VarHelpID = -1

Private Type TBinding
    Source As Object
    SourceProperty As String
End Type

Private this As TBinding

Public Sub Initialize(ByVal Control As msForms.checkbox, ByVal Source As Object, ByVal SourceProperty As String)
    Set UI = Control
    Set this.Source = Source
    this.SourceProperty = SourceProperty
    If TypeOf Source Is INotifyPropertyChanged Then RegisterPropertyChanged Source
End Sub
 
Private Sub RegisterPropertyChanged(ByVal Source As INotifyPropertyChanged)
    Source.RegisterHandler Me
End Sub
 
Private Sub IHandlePropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal Name As String)
    'update UI when source property change
    If Source Is this.Source And Name = this.SourceProperty Then
        UI.value = VBA.Interaction.CallByName(this.Source, this.SourceProperty, VbGet)
    End If
End Sub
 
Private Sub UI_Change()
    'update source property when UI change
    VBA.Interaction.CallByName this.Source, this.SourceProperty, VbLet, UI.value
End Sub

