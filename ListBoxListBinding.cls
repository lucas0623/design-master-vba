VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListBoxListBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder MVVM.Infrastructure.Binding

Option Explicit
Implements IHandlePropertyChanged
Private WithEvents UI As msForms.listBox
Attribute UI.VB_VarHelpID = -1
Private genFuncs As clsGeneralFunctions

Private Type TBinding
    Source As Object
    SourceProperty As String
    ValueValiadator As IValueValidator
End Type

Private this As TBinding

Public Sub Initialize(ByVal Control As msForms.listBox, ByVal Source As Object, ByVal SourceProperty As String)
    Set UI = Control
    Set this.Source = Source
    this.SourceProperty = SourceProperty
    If TypeOf Source Is INotifyPropertyChanged Then RegisterPropertyChanged Source
    Set genFuncs = New clsGeneralFunctions
    Set this.ValueValiadator = New ValidateListBoxInput
End Sub
 
Private Sub RegisterPropertyChanged(ByVal Source As INotifyPropertyChanged)
    Source.RegisterHandler Me
End Sub
 
Private Sub IHandlePropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal name As String)
    'update UI when source property change
    If Source Is this.Source And name = this.SourceProperty Then
        If this.ValueValiadator.isValid(VBA.Interaction.CallByName(this.Source, this.SourceProperty, VbGet), this.Source, UI) Then
            UI.List = VBA.Interaction.CallByName(this.Source, this.SourceProperty, VbGet)
        Else
            UI.Clear
        'UI.List = VBA.Interaction.CallByName(this.Source, this.SourceProperty, VbGet)
        End If
        
    End If
End Sub


'Private Sub UI_Change()
'    'update source property when UI change
'    If Not UI.ListCount = UBound(VBA.Interaction.CallByName(this.Source, this.SourceProperty, VbGet)) + 1 Then
'        VBA.Interaction.CallByName this.Source, this.SourceProperty, VbLet, genFuncs.ListBoxToArray(UI)
'    End If
'End Sub


