VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OneWayPropertyBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("MVVM.Infrastructure.Binding")

'
'Option Explicit
'Implements IHandlePropertyChanged
'Private UI As MSForms.Control
'Private UIProperty As String
'Private Type TBinding
'    Source As Object
'    SourceProperty As String
'    ValueValiadator As IValueValidator
'End Type
'
'Private this As TBinding
'
'Public Sub Initialize(ByVal Control As MSForms.Control, ByVal ControlProperty As String, ByVal Source As Object, ByVal SourceProperty As String, _
'                    Optional InvertBoolean As Boolean = False, Optional ValueValiadator As IValueValidator)
'    Set UI = Control
'    UIProperty = ControlProperty
'    Set this.Source = Source
'    this.SourceProperty = SourceProperty
'    this.ValueValiadator = ValueValiadator
'    If TypeOf Source Is INotifyPropertyChanged Then RegisterPropertyChanged Source
'End Sub
'
'Private Sub RegisterPropertyChanged(ByVal Source As INotifyPropertyChanged)
'    Source.RegisterHandler Me
'End Sub
'
'Private Sub IHandlePropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal Name As String)
'    'update UI when source property change
'    If Source Is this.Source And Name = this.SourceProperty Then
'        If Not this.ValueValiadator Is Nothing Then
'            If this.ValueValiadator.IsValid() Then
'                VBA.Interaction.CallByName UI, UIProperty, VbLet, VBA.Interaction.CallByName(this.Source, this.SourceProperty, VbGet)
'            End If
'        Else
'            VBA.Interaction.CallByName UI, UIProperty, VbLet, VBA.Interaction.CallByName(this.Source, this.SourceProperty, VbGet)
'        End If
'    End If
'End Sub
'
''Private Sub UI_Change()
''    'update source property when UI change
''    VBA.Interaction.CallByName this.Source, this.SourceProperty, VbLet, VBA.Interaction.CallByName(UI, UIProperty, VbGet)
''End Sub
'
'
'
'
