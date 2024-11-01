VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BindingsManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.Infrastructure.Binding"
Option Explicit
Private BindingColl As New Collection

'@Description "Binds a MSForms.Control property to a source property"
'Public Sub BindOneWayProperty(ByVal Control As MSForms.Control, ByVal ControlProperty As String, ByVal SourceProperty As String, ByVal Source As Object, Optional ByVal InvertBoolean As Boolean = False)
'
'    Dim Binding As OneWayPropertyBinding
'    Set Binding = New OneWayPropertyBinding
'
'    Binding.Initialize Control, ControlProperty, Source, SourceProperty, InvertBoolean
'
'    BindingColl.Add Binding
'
'End Sub

'@Description "Binds a MSForms.Control property to a source property"
Public Sub BindOneTimeProperty(ByVal Control As msForms.Control, ByVal ControlProperty As String, ByVal SourceProperty As String, ByVal Source As Object)
     
    Dim Binding As OneTimePropertyBinding
    Set Binding = New OneTimePropertyBinding
     
    Binding.Initialize Control, ControlProperty, Source, SourceProperty
     
    BindingColl.Add Binding
 
End Sub
''@Description "Binds a MSForms.Control property to a source property"
'Public Sub BindTwoWayProperty(ByVal Control As MSForms.Control, ByVal ControlProperty As String, ByVal SourceProperty As String, ByVal Source As Object, Optional ByVal InvertBoolean As Boolean = False)
'
'    Dim Binding As TwoWayPropertyBinding
'    Set Binding = New TwoWayPropertyBinding
'
'    Binding.Initialize Control, ControlProperty, Source, SourceProperty, InvertBoolean
'
'    'Set BindTwoWayProperty = Binding
'    BindingColl.Add Binding
'End Sub

'@Description "Binds the Text/Value of a MSForms.TextBox to a source property"
Public Sub BindTextBox(ByVal Control As msForms.Textbox, ByVal SourceProperty As String, ByVal Source As Object)
     
    Dim Binding As TextBoxValueBinding
    Set Binding = New TextBoxValueBinding
     
    Binding.Initialize Control, Source, SourceProperty
     
    'Set BindTextBox = Binding
     BindingColl.Add Binding
End Sub
 
'@Description "Binds the Text of a MSForms.ComboBox to a String source property"
Public Sub BindComboBox(ByVal Control As msForms.ComboBox, ByVal SourceProperty As String, ByVal Source As Object)
     
    Dim Binding As ComboBoxValueBinding
    Set Binding = New ComboBoxValueBinding
     
    Binding.Initialize Control, Source, SourceProperty
     
    'Set BindComboBox = Binding
    BindingColl.Add Binding
End Sub
 
'@Description "Binds the Value of a MSForms.CheckBox to a Boolean source property"
Public Sub BindCheckBox(ByVal Control As msForms.checkbox, ByVal SourceProperty As String, ByVal Source As Object)

    Dim Binding As CheckBoxValueBinding
    Set Binding = New CheckBoxValueBinding

    Binding.Initialize Control, Source, SourceProperty

    'Set BindCheckBox = Binding
    BindingColl.Add Binding
End Sub

'@Description "Binds Command"
Public Sub BindCommand(ByVal ViewModel As Object, ByVal Control As msForms.CommandButton, ByVal Command As ICommand, _
                            Optional ByVal CanExecuteContext As Object, Optional ByVal ExecuteContext As Object)

    Dim Binding As CommandBinding
    Set Binding = New CommandBinding

    Binding.Initialize ViewModel, Control, Command, CanExecuteContext, ExecuteContext
    
    'Set BindCommand = Binding
    BindingColl.Add Binding
End Sub

'@Description "Binds Label Command"
Public Sub BindImageCommand(ByVal ViewModel As Object, ByVal Control As msForms.Image, ByVal Command As ICommand, _
                            Optional ByVal CanExecuteContext As Object, Optional ByVal ExecuteContext As Object)

    Dim Binding As ImageCommandBinding
    Set Binding = New ImageCommandBinding

    Binding.Initialize ViewModel, Control, Command, CanExecuteContext, ExecuteContext
    
    'Set BindCommand = Binding
    BindingColl.Add Binding
End Sub
'@Description "Binds the Value of a MSForms.Listbox to a Boolean source property"
Public Sub BindListBox(ByVal Control As msForms.listBox, ByVal SourceProperty As String, ByVal Source As Object)

    Dim Binding As ListBoxListBinding
    Set Binding = New ListBoxListBinding

    Binding.Initialize Control, Source, SourceProperty

    'Set BindCheckBox = Binding
    BindingColl.Add Binding
End Sub
