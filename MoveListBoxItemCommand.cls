VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MoveListBoxItemCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder Userform.Command
'@ModuleDescription "A command that move the selected items in ListBox1 to ListBox2, and update corresponding properties in the VM."
Option Explicit
Implements ICommand
Private genFuncs As New clsGeneralFunctions

Private Type TState
    ViewModel As Object
    ListBox1 As msForms.listBox
    ListBox2 As msForms.listBox
    PropName1 As String
    PropName2 As String
End Type
 
Private this As TState

'@Description "Creates a new instance of this command."
Public Function Create(ByVal ListBox1 As msForms.listBox, ByVal ListBox2 As msForms.listBox, ViewModel As Object, _
                        VMProp1 As String, VMProp2 As String) As ICommand
    Dim result As MoveListBoxItemCommand
    Set result = New MoveListBoxItemCommand
    Set result.ViewModel = ViewModel
    Set result.ListBox1 = ListBox1
    Set result.ListBox2 = ListBox2
    result.PropName1 = VMProp1
    result.PropName2 = VMProp2
    
    result.ViewModel.RegisterCommand result
    Set Create = result
End Function

Public Property Get ViewModel() As Object
    Set ViewModel = this.ViewModel
End Property

Public Property Set ViewModel(ByVal rhs As Object)
    Set this.ViewModel = rhs
End Property

Public Property Get ListBox1() As msForms.listBox
    Set ListBox1 = this.ListBox1
End Property

Public Property Set ListBox1(ByVal rhs As msForms.listBox)
    Set this.ListBox1 = rhs
End Property

Public Property Get ListBox2() As msForms.listBox
    Set ListBox2 = this.ListBox2
End Property

Public Property Set ListBox2(ByVal rhs As msForms.listBox)
    Set this.ListBox2 = rhs
End Property

Public Property Get PropName1() As String
    PropName1 = this.PropName1
End Property

Public Property Let PropName1(ByVal rhs As String)
    this.PropName1 = rhs
End Property

Public Property Get PropName2() As String
    PropName2 = this.PropName2
End Property

Public Property Let PropName2(ByVal rhs As String)
    this.PropName2 = rhs
End Property

Private Function ICommand_CanExecute(ByVal Context As Object) As Boolean
    'If mViewModel.isLoadDefault Then
    ICommand_CanExecute = True
    'End If
End Function

Private Property Get ICommand_Description() As String
    ICommand_Description = "Trasnfer Selected Items from ListBox 1 to ListBox 2."
End Property

Private Sub ICommand_Execute(ByVal Context As Object)
    Dim counter As Integer
    Dim i As Integer

    counter = 0
    For i = 0 To this.ListBox1.ListCount - 1
        If this.ListBox1.Selected(i - counter) Then
            this.ListBox2.AddItem this.ListBox1.List(i - counter)
            this.ListBox1.RemoveItem (i - counter)
            counter = counter + 1
        End If
    Next i
    
    VBA.Interaction.CallByName this.ViewModel, this.PropName1, VbLet, genFuncs.ListBoxToArray(this.ListBox1)
    VBA.Interaction.CallByName this.ViewModel, this.PropName2, VbLet, genFuncs.ListBoxToArray(this.ListBox2)
End Sub

