VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder MVVM.Infrastructure
'@ModuleDescription "An object that represents an executable command."
'@Interface
'@Exposed
Option Explicit
 
'@Description "Returns True if the command is enabled given the provided binding context (ViewModel)."
Public Function CanExecute(ByVal Context As Object) As Boolean
End Function
 
'@Description "Executes the command given the provided binding context (ViewModel)."
Public Sub Execute(ByVal Context As Object)
End Sub
 
'@Description "Gets a user-friendly description of the command."
Public Property Get Description() As String
End Property
