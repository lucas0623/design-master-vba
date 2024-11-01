VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PropertyChangeNotification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder MVVM.Infrastructure
Option Explicit
Private Handlers As VBA.Collection
 
Public Sub AddHandler(ByVal handler As IHandlePropertyChanged)
    Handlers.Add handler
End Sub
 
Public Sub Notify(ByVal Source As Object, ByVal name As String)
    Dim handler As IHandlePropertyChanged
    For Each handler In Handlers
        handler.OnPropertyChanged Source, name
    Next
End Sub
 
Private Sub Class_Initialize()
    Set Handlers = New VBA.Collection
End Sub
