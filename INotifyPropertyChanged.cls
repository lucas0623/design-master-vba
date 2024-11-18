VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "INotifyPropertyChanged"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder MVVM.Infrastructure
Option Explicit
Public Sub OnPropertyChanged(ByVal Source As Object, ByVal Name As String)
End Sub
Public Sub RegisterHandler(ByVal handler As IHandlePropertyChanged)
End Sub
