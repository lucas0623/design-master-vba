VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDataFormatConvertor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Operation.ImportData"
'@Interface

Option Explicit

Private mDfEle As clsDataFrame
Private mDfForce As clsDataFrame
Private mDfJoint As clsDataFrame


Public Function GetUserInput() As String
End Function

Public Function ReadData(filePath As String) As Integer
End Function

Public Property Get DfEle() As clsDataFrame

End Property

Public Property Let DfEle(value As clsDataFrame)

End Property

Public Property Get DfForce() As clsDataFrame
 
End Property

Public Property Let DfForce(value As clsDataFrame)

End Property

Public Property Get DfJoint() As clsDataFrame

End Property

Public Property Let DfJoint(value As clsDataFrame)

End Property
