VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DesignWsIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Operation.SummaryToWS")

' The class module name should be "DesignWsIO"
Option Explicit

' Private member variables
Private pInput As Object ' Dictionary
Private pOutput As Object ' Dictionary
Private pName As String

' Class initializer
Private Sub Class_Initialize()
    Set pInput = CreateObject("Scripting.Dictionary")
    Set pOutput = CreateObject("Scripting.Dictionary")
End Sub

' Property for Input (Read-only)
Public Property Get WsInput() As Object
    Set WsInput = pInput
End Property

' Property for Output (Read-only)
Public Property Get WsOutput() As Object
    Set WsOutput = pOutput
End Property

' Property for name (Read and Write)
Public Property Get name() As String
    name = pName
End Property

Public Property Let name(value As String)
    pName = value
End Property

' Method to verify if a string is a cell address
Private Function isAddress(ByVal address As String) As Boolean
    On Error Resume Next
    isAddress = Not IsError(Range(address).address)
    On Error GoTo 0
End Function

' Method to add an input to the Input dictionary
Public Function AddInput(ByVal key As String, ByVal address As String) As Long
    If isKeyExist(pInput, key) Then GoTo ExitFunc
    If Not isAddress(address) Then GoTo ExitFunc
    pInput.Add key, address
    Exit Function
    
ExitFunc:
    AddInput = -1
End Function

' Method to add an output to the Output dictionary
Public Function AddOutput(ByVal key As String, ByVal address As String) As Long
    If isKeyExist(pOutput, key) Then GoTo ExitFunc
    If Not isAddress(address) Then GoTo ExitFunc
    pOutput.Add key, address
    Exit Function
    
ExitFunc:
    AddOutput = -1
End Function

' Helper function to check if a key exists in a dictionary
Private Function isKeyExist(dict As Object, sKey As String) As Boolean
    isKeyExist = dict.exists(sKey)
End Function

' Method to initialize the class properties
Public Sub Initialize(ByVal name As String)
    Me.name = name
    Set pInput = CreateObject("Scripting.Dictionary")
    Set pOutput = CreateObject("Scripting.Dictionary")
End Sub
