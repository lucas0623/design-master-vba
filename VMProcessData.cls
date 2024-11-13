VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VMProcessData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Userform.ProcessData")
Option Explicit
Implements INotifyPropertyChanged
Private Notification As New PropertyChangeNotification
Private ds_sys As New DataSheetSystem
Private Type VMProp
    isCreateMember As Boolean
    isIdentifyRestraint As Boolean
    isMapBySection As Boolean
    isMapByKinkAngle As Boolean
    
    Commands As New Collection
End Type

Private this As VMProp

Private Sub OnPropertyChanged(ByVal Name As String)
    INotifyPropertyChanged_OnPropertyChanged Me, Name
End Sub
 
Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal Name As String)
    Notification.Notify Source, Name
End Sub
 
Private Sub INotifyPropertyChanged_RegisterHandler(ByVal handler As IHandlePropertyChanged)
    Notification.AddHandler handler
End Sub

Public Sub RegisterCommand(ByVal cmd As ICommand)
    this.Commands.Add cmd
End Sub

Public Function CountCommands() As Integer
    CountCommands = this.Commands.count
End Function

Public Sub Initialize()
    Dim ret As Integer
    ret = LoadLastSaveVal
    If Not ret = 0 Then LoadDefaultVal
End Sub

Public Sub LoadDefaultVal()
    Me.isCreateMember = True
    Me.isIdentifyRestraint = True
    Me.isMapBySection = True
    Me.isMapByKinkAngle = True
End Sub

Public Function LoadLastSaveVal() As Integer
    'Load Save Input if possible
    Dim propGrpName As String
    Dim ret As Integer
    
    propGrpName = "ProcessDataOperationSave"
    '1.Check if there is existing save. If yes, try loading the data
    If Not ds_sys.isPropGrpExist(propGrpName) Then
        ret = -1
        GoTo ExitFunc
    End If
    
    Me.isCreateMember = ds_sys.prop(propGrpName, "isCreateMember")
    Me.isIdentifyRestraint = ds_sys.prop(propGrpName, "isIdentifyRestraint")
    Me.isMapBySection = ds_sys.prop(propGrpName, "isMapBySection")
    Me.isMapByKinkAngle = ds_sys.prop(propGrpName, "isMapByKinkAngle")

ExitFunc:
    LoadLastSaveVal = ret
End Function

Public Sub SaveInput()
    'Save User Input to the
    Dim propGrpName As String
    
    propGrpName = "ProcessDataOperationSave"
    '1.Check if there is existing save. If yes, clear existing data. If no, create a new save
    If ds_sys.isPropGrpExist(propGrpName) Then
        'Clear Existing Data
        ds_sys.ClearPropGrpVal propGrpName
    End If
    
    '2. Write Data
    ds_sys.prop(propGrpName, "isCreateMember") = Me.isCreateMember
    ds_sys.prop(propGrpName, "isIdentifyRestraint") = Me.isIdentifyRestraint
    ds_sys.prop(propGrpName, "isMapBySection") = Me.isMapBySection
    ds_sys.prop(propGrpName, "isMapByKinkAngle") = Me.isMapByKinkAngle
    
End Sub

Public Property Get isCreateMember() As Boolean
    isCreateMember = this.isCreateMember
End Property

Public Property Let isCreateMember(rhs As Boolean)
    this.isCreateMember = rhs
    OnPropertyChanged "isCreateMember"
End Property

Public Property Get isIdentifyRestraint() As Boolean
    isIdentifyRestraint = this.isIdentifyRestraint
End Property

Public Property Let isIdentifyRestraint(ByVal rhs As Boolean)
    this.isIdentifyRestraint = rhs
    OnPropertyChanged "isIdentifyRestraint"
End Property

Public Property Get isMapBySection() As Boolean
    isMapBySection = this.isMapBySection
End Property
 
Public Property Let isMapBySection(ByVal rhs As Boolean)
    this.isMapBySection = rhs
    OnPropertyChanged "isMapBySection"
End Property

Public Property Get isMapByKinkAngle() As Boolean
    isMapByKinkAngle = this.isMapByKinkAngle
End Property

Public Property Let isMapByKinkAngle(ByVal rhs As Boolean)
    this.isMapByKinkAngle = rhs
    OnPropertyChanged "isMapByKinkAngle"
End Property
