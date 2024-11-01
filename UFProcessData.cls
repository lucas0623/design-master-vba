VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UFProcessData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Userform.ProcessData")
'********************************************************
'Arthor: Lucas LEUNG
'Update Log
'Aug 2023 - Initial
'*******************************************************
Option Explicit
Implements IView

Private Type TView
    'IView state:
    ViewModel As VMProcessData
    'Data binding helper dependency:
    Bindings As New BindingsManager
End Type
 
Private this As TView

Private genFuncs As New clsGeneralFunctions
Private form1 As New UFBasic

Private mcbCreateMember As msForms.checkbox, mcbMapByFrameConnectivity As msForms.checkbox
Private mcbMapBySection As msForms.checkbox, mcbMapByKinkAngle As msForms.checkbox

Private mcbIdentifyRestraint As msForms.checkbox, mcbMapConnection As msForms.checkbox
Private mImageOkButton As msForms.Image
Private mbtnLoadInput As msForms.CommandButton

Private miCloseState As Integer

Private Sub Class_Initialize()

    form1.Initialize 300, True '3, genFuncs.CStr_arr(Split("General,Member Filter,Load Comb Filter", ",")), 400, False
    form1.TitleBarCaption = "Create Data Option"

    form1.AddLabel "DATA TO BE CREATED", fontSize:=11
    form1.AddSpace 2
    form1.AddCheckBox mcbCreateMember, title:="Member?", isCheck:=True, labelWidth:=140
    form1.AddSpace -4
    form1.AddCheckBox mcbIdentifyRestraint, title:="Restraint of Members", isCheck:=True, labelWidth:=140
    'form1.AddCheckBox mcbMapConnection, title:="Is Map Connection Type to Nodes?", isCheck:=False
    form1.AddSpace
    form1.AddLabel "CREATE MEMBER SETTING", fontSize:=11
    form1.AddCheckBox mcbMapByFrameConnectivity, title:="When the End-J of a frame connected to End-I of another frame", isCheck:=True, labelWidth:=220
    mcbMapByFrameConnectivity.Locked = True
    form1.AddSpace 4
    form1.AddCheckBox mcbMapBySection, title:="When 2 frames are the same Section", isCheck:=True, labelWidth:=220
    form1.AddCheckBox mcbMapByKinkAngle, title:="When the kink angle between 2 frames < 30 Deg", isCheck:=True, labelWidth:=220

    form1.AdjustHeight
    'Set mbtnLoadInput = form1.AddButton("Load Default", width:=100, height:=20, top:=230, left:=10, pageNum:=-1)
    Set mImageOkButton = form1.OKButton
End Sub


Private Sub IView_Initialize(vm As Object)
    Set this.ViewModel = vm
    ConfigureBindings
End Sub

Private Sub IView_Show()
    form1.Show
    If Not form1.CloseState = 0 Then GoTo CloseProcedure
    Exit Sub
CloseProcedure:
    miCloseState = -1
End Sub

Private Sub ConfigureBindings()
 
    Const EnabledProperty As String = "Enabled"
    Const ListProperty As String = "List"

    this.Bindings.BindCheckBox mcbCreateMember, "isCreateMember", this.ViewModel
    this.Bindings.BindCheckBox mcbIdentifyRestraint, "isIdentifyRestraint", this.ViewModel
    this.Bindings.BindCheckBox mcbMapBySection, "isMapBySection", this.ViewModel
    this.Bindings.BindCheckBox mcbMapByKinkAngle, "isMapByKinkAngle", this.ViewModel


    Dim SwitchUFValueCommand As New SwitchUFValueCommand
    
    
    'this.Bindings.BindCommand this.ViewModel, mbtnLoadInput, SwitchUFValueCommand.Create(this.ViewModel), ExecuteContext:=mbtnLoadInput ', this.ViewModel, ExecuteContext:=mbtnLoadInput
    Dim SaveInputCommand As New SaveInputCommand
    this.Bindings.BindImageCommand this.ViewModel, mImageOkButton, SaveInputCommand.Create(this.ViewModel), ExecuteContext:=mImageOkButton
    
'    Dim SaveInputCommand As New SaveInputCommand
'    this.Bindings.BindCommand this.ViewModel, mbtnLoadInput, SaveInputCommand.Create(this.ViewModel) ', this.ViewModel, ExecuteContext:=mbtnLoadInput
    'this.Bindings.BindCommand this.ViewModel, mbtnLoadInput, this.ViewModel.SwitchUFValueCommand
    'Debug.Print mbtnLoadInput.caption
End Sub

Private Property Get IView_ViewModel() As Object
    Set IView_ViewModel = this.ViewModel
End Property

Private Property Get IView_CloseState() As Integer
    IView_CloseState = miCloseState
End Property

'Private Property Let IView_ViewModel(vm As Object)
'    Set IView_ViewModel = this.ViewModel
'End Property





