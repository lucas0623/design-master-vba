VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UFExtractForceMemForGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Userform.ForceExtraction")

'********************************************************
'This is the userform for the 'extract frame force' operation
'This class contains only the UI elements only. The input/ process of UF data is in the 'UFControl' Class

'Arthor: Lucas LEUNG
'Update Log
'Aug 2023 - Initial
'*******************************************************
Option Explicit
Implements IView

Private Type TView
    'IView state:
    ViewModel As VMForceExtraction

    'ICancellable state:
    'IsCancelled As Boolean

    'Data binding helper dependency:
    Bindings As New BindingsManager
End Type
 
Private this As TView

Private genFuncs As clsGeneralFunctions
Private form1 As UFBasicMulti
Private mcboOutputSheetName As msForms.ComboBox
Private mcboGroupBy As msForms.ComboBox
Private mcboForceType As msForms.ComboBox

Private mlstFrame As msForms.listBox, mlstFrameExcl As msForms.listBox, mbtnFrame1 As msForms.CommandButton, mbtnFrame2 As msForms.CommandButton
Private mlstLc As msForms.listBox, mlstLcExcl As msForms.listBox, mbtnLc1 As msForms.CommandButton, mbtnLc2 As msForms.CommandButton
Private mImageOkButton As msForms.Image
Private WithEvents mbtnLoadInput As msForms.CommandButton
Attribute mbtnLoadInput.VB_VarHelpID = -1
Private miCloseState As Integer

Private Sub Class_Initialize()

    Set genFuncs = New clsGeneralFunctions
    
    Set form1 = New UFBasicMulti
    form1.Initialize 3, genFuncs.CStr_arr(Split("General,Member Filter,Load Comb Filter", ",")), 400, False
    form1.TitleBarCaption = "Extract Frame Data To Summary Table (for Plot Graph)"
    
    
    form1.CurrentPage = 1
    
    form1.AddComboBox_Empty mcboOutputSheetName, "Extract Force to Summary Worksheet:"
    
    form1.AddComboBox Split("By Member", ","), mcboGroupBy, "Output to be Grouped by:"
    form1.AddComboBox Split("All Forces", ","), mcboForceType, "Output Force Cases:"

    
    form1.CurrentPage = 2
    form1.AddSelectionBoxMulti_Empty mlstFrame, "SELECTED Members", height_LB:=300, _
                                    is_reListBox2:=True, reListBox2:=mlstFrameExcl, _
                                    title2:="EXCLUDED Members", _
                                    reCmdBtn1:=mbtnFrame1, reCmdBtn2:=mbtnFrame2  ', isCreateFrame:=True, frameTitle:="FRAMES"
    
    form1.CurrentPage = 3
    form1.AddSelectionBoxMulti_Empty mlstLc, "SELECTED Load Comb.", height_LB:=300, _
                                    is_reListBox2:=True, reListBox2:=mlstLcExcl, _
                                    title2:="EXCLUDED Load Comb.", _
                                    reCmdBtn1:=mbtnLc1, reCmdBtn2:=mbtnLc2

    form1.AdjustHeight
    Set mbtnLoadInput = form1.AddButton("Load Default", width:=100, height:=20, top:=376.5, left:=10, pageNum:=-1)
    Set mImageOkButton = form1.OKButton
    'form1.AddEvent mbtnLoadInput
    
End Sub

Public Sub IView_Initialize(vm As Object)
    Set this.ViewModel = vm
    'Set this.Bindings = New PropertyBindings
    'this.Bindings.Initialize VM
    mcboOutputSheetName.List = this.ViewModel.wsList
    ConfigureBindings
End Sub

Public Sub IView_Show()
    form1.Show
    If Not form1.CloseState = 0 Then GoTo CloseProcedure
    Exit Sub
CloseProcedure:
    miCloseState = -1
End Sub

Private Sub ConfigureBindings()
 
    Const EnabledProperty As String = "Enabled"
    Const ListProperty As String = "List"

    this.Bindings.BindComboBox mcboOutputSheetName, "wsSumName", this.ViewModel
    'this.Bindings.BindComboBox mcboGroupBy, "GroupBy", this.ViewModel

    Dim MoveListBoxItemCommand As New MoveListBoxItemCommand
       
    this.Bindings.BindListBox mlstFrame, "MemberNames", this.ViewModel
    this.Bindings.BindListBox mlstFrameExcl, "MemberNamesExcl", this.ViewModel
    this.Bindings.BindCommand this.ViewModel, mbtnFrame1, _
                            MoveListBoxItemCommand.Create(mlstFrame, mlstFrameExcl, this.ViewModel, "MemberNames", "MemberNamesexcl")
    this.Bindings.BindCommand this.ViewModel, mbtnFrame2, _
                            MoveListBoxItemCommand.Create(mlstFrameExcl, mlstFrame, this.ViewModel, "MemberNamesExcl", "MemberNames")
                            
    this.Bindings.BindListBox mlstLc, "lc", this.ViewModel
    this.Bindings.BindListBox mlstLcExcl, "lcExcl", this.ViewModel
    this.Bindings.BindCommand this.ViewModel, mbtnLc1, _
                            MoveListBoxItemCommand.Create(mlstLc, mlstLcExcl, this.ViewModel, "lc", "lcExcl")
    this.Bindings.BindCommand this.ViewModel, mbtnLc2, _
                            MoveListBoxItemCommand.Create(mlstLcExcl, mlstLc, this.ViewModel, "lcExcl", "lc")
'    this.Bindings.BindListBox mlstSec, "Sections", this.ViewModel
'    this.Bindings.BindListBox mlstSecExcl, "SectionsExcl", this.ViewModel
'    this.Bindings.BindListBox mlstFrame, "MemberNames", this.ViewModel
'    this.Bindings.BindListBox mlstFrameExcl, "MemberNamesexcl", this.ViewModel
'    this.Bindings.BindListBox mlstLc, "lc", this.ViewModel
'    this.Bindings.BindListBox mlstLcExcl, "lcExcl", this.ViewModel

    Dim SwitchUFValueCommand As New SwitchUFValueCommand
    this.Bindings.BindCommand this.ViewModel, mbtnLoadInput, SwitchUFValueCommand.Create(this.ViewModel), ExecuteContext:=mbtnLoadInput ', this.ViewModel, ExecuteContext:=mbtnLoadInput
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




