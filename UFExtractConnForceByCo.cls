VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UFExtractConnForceByCo"
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
Private mcboMaxMinEnv(5) As msForms.checkbox

Private mcboPosition As msForms.ComboBox
Private mcbCase(11) As msForms.checkbox
Private mcbIsEachLC As msForms.checkbox

Private mlstConnType As msForms.listBox, mlstConnTypeExcl As msForms.listBox, mbtnConnType1 As msForms.CommandButton, mbtnConnType2 As msForms.CommandButton
Private mlstNode As msForms.listBox, mlstNodeExcl As msForms.listBox, mbtnNode1 As msForms.CommandButton, mbtnNode2 As msForms.CommandButton
Private mlstLc As msForms.listBox, mlstLcExcl As msForms.listBox, mbtnLc1 As msForms.CommandButton, mbtnLc2 As msForms.CommandButton

Private WithEvents mbtnLoadInput As msForms.CommandButton
Attribute mbtnLoadInput.VB_VarHelpID = -1
Private mImageOkButton As msForms.Image
Private miCloseState As Integer

Private Sub Class_Initialize()

    Set genFuncs = New clsGeneralFunctions
    
    Set form1 = New UFBasicMulti
    form1.Initialize 4, genFuncs.CStr_arr(Split("General,Connection Type Filter,Node Filter,Load Comb Filter", ",")), 400, False
    form1.TitleBarCaption = "Extract Connection Data To Summary Table"
    
    form1.CurrentPage = 1
    
    form1.AddComboBox_Empty mcboOutputSheetName, "Extract Force to Summary Worksheet:"
    form1.AddComboBox Split("By Node,By Connection Type", ","), mcboGroupBy, "Output to be Grouped by:"
    form1.AddComboBox Split("Correspondence Cases", ","), mcboForceType, "Output Force Cases:"
    
    form1.AddCheckBox_double mcbCase(0), mcbCase(1), "Please select the correspondence case(s) for output:", "Max P", "Min P", True, True, _
                            tipText:="Example: Case 'Max P'= Locate the position for Max P under certain Load Comb of certain Section/Member --> Extract all permutations of forces of that Load Comb at that location."
    form1.AddCheckBox_double mcbCase(2), mcbCase(3), vbNullString, "Max Vy", "Min Vy", True, True
    form1.AddCheckBox_double mcbCase(4), mcbCase(5), vbNullString, "Max Vz", "Min Vz", True, True
    form1.AddCheckBox_double mcbCase(6), mcbCase(7), vbNullString, "Max T", "Min T", True, True
    form1.AddCheckBox_double mcbCase(8), mcbCase(9), vbNullString, "Max Mz", "Min Mz", True, True
    form1.AddCheckBox_double mcbCase(10), mcbCase(11), vbNullString, "Max My", "Min My", True, True
    
    form1.AddCheckBox mcbIsEachLC, "Output result seperately for EACH load combination?", isCheck:=False, _
                    tipText:="Checked: for Each Selected Section/Member --> for Each Selected Load Combination --> 1 sets of force and moments"

    form1.AddCheckBox_double mcboMaxMinEnv(0), mcboMaxMinEnv(3), "Envelope Option for Max & Min Result", "P", "T", False, True, _
    "Checked --> consider ABS MAX for the force/moment; Unchecked--> All Permutations will be Considered.  Example: Unchecked P and M3 for a Env Load Combination: Max/Min P/M3 will create 4 permutations"
'    "For Env Load Combination (LC) Output Data, there are 'max' and 'min' (2 rows of data) for certain element at certain location under certain LC." & vbNewLine & _
'    "Hence, there will be 2^6 permutations of design data, if it is uncertain whether 'max' or 'min' is more critical. (keep all boxex unchecked for this cases)" & vbNewLine & _
'    "In Opposite, if only would like to consider the abs max of certain force(s)/moment(s), then check the corresponding box(es). "
    form1.AddCheckBox_double mcboMaxMinEnv(1), mcboMaxMinEnv(4), "", "Vz (V2) ", "Mz (M2)", True, True
    form1.AddCheckBox_double mcboMaxMinEnv(2), mcboMaxMinEnv(5), "", "Vy (V3) ", "My (M3)", True, True
    
    form1.CurrentPage = 2
    form1.AddSelectionBoxMulti_Empty mlstConnType, "SELECTED Connection Types", height_LB:=300, _
                                    is_reListBox2:=True, reListBox2:=mlstConnTypeExcl, _
                                    title2:="EXCLUDED Connection Types", _
                                    reCmdBtn1:=mbtnConnType1, reCmdBtn2:=mbtnConnType2
    form1.CurrentPage = 3
    form1.AddSelectionBoxMulti_Empty mlstNode, "SELECTED Nodes", height_LB:=300, _
                                    is_reListBox2:=True, reListBox2:=mlstNodeExcl, _
                                    title2:="EXCLUDED Nodes", _
                                    reCmdBtn1:=mbtnNode1, reCmdBtn2:=mbtnNode2  ', isCreateFrame:=True, frameTitle:="FRAMES"
    
    form1.CurrentPage = 4
    form1.AddSelectionBoxMulti_Empty mlstLc, "SELECTED Load Comb.", height_LB:=300, _
                                    is_reListBox2:=True, reListBox2:=mlstLcExcl, _
                                    title2:="EXCLUDED Load Comb.", _
                                    reCmdBtn1:=mbtnLc1, reCmdBtn2:=mbtnLc2

    form1.AdjustHeight
    Set mbtnLoadInput = form1.AddButton("Load Default", width:=100, height:=20, top:=388.5, left:=10, pageNum:=-1)
    Set mImageOkButton = form1.OKButton
'    Debug.Print form1.OKButtonActive.top
'    Debug.Print form1.OKButtonActive.top + form1.OKButtonActive.height
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
    this.Bindings.BindComboBox mcboGroupBy, "GroupByNodeOrConnType", this.ViewModel
    
    this.Bindings.BindCheckBox mcbCase(0), "ExtremeCase0", this.ViewModel
    this.Bindings.BindCheckBox mcbCase(1), "ExtremeCase1", this.ViewModel
    this.Bindings.BindCheckBox mcbCase(2), "ExtremeCase2", this.ViewModel
    this.Bindings.BindCheckBox mcbCase(3), "ExtremeCase3", this.ViewModel
    this.Bindings.BindCheckBox mcbCase(4), "ExtremeCase4", this.ViewModel
    this.Bindings.BindCheckBox mcbCase(5), "ExtremeCase5", this.ViewModel
    this.Bindings.BindCheckBox mcbCase(6), "ExtremeCase6", this.ViewModel
    this.Bindings.BindCheckBox mcbCase(7), "ExtremeCase7", this.ViewModel
    this.Bindings.BindCheckBox mcbCase(8), "ExtremeCase8", this.ViewModel
    this.Bindings.BindCheckBox mcbCase(9), "ExtremeCase9", this.ViewModel
    this.Bindings.BindCheckBox mcbCase(10), "ExtremeCase10", this.ViewModel
    this.Bindings.BindCheckBox mcbCase(11), "ExtremeCase11", this.ViewModel
    this.Bindings.BindCheckBox mcboMaxMinEnv(0), "MaxMinEnv0", this.ViewModel
    this.Bindings.BindCheckBox mcboMaxMinEnv(1), "MaxMinEnv1", this.ViewModel
    this.Bindings.BindCheckBox mcboMaxMinEnv(2), "MaxMinEnv2", this.ViewModel
    this.Bindings.BindCheckBox mcboMaxMinEnv(3), "MaxMinEnv3", this.ViewModel
    this.Bindings.BindCheckBox mcboMaxMinEnv(4), "MaxMinEnv4", this.ViewModel
    this.Bindings.BindCheckBox mcboMaxMinEnv(5), "MaxMinEnv5", this.ViewModel
    this.Bindings.BindCheckBox mcbIsEachLC, "IsEachLC", this.ViewModel
    
    Dim MoveListBoxItemCommand As New MoveListBoxItemCommand
    
    this.Bindings.BindListBox mlstConnType, "ConnTypes", this.ViewModel
    this.Bindings.BindListBox mlstConnTypeExcl, "ConnTypesExcl", this.ViewModel
    this.Bindings.BindCommand this.ViewModel, mbtnConnType1, _
                            MoveListBoxItemCommand.Create(mlstConnType, mlstConnTypeExcl, this.ViewModel, "ConnTypes", "ConnTypesExcl")
    this.Bindings.BindCommand this.ViewModel, mbtnConnType2, _
                            MoveListBoxItemCommand.Create(mlstConnTypeExcl, mlstConnType, this.ViewModel, "ConnTypesExcl", "ConnTypes")
    
    this.Bindings.BindListBox mlstNode, "Nodes", this.ViewModel
    this.Bindings.BindListBox mlstNodeExcl, "NodesExcl", this.ViewModel
    this.Bindings.BindCommand this.ViewModel, mbtnNode1, _
                            MoveListBoxItemCommand.Create(mlstNode, mlstNodeExcl, this.ViewModel, "Nodes", "NodesExcl")
    this.Bindings.BindCommand this.ViewModel, mbtnNode2, _
                            MoveListBoxItemCommand.Create(mlstNodeExcl, mlstNode, this.ViewModel, "NodesExcl", "Nodes")
                            
    this.Bindings.BindListBox mlstLc, "lc", this.ViewModel
    this.Bindings.BindListBox mlstLcExcl, "lcExcl", this.ViewModel
    this.Bindings.BindCommand this.ViewModel, mbtnLc1, _
                            MoveListBoxItemCommand.Create(mlstLc, mlstLcExcl, this.ViewModel, "lc", "lcExcl")
    this.Bindings.BindCommand this.ViewModel, mbtnLc2, _
                            MoveListBoxItemCommand.Create(mlstLcExcl, mlstLc, this.ViewModel, "lcExcl", "lc")

    Dim SwitchUFValueCommand As New SwitchUFValueCommand
    this.Bindings.BindCommand this.ViewModel, mbtnLoadInput, SwitchUFValueCommand.Create(this.ViewModel), ExecuteContext:=mbtnLoadInput ', this.ViewModel, ExecuteContext:=mbtnLoadInput
    Dim SaveInputCommand As New SaveInputCommand
    this.Bindings.BindImageCommand this.ViewModel, mImageOkButton, SaveInputCommand.Create(this.ViewModel), ExecuteContext:=mImageOkButton
End Sub

Private Property Get IView_ViewModel() As Object
    Set IView_ViewModel = this.ViewModel
End Property

Private Property Get IView_CloseState() As Integer
    IView_CloseState = miCloseState
End Property





