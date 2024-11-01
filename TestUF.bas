Attribute VB_Name = "TestUF"
'@Folder "Tests.TestModule"

'********************************************************
'Arthor: Lucas LEUNG
'Update Log
'Aug 2023 - Initial
'*******************************************************

Sub TestUF()
    Dim var As Variant
    Dim form1 As UFBasic
    
    'Try Load user last input, if fail, then load default

    '1st page of the userform
    Dim outputSheetName As Variant, cb_outputSheetName As msForms.ComboBox, def_sheetName As String
    Dim sGroup As Variant, cb_group As msForms.ComboBox, def_group As String
    Dim sForceType As Variant, cb_sForceType As msForms.ComboBox
    Dim cb_pMaxMinEnv(5) As msForms.checkbox
    
    Dim sPosition As Variant, cb_position As msForms.ComboBox
    Dim sCaseQuickToggle As Variant, cb_caseQuickToggle As msForms.ComboBox
    Dim cb_cCase(11) As msForms.checkbox, cb_checkAll As msForms.checkbox
    Dim arrNum_all(11) As Long, arrNum_PnM(5) As Long, arrNum_major(7) As Long
    
    Dim pIsEachLC_cb As msForms.checkbox
    
    Dim arr(1) As Object
    Dim i As Integer
    outputSheetName = Split("Summary_Sample,DesignData_Sample", ",")
    sGroup = Split("by Section,by Member", ",")
    sForceType = Split("Correspondence Cases", ",")
    sPosition = Split("NO Filter (ALL Positions),Both Ends ONLY,I-End Only,J-End Only", ",")
    sCaseQuickToggle = Split("ALL,Major Axis Plane Forces,Axial and Bending,Custom", ",")
    
    For i = 0 To 11
        arrNum_all(i) = i
    Next i
    
    arrNum_PnM(0) = 0
    arrNum_PnM(1) = 1
    arrNum_PnM(2) = 8
    arrNum_PnM(3) = 9
    arrNum_PnM(4) = 10
    arrNum_PnM(5) = 11
    
    arrNum_major(0) = 0
    arrNum_major(1) = 1
    arrNum_major(2) = 2
    arrNum_major(3) = 3
    arrNum_major(4) = 6
    arrNum_major(5) = 7
    arrNum_major(6) = 10
    arrNum_major(7) = 11

    Set form1 = New UFBasic
    form1.TitleBarCaption = "Extract Frame Data To Summary Table - Correspondence Cases"
    form1.width = 300
    'form1.Height = 600
    form1.AddComboBox outputSheetName, cb_outputSheetName, "Extract Force to Summary Worksheet:", "Summary_Sample"
    form1.AddComboBox sGroup, cb_group, "Output to be Grouped by:"
    form1.AddComboBox sForceType, cb_sForceType, "Output Force Cases:"
    form1.AddComboBox sPosition, cb_position, "Position Filter:"
    
    form1.AddCheckBox_double cb_pMaxMinEnv(0), cb_pMaxMinEnv(3), "Envelope Option for Max & Min Result", "P", "T", False, True
    form1.AddCheckBox_double cb_pMaxMinEnv(1), cb_pMaxMinEnv(4), "", "Vz (V2) ", "Mz (M2)", True, True
    form1.AddCheckBox_double cb_pMaxMinEnv(2), cb_pMaxMinEnv(5), "", "Vy (V3) ", "My (M3)", True, True
'
    form1.AddComboBox sCaseQuickToggle, cb_caseQuickToggle, "Please select the correspondence case(s) for output:"
    form1.AddCheckBox_double cb_cCase(0), cb_cCase(1), "", "Max P", "Min P", True, True
    form1.AddCheckBox_double cb_cCase(2), cb_cCase(3), "", "Max Vy", "Min Vy", True, True
    form1.AddCheckBox_double cb_cCase(4), cb_cCase(5), "", "Max Vz", "Min Vz", True, True
    form1.AddCheckBox_double cb_cCase(6), cb_cCase(7), "", "Max T", "Min T", True, True
    form1.AddCheckBox_double cb_cCase(8), cb_cCase(9), "", "Max Mz", "Min Mz", True, True
    form1.AddCheckBox_double cb_cCase(10), cb_cCase(11), "", "Max My", "Min My", True, True
    form1.AddCheckBox pIsEachLC_cb, "Output result seperately for EACH load combination?", isCheck:=False
    form1.AdjustHeight
    form1.Show
    
    'If Not form1.closeState = 0 Then GoTo CloseProcedure
 
    '2nd page of the userform
    'Create user form for getting user input range
'    Dim form2 As UF_BasicUserform
'    Set form2 = New UF_BasicUserform
'    Dim lb_sec As MSForms.listBox, lb_frame As MSForms.listBox
'    Dim lb_loadComb As MSForms.listBox
'
'
'    form2.Caption = "Extract Frame Data To Summary Table - Correspondence Cases"
'    form2.AddSelectionBoxMulti pSections, lb_sec, "SELECTED Sections", "EXCLUDED Sections", is_reListBox2:=False  ', isCreateFrame:=True, frameTitle:="pSections"
'    form2.AddSelectionBoxMulti pMemberNames, lb_frame, "SELECTED Member", "EXCLUDED Member", is_reListBox2:=False ', isCreateFrame:=True, frameTitle:="FRAMES"
'    form2.Show
'    If Not form2.closeState = 0 Then GoTo CloseProcedure
End Sub

