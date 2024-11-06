VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ViewWorkbookStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Operation")

Private dsManager As New DataSheetManager ', model As StrModel
Private DsSys As New DataSheetSystem

Private wsInteract As New clsWorksheetsInteraction
Private genFunc As New clsGeneralFunctions
Private UI As New clsUIManager

Sub ViewWorkbookStatus()
    Dim form1 As UFBasic

    Set form1 = New UFBasic
    form1.Initialize 350, True
    form1.TitleBarCaption = "View Workbook Status"
    
    'Data Imported Status
    Dim arr As Variant, arr_header As Variant
    ReDim arr(3, 1)
    arr(0, 0) = "Joint/Node Coordinates"
    arr(1, 0) = "Frame/Element Connectivity"
    arr(2, 0) = "Frame/Element Section Assignment"
    arr(3, 0) = "Joint/Node Restraint Assignment"
    arr(0, 1) = CStr(DsSys.prop("isWSImported", "ws_joint"))
    arr(1, 1) = CStr(DsSys.prop("isWSImported", "ws_frame"))
    arr(2, 1) = CStr(DsSys.prop("isWSImported", "ws_frameForce"))
    arr(3, 1) = CStr(DsSys.prop("isWSImported", "ws_restraint"))
    arr_header = Array("Data Type", "Is Imported?")
    form1.AddMultiColumnsListBox arr, arr_header, width_LB:=320, colWidth:="220;80", title:="Imported Data Status", height_LB:=60
    
    'Member creation status
    Dim arr2 As Variant, arr_header2 As Variant
    ReDim arr2(4, 1)
    arr2(0, 0) = "Is Member Created?"
    arr2(1, 0) = "   Applied Condition: Frame Connectivity?"
    arr2(2, 0) = "   Applied Condition: Matching Section?"
    arr2(3, 0) = "   Applied Condition: Limiting Kink Angle?"
    arr2(4, 0) = "Is Identified Restraint Along Member?"
    arr2(0, 1) = CStr(DsSys.prop("isCreated", "member"))
    If DsSys.prop("isCreated", "member") Then
        arr2(1, 1) = CStr(DsSys.prop("MemberCreateMethod", "FrameConnectivity"))
        arr2(2, 1) = CStr(DsSys.prop("MemberCreateMethod", "MatchSection"))
        arr2(3, 1) = CStr(DsSys.prop("MemberCreateMethod", "LimitKinkAngle"))
    Else
        arr2(1, 1) = "False"
        arr2(2, 1) = "False"
        arr2(3, 1) = "False"
    End If
    arr2(4, 1) = CStr(DsSys.prop("isCreated", "isIdentifiedConn"))
    arr_header2 = Array("Item", "True/False")
    form1.AddMultiColumnsListBox arr2, arr_header2, width_LB:=320, colWidth:="220;80", title:="Member Creation Status", height_LB:=65
    
    'Connection Creation Status
    Dim arr3 As Variant, arr_header3 As Variant
    ReDim arr3(1, 1)
    arr3(0, 0) = "Is Identified Connection Node?"
    arr3(1, 0) = "Is Mapped Connection Type to Node?"
    arr3(0, 1) = CStr(DsSys.prop("isCreated", "isIdentifiedConn"))
    arr3(1, 1) = CStr(DsSys.prop("isCreated", "isMappedConn"))
    arr_header3 = Array("Item", "True/False")
    form1.AddMultiColumnsListBox arr3, arr_header3, width_LB:=320, colWidth:="220;80", title:="Connection Creation Status", height_LB:=40
    
    'What Operation Can be carried out
    Dim arr4 As Variant, arr_header4 As Variant
    ReDim arr4(5, 1)
    arr4(0, 0) = "Create Member By Section"
    arr4(1, 0) = "Create Member By Kink Angle"
    arr4(2, 0) = "Identify Restraint along Members"
    arr4(3, 0) = "Map Connection"
    arr4(4, 0) = "Extract Frame Force"
    arr4(5, 0) = "Extract Connection Force"
    
    arr4(0, 1) = CStr(DsSys.prop("WhatOperationCanProcess", "CreateMemberBySection"))
    arr4(1, 1) = CStr(DsSys.prop("WhatOperationCanProcess", "CreateMemberByKinkAngle"))
    arr4(2, 1) = CStr(DsSys.prop("WhatOperationCanProcess", "IdentifyRestraint"))
    arr4(3, 1) = CStr(DsSys.prop("WhatOperationCanProcess", "MapConnection"))
    arr4(4, 1) = CStr(DsSys.prop("WhatOperationCanProcess", "ExtractFrameForce"))
    arr4(5, 1) = CStr(DsSys.prop("WhatOperationCanProcess", "ExtractConnectionForce"))
    
    arr_header4 = Array("Operation", "Status/ Can I Run?")
    form1.AddMultiColumnsListBox arr4, arr_header4, width_LB:=320, colWidth:="220;80", title:="What Operation Can Be Carried Out?", height_LB:=90
    
    'Design Worksheet Output Data
    form1.AdjustHeight
    form1.Show
End Sub

