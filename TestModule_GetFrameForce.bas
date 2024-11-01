Attribute VB_Name = "TestModule_GetFrameForce"
'Option Explicit
'Option Private Module
'
''@TestModule
''@Folder("Tests")
'
'Private Assert As Object
'Private Fakes As Object
'
''@ModuleInitialize
'Private Sub ModuleInitialize()
'    'this method runs once per module.
'    Set Assert = CreateObject("Rubberduck.AssertClass")
'    Set Fakes = CreateObject("Rubberduck.FakesProvider")
'End Sub
'
''@ModuleCleanup
'Private Sub ModuleCleanup()
'    'this method runs once per module.
'    Set Assert = Nothing
'    Set Fakes = Nothing
'End Sub
'
''@TestInitialize
'Private Sub TestInitialize()
'    'This method runs before every test in the module..
'End Sub
'
''@TestCleanup
'Private Sub TestCleanup()
'    'this method runs after every test in the module.
'End Sub
'
''@TestMethod("clsStrFrame_GetFrameForcesMethod")
'Private Sub TestGetFrameForce_StationFilter_1LC()
'    '1 load comb, 1 step type, 1 station filter
'
'    On Error GoTo TestFail
'
'    'Initialize
'    Dim cFrms As New Collection, cFrmForces As New Collection
'    Dim frm As StrFrame, jtI As StrJoint, jtJ As StrJoint
'    Dim frmForce As StrFrameForce
'    Dim i As Long, Position As Double
'
'    Set jtI = New StrJoint
'    Set jtJ = New StrJoint
'    Set frm = New StrFrame
'    frm.Init "C1", "PB1", jtI, jtJ, length:=10
'
'
'    For i = 0 To 10
'        Set frmForce = New StrFrameForce
'        frmForce.Init2 frm, CDbl(i), "C1", "Max", 100, 100, 100, 1001, 100, 100
'        frm.AddFrameForceToColl frmForce
'    Next i
'    Set cFrmForces = frm.frmForces
'    Position = 2
'    Set frmForce = frm.GetFrameForceAtStation(cFrmForces, Position)
'    If frmForce.station = Position Then
'        Assert.Succeed
'    Else
'        Assert.Fail "Runtime Error"
'    End If
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
'End Sub
'
''@TestMethod("clsStrFrame_GetFrameForcesMethod")
'Private Sub TestGetFrameForce_StationFilter_3LC()
'    '3 load comb, 1 step type, 1 station filter
'
'    On Error GoTo TestFail
'
'    'Initialize
'    Dim cFrms As New Collection, cFrmForces As New Collection
'    Dim frm As StrFrame, jtI As StrJoint, jtJ As StrJoint
'    Dim frmForce As StrFrameForce
'    Dim i As Long, Position As Double
'
'    Set jtI = New StrJoint
'    Set jtJ = New StrJoint
'    Set frm = New StrFrame
'    frm.Init "C1", "PB1", jtI, jtJ, length:=10
'
'
'    For i = 0 To 10
'        Set frmForce = New StrFrameForce
'        frmForce.Init2 frm, CDbl(i), "C1", "Max", 100, 100, 100, 1001, 100, 100
'        frm.AddFrameForceToColl frmForce
'    Next i
'    For i = 0 To 10
'        Set frmForce = New StrFrameForce
'        frmForce.Init2 frm, CDbl(i), "C2", "Max", 100, 100, 100, 1001, 100, 100
'        frm.AddFrameForceToColl frmForce
'    Next i
'    For i = 0 To 10
'        Set frmForce = New StrFrameForce
'        frmForce.Init2 frm, CDbl(i), "C3", "Max", 100, 100, 100, 1001, 100, 100
'        frm.AddFrameForceToColl frmForce
'    Next i
'
'    Set cFrmForces = frm.frmForces
'    Position = 2
'    Set cFrmForces = frm.GetFrameForces(SpecifiedStation:=Position)
'    If cFrmForces.count = 3 Then
'        Assert.Succeed
'    Else
'        Assert.Fail "Runtime Error"
'    End If
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
'End Sub
'
''@TestMethod("clsStrFrame_GetFrameForcesMethod")
'Private Sub TestGetFrameForce_StationFilter_2LC_2StepType()
'    '2 load comb, 2 step type, 1 station filter
'
'    On Error GoTo TestFail
'
'    'Initialize
'    Dim cFrms As New Collection, cFrmForces As New Collection
'    Dim frm As StrFrame, jtI As StrJoint, jtJ As StrJoint
'    Dim frmForce As StrFrameForce
'    Dim i As Long, Position As Double
'
'    Set jtI = New StrJoint
'    Set jtJ = New StrJoint
'    Set frm = New StrFrame
'    frm.Init "C1", "PB1", jtI, jtJ, length:=10
'
'
'    For i = 0 To 10
'        Set frmForce = New StrFrameForce
'        frmForce.Init2 frm, CDbl(i), "C1", "Max", 100, 100, 100, 1001, 100, 100
'        frm.AddFrameForceToColl frmForce
'    Next i
'    For i = 0 To 10
'        Set frmForce = New StrFrameForce
'        frmForce.Init2 frm, CDbl(i), "C2", "Max", 100, 100, 100, 1001, 100, 100
'        frm.AddFrameForceToColl frmForce
'    Next i
'    For i = 0 To 10
'        Set frmForce = New StrFrameForce
'        frmForce.Init2 frm, CDbl(i), "C2", "Min", 100, 100, 100, 1001, 100, 100
'        frm.AddFrameForceToColl frmForce
'    Next i
'
'    Set cFrmForces = frm.frmForces
'    Position = 3
'    Set cFrmForces = frm.GetFrameForces(stepTypeFilter:="MAX", SpecifiedStation:=Position)
'    If cFrmForces.count = 2 Then
'        Assert.Succeed
'    Else
'        Assert.Fail "Runtime Error"
'    End If
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
'End Sub
'
''@TestMethod("clsStrModel_FrmForceAnalyser")
'Private Sub TestFindExtremeForce_oneCase_ret2()
'    '2 load comb, 2 step type, 1 station filter
'
'    On Error GoTo TestFail
'
'    'Initialize
'    Dim cFrms As New Collection, cFrmForces As New Collection
'    Dim frm As StrFrame, jtI As StrJoint, jtJ As StrJoint
'    Dim frmForce As StrFrameForce
'    Dim i As Long, Position As Double
'    Dim model As StrModel
'
'    Set model = New StrModel
'    Set jtI = New StrJoint
'    Set jtJ = New StrJoint
'    Set frm = New StrFrame
'    frm.Init "C1", "PB1", jtI, jtJ, length:=10
'    model.AddStrObjToColl frm, obj_frm
'
'    For i = 0 To 10
'        Set frmForce = New StrFrameForce
'        frmForce.Init2 frm, CDbl(i), "C1", "Max", 100, 100, 100, 1001, 100, 100
'        frm.AddFrameForceToColl frmForce
'        model.AddStrObjToColl frmForce, obj_frmForce
'    Next i
'    For i = 0 To 10
'        Set frmForce = New StrFrameForce
'        frmForce.Init2 frm, CDbl(i), "C1", "Min", -100, -100, -100, -1001, -100, -100
'        frm.AddFrameForceToColl frmForce
'        model.AddStrObjToColl frmForce, obj_frmForce
'    Next i
'    For i = 0 To 10
'        Set frmForce = New StrFrameForce
'        frmForce.Init2 frm, CDbl(i), "C2", "Max", 100, 100, 100, 1001, 100, 100
'        frm.AddFrameForceToColl frmForce
'        model.AddStrObjToColl frmForce, obj_frmForce
'    Next i
'    For i = 0 To 10
'        Set frmForce = New StrFrameForce
'        frmForce.Init2 frm, CDbl(i), "C2", "Min", -100, -100, -100, -1001, -100, -100
'        frm.AddFrameForceToColl frmForce
'        model.AddStrObjToColl frmForce, obj_frmForce
'    Next i
'
'    Set cFrmForces = frm.frmForces
'    Set cFrmForces = model.frmForceAnalyser.FindExtremeForce_oneCase(cFrmForces, MaxP)
'
'    If cFrmForces.count = 2 Then
'        Assert.Succeed
'    Else
'        Assert.Fail "Runtime Error"
'    End If
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
'End Sub
'
''@TestMethod("clsStrModel_FrmForceAnalyser")
'Private Sub TestFindExtremeForce_oneCase_ret1()
'    '2 load comb, 2 step type, 1 station filter
'
'    On Error GoTo TestFail
'
'    'Initialize
'    Dim cFrms As New Collection, cFrmForces As New Collection
'    Dim frm As StrFrame, jtI As StrJoint, jtJ As StrJoint
'    Dim frmForce As StrFrameForce
'    Dim i As Long, Position As Double
'    Dim model As StrModel
'
'    Set model = New StrModel
'    Set jtI = New StrJoint
'    Set jtJ = New StrJoint
'    Set frm = New StrFrame
'    frm.Init "C1", "PB1", jtI, jtJ, length:=10
'    model.AddStrObjToColl frm, obj_frm
'
'    For i = 0 To 10
'        Set frmForce = New StrFrameForce
'        frmForce.Init2 frm, CDbl(i), "C1", "", 1000, 100, 100, 1001, 100, 100
'        frm.AddFrameForceToColl frmForce
'        model.AddStrObjToColl frmForce, obj_frmForce
'    Next i
'    For i = 0 To 10
'        Set frmForce = New StrFrameForce
'        frmForce.Init2 frm, CDbl(i), "C3", "Min", -100, -100, -100, -1001, -100, -100
'        frm.AddFrameForceToColl frmForce
'        model.AddStrObjToColl frmForce, obj_frmForce
'    Next i
'    For i = 0 To 10
'        Set frmForce = New StrFrameForce
'        frmForce.Init2 frm, CDbl(i), "C2", "Max", 100, 100, 100, 1001, 100, 100
'        frm.AddFrameForceToColl frmForce
'        model.AddStrObjToColl frmForce, obj_frmForce
'    Next i
'    For i = 0 To 10
'        Set frmForce = New StrFrameForce
'        frmForce.Init2 frm, CDbl(i), "C2", "Min", -100, -100, -100, -1001, -100, -100
'        frm.AddFrameForceToColl frmForce
'        model.AddStrObjToColl frmForce, obj_frmForce
'    Next i
'
'    Set cFrmForces = frm.frmForces
'    Set cFrmForces = model.frmForceAnalyser.FindExtremeForce_oneCase(cFrmForces, MaxP)
'
'    If cFrmForces.count = 1 Then
'        Assert.Succeed
'    Else
'        Assert.Fail "Runtime Error"
'    End If
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
'End Sub
'
''@TestMethod("clsStrModel_FrmForceAnalyser")
'Private Sub TestFindExtremeForce()
'    '2 load comb, 2 step type, 1 station filter
'
'    On Error GoTo TestFail
'
'    'Initialize
'    Dim cFrms As New Collection, cFrmForces As New Collection
'    Dim frm As StrFrame, jtI As StrJoint, jtJ As StrJoint
'    Dim frmForce As StrFrameForce
'    Dim i As Long, Position As Double
'    Dim model As StrModel
'
'    Set model = New StrModel
'    Set jtI = New StrJoint
'    Set jtJ = New StrJoint
'    Set frm = New StrFrame
'    frm.Init "C1", "PB1", jtI, jtJ, length:=10
'    model.AddStrObjToColl frm, obj_frm
'
'    For i = 0 To 10
'        Set frmForce = New StrFrameForce
'        frmForce.Init2 frm, CDbl(i), "C1", "", 1000, 100, 100, 1001, 100, 100
'        frm.AddFrameForceToColl frmForce
'        model.AddStrObjToColl frmForce, obj_frmForce
'    Next i
'    For i = 0 To 10
'        Set frmForce = New StrFrameForce
'        frmForce.Init2 frm, CDbl(i), "C3", "Min", -100, -100, -100, -1001, -100, -100
'        frm.AddFrameForceToColl frmForce
'        model.AddStrObjToColl frmForce, obj_frmForce
'    Next i
'    For i = 0 To 10
'        Set frmForce = New StrFrameForce
'        frmForce.Init2 frm, CDbl(i), "C2", "Max", 100, 100, 100, 1001, 100, 100
'        frm.AddFrameForceToColl frmForce
'        model.AddStrObjToColl frmForce, obj_frmForce
'    Next i
'    For i = 0 To 10
'        Set frmForce = New StrFrameForce
'        frmForce.Init2 frm, CDbl(i), "C2", "Min", -100, -100, -100, -1001, -100, -100
'        frm.AddFrameForceToColl frmForce
'        model.AddStrObjToColl frmForce, obj_frmForce
'    Next i
'
'    Set cFrmForces = frm.frmForces
'    Set cFrmForces = model.frmForceAnalyser.FindExtremeForce_oneCase(cFrmForces, MaxP)
'
'    If cFrmForces.count = 1 Then
'        Assert.Succeed
'    Else
'        Assert.Fail "Runtime Error"
'    End If
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
'End Sub
