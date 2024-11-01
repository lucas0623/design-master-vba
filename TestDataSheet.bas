Attribute VB_Name = "TestDataSheet"
''@TestModule
''@Folder("Tests")
'
'
'Option Explicit
'Option Private Module
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
''@TestMethod("DataSheet.Init")
'Private Sub ReadJtData()
'    Dim dsManager As DataSheetManager
'    Set dsManager = New DataSheetManager
'
'    Dim dsJt As oDataSheet
'    Set dsJt = dsManager.DSJointCoor
'
'    Dim var As Variant
'    With dsJt.tagSelector
'        var = dsJt.GetTagsText(.ID, .x, .y, .z)
'        If var(0) = .ID And var(3) = .z Then
'            Assert.Succeed
'        Else
'            Assert.Fail "Runtime Error"
'        End If
'    End With
'
'TestExit:
'    Application.ScreenUpdating = True
'    Application.Calculation = xlAutomatic
'    'mLog.CloseFile
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
'End Sub
