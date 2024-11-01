Attribute VB_Name = "TestFormStrObj"
''@TestModule
''@Folder("Tests")
'
'
'Option Explicit
'Option Private Module
'
'#Const LateBind = LateBindTests
'
'#If LateBind Then
'    Private Assert As Object
'    Private Fakes As Object
'#Else
'    Private Assert As Rubberduck.AssertClass
'    Private Fakes As Rubberduck.FakesProvider
'#End If
'
''@ModuleInitialize
'Private Sub ModuleInitialize()
'    'this method runs once per module.
'    #If LateBind Then
'        Set Assert = CreateObject("Rubberduck.AssertClass")
'        Set Fakes = CreateObject("Rubberduck.FakesProvider")
'    #Else
'        Set Assert = New Rubberduck.AssertClass
'        Set Fakes = New Rubberduck.FakesProvider
'    #End If
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
''@TestMethod("StrModel")
'Private Sub FormJt()
'    Dim model As StrModel, ret As Integer
'    Set model = New StrModel
'    ret = model.Constructor.FormJointObj
'    If ret = 0 Then
'        Assert.Succeed
'    Else
'        Assert.Fail
'    End If
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
'
''@TestMethod("StrModel")
'Private Sub FormFrmAll()
'    Dim model As StrModel, ret As Integer
'    Set model = New StrModel
'    ret = model.Constructor.FormFrmObj
'    If ret = 0 Then
'        Assert.Succeed
'    Else
'        Assert.Fail
'    End If
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
'
''@TestMethod("StrModel")
'Private Sub FormFrmForceAll()
'    Dim model As StrModel, ret As Integer
'    Set model = New StrModel
'    ret = model.Constructor.FormFrmForceObj
'    If ret = 0 Then
'        Assert.Succeed
'    Else
'        Assert.Fail
'    End If
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
