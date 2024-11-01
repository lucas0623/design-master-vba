Attribute VB_Name = "TestImportGSA"
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
''Private mLog As clsLog
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
''@TestMethod("ImportGSA.ImportData")
'Private Sub TestImportGSADataSet1()
'    'Import Joint Table + Ele Table + Beam Force Table
'    'Expect: complete without error
'    Dim filePath As String
'    filePath = "C:\Users\leung\OneDrive - The University of Hong Kong\12 Engineering\01 Resource Lib\02 Design Spreadsheet\01 Structural Design\21 Design Master\Data File for Test\01 Test Import\GSA\DataSet01.csv"
'    TestImportGSADataSet filePath, 0
'End Sub
'
''@TestMethod("ImportGSA.ImportData")
'Private Sub TestImportGSADataSet2()
'    'Import Joint Table + Ele Table + Beam Force Table
'    'Expect: complete without error
'    Dim filePath As String
'    filePath = "C:\Users\lucasleung\OneDrive - connect.hku.hk\12 Engineering\01 Resource Lib\02 Design Spreadsheet\01 Structural Design\21 Design Master\Data File for Test\01 Test Import\GSA\DataSet02.csv"
'    TestImportGSADataSet filePath, 0
'End Sub
'
''@TestMethod("ImportGSA.ImportData")
'Private Sub TestImportGSADataSet3()
'    'Import Joint Table + Ele Table + Beam Force Table
'    'Expect: complete without error
'    Dim filePath As String
'    filePath = "C:\Users\lucasleung\OneDrive - connect.hku.hk\12 Engineering\01 Resource Lib\02 Design Spreadsheet\01 Structural Design\21 Design Master\Data File for Test\01 Test Import\GSA\DataSet03.csv"
'    TestImportGSADataSet filePath, 0
'End Sub
'
''@TestMethod("ImportGSA.ImportData")
'Private Sub TestImportGSADataSet4()
'    'Import Joint Table + Ele Table + Beam Force Table
'    'Expect: complete without error
'    Dim filePath As String
'    filePath = "C:\Users\lucasleung\OneDrive - connect.hku.hk\12 Engineering\01 Resource Lib\02 Design Spreadsheet\01 Structural Design\21 Design Master\Data File for Test\01 Test Import\GSA\DataSet04.csv"
'    TestImportGSADataSet filePath, -1
'End Sub
'
''@TestMethod("ImportGSA.ImportData")
'Private Sub TestImportGSADataSet5()
'    'Import Joint Table + Ele Table + Beam Force Table
'    'Expect: complete without error
'    Dim filePath As String
'    filePath = "C:\Users\lucasleung\OneDrive - connect.hku.hk\12 Engineering\01 Resource Lib\02 Design Spreadsheet\01 Structural Design\21 Design Master\Data File for Test\01 Test Import\GSA\DataSet05.csv"
'    TestImportGSADataSet filePath, -1
'End Sub
'
''@TestMethod("ImportGSA.ImportData")
'Private Sub TestImportGSADataSet6()
'    'Import Joint Table + Ele Table + Beam Force Table
'    'Expect: complete without error
'    Dim filePath As String
'    filePath = "C:\Users\lucasleung\OneDrive - connect.hku.hk\12 Engineering\01 Resource Lib\02 Design Spreadsheet\01 Structural Design\21 Design Master\Data File for Test\01 Test Import\GSA\DataSet06.csv"
'    TestImportGSADataSet filePath, 0
'End Sub
'
'
'Private Sub TestImportGSADataSet(filePath As String, expectedResult As Integer)
'    'Import Joint Table + Ele Table + Beam Force Table
'    'Expect: complete without error
'
'    On Error GoTo TestFail
'
'    'Initialize
'    Dim ret As Integer
'    Dim obj As IDataFormatConvertor
'    Set obj = New DataConvertorGSA
'
'    Fakes.MsgBox.Returns 42
'
'    ret = obj.ReadData(filePath)
'
'    If ret = expectedResult Then
'        Assert.Succeed
'    Else
'        Assert.Fail "Runtime Error"
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
