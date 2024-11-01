Attribute VB_Name = "TestImportSAP"
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
''@TestMethod("ImportSAP.ImportData")
'Private Sub TestImportSAPDataSet1()
'    'Import Joint Table + Ele Table + Beam Force Table
'    'Expect: complete without error
'    Dim filePath As String
'    filePath = "C:\Users\leung\OneDrive - The University of Hong Kong\12 Engineering\01 Resource Lib\02 Design Spreadsheet\01 Structural Design\21 Design Master\Data File for Test\01 Test Import\SAP\DataSet01.xlsx"
'    TestImportDataSet filePath, 0
'End Sub
'
''@TestMethod("ImportSAP.ImportData")
'Private Sub TestImportSAPDataSet2()
'    'Import Joint Table + Ele Table + Beam Force Table
'    'Expect: complete without error
'    Dim filePath As String
'    filePath = "C:\Users\leung\OneDrive - The University of Hong Kong\12 Engineering\01 Resource Lib\02 Design Spreadsheet\01 Structural Design\21 Design Master\Data File for Test\01 Test Import\SAP\DataSet02.xlsx"
'    TestImportDataSet filePath, 0
'End Sub
'
''@TestMethod("ImportSAP.ImportData")
'Private Sub TestImportSAPDataSet3()
'    'Import Joint Table + Ele Table + Beam Force Table
'    'Expect: complete without error
'    Dim filePath As String
'    filePath = "C:\Users\leung\OneDrive - The University of Hong Kong\12 Engineering\01 Resource Lib\02 Design Spreadsheet\01 Structural Design\21 Design Master\Data File for Test\01 Test Import\SAP\DataSet03.xlsx"
'    TestImportDataSet filePath, 0
'End Sub
'
''@TestMethod("ImportSAP.ImportData")
'Private Sub TestImportSAPDataSet4()
'    'Import Joint Table + Ele Table + Beam Force Table
'    'Expect: complete without error
'    Dim filePath As String
'    filePath = "C:\Users\leung\OneDrive - The University of Hong Kong\12 Engineering\01 Resource Lib\02 Design Spreadsheet\01 Structural Design\21 Design Master\Data File for Test\01 Test Import\SAP\DataSet04.xlsx"
'    TestImportDataSet filePath, 0
'End Sub
'
''@TestMethod("ImportSAP.ImportData")
'Private Sub TestImportSAPDataSet5()
'    'Import Joint Table + Ele Table + Beam Force Table
'    'Expect: complete without error
'    Dim filePath As String
'    filePath = "C:\Users\leung\OneDrive - The University of Hong Kong\12 Engineering\01 Resource Lib\02 Design Spreadsheet\01 Structural Design\21 Design Master\Data File for Test\01 Test Import\SAP\DataSet05.xlsx"
'    TestImportDataSet filePath, -1
'End Sub
'
''@TestMethod("ImportSAP.ImportData")
'Private Sub TestImportSAPDataSet6()
'    Dim filePath As String
'    filePath = "C:\Users\leung\OneDrive - The University of Hong Kong\12 Engineering\01 Resource Lib\02 Design Spreadsheet\01 Structural Design\21 Design Master\Data File for Test\01 Test Import\SAP\DataSet06.xlsx"
'    TestImportDataSet filePath, -1
'End Sub
'
''@TestMethod("ImportSAP.ImportData")
'Private Sub TestImportSAPDataSet7()
'    'Import Joint Table + Ele Table + Beam Force Table
'    'Expect: complete without error
'    Dim filePath As String
'    filePath = "C:\Users\leung\OneDrive - The University of Hong Kong\12 Engineering\01 Resource Lib\02 Design Spreadsheet\01 Structural Design\21 Design Master\Data File for Test\01 Test Import\SAP\DataSet07.xlsx"
'    TestImportDataSet filePath, 0
'End Sub
'
'Private Sub TestImportDataSet(filePath As String, expectedResult As Integer)
'    'Import Joint Table + Ele Table + Beam Force Table
'    'Expect: complete without error
'
'    On Error GoTo TestFail
'
'    'Initialize
'    Dim ret As Integer
'    Dim obj As IDataFormatConvertor
'    Set obj = New DataConvertorSAP
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
'
