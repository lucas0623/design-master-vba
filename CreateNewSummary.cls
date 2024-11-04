VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CreateNewSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Operation"
Option Explicit
Private newSheet As Worksheet
Public Sub CreateEmptySummary()
    CopyHiddenSheet "Summary_TemplateEmpty"
End Sub
Public Sub CreateForGeneralPurpose()

    CopyHiddenSheet "Summary_TemplateGeneral"
End Sub
Public Sub CreateForSteelMember()
    CopyHiddenSheet "Summary_TemplateStlMem"
End Sub
Public Sub CreateForSteelConnection()
    CopyHiddenSheet "Summary_TemplateStlConn"
End Sub
Public Sub CreateForRC()
    CopyHiddenSheet "Summary_TemplateRC"
End Sub

Private Sub CopyHiddenSheet(name As String)
    Dim oldSheet As Worksheet
    Set oldSheet = Worksheets(name)
    oldSheet.Visible = xlSheetVisible
    oldSheet.Copy Before:=ActiveSheet
    oldSheet.Visible = xlSheetHidden
End Sub


Sub InitializeSummaryTemplate(sheetName As String)
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim uniqueSheetName As String
    
    ' Reference the current workbook
    Set wb = ThisWorkbook
    
    ' Ensure a unique sheet name
    uniqueSheetName = GetUniqueSheetName(sheetName)
    
    ' Always create a new sheet with a unique name
    Set ws = wb.Worksheets.Add
    ws.name = uniqueSheetName
    
    ' Initialize the worksheet with the specified content
    With ws
        ' Set the headers and content
        .Cells(1, 1).value = "targetWS"
        .Cells(1, 2).value = "user input 1"
        .Cells(1, 3).value = "user input 2"
        .Cells(1, 4).value = "user input 3"
        
        .Cells(3, 1).value = "Summary of Design"
        .Cells(3, 1).Font.size = 18  ' Set font size of cell A3 to 18
        
        .Cells(4, 1).value = "Design Input"
        .Cells(4, 3).value = "Design Input"
        
        .Cells(5, 1).value = "Design Worksheet Name"
        .Cells(5, 2).value = "user input 1"
        .Cells(5, 3).value = "user input 2"
        .Cells(5, 4).value = "user input 3"
        
        ' Format the cells
        .rows("3:5").Font.Bold = True  ' Bold Rows 3 to 5
        .Columns("A:D").AutoFit
        
        ' Add specific colors to each cell in Row 4
        .Cells(4, 1).Interior.Color = RGB(255, 255, 0) ' Yellow for cell A4
        .Cells(4, 2).Interior.Color = RGB(255, 255, 0) ' Yellow for cell B4
        .Cells(4, 3).Interior.Color = RGB(255, 255, 0) ' Yellow for cell C4
        .Cells(4, 4).Interior.Color = RGB(255, 255, 0) ' Yellow for cell D4
        
        ' Add specific colors to each cell in Row 5
        .Cells(5, 1).Interior.Color = RGB(173, 216, 230) ' Light blue for cell A5
        .Cells(5, 2).Interior.Color = RGB(173, 216, 230) ' Light blue for cell B5
        .Cells(5, 3).Interior.Color = RGB(173, 216, 230) ' Light blue for cell C5
        .Cells(5, 4).Interior.Color = RGB(173, 216, 230) ' Light blue for cell D5
    End With
    
    ' Set page setup options
    With ws.PageSetup
        .PaperSize = xlPaperA3
        .Orientation = xlLandscape
        .PrintTitleRows = "$3:$5"
        
        ' Set headers and footers
        .RightHeader = "&""Aptos,Regular""&11Page &P of &N"
        .RightFooter = "&""Aptos Narrow,Regular""&8Printed at &D &T" & Chr(10) & "&[Path]&[File]"
    End With
End Sub

Function GetUniqueSheetName(baseName As String) As String
    Dim newName As String
    Dim counter As Integer
    counter = 1
    newName = baseName
    
    ' Loop until a unique name is found
    Do While WorksheetExists(newName)
        newName = baseName & "_" & counter
        counter = counter + 1
    Loop
    
    GetUniqueSheetName = newName
End Function

Function WorksheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    WorksheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function
