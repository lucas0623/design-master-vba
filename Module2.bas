Attribute VB_Name = "Module2"

'@Folder("Module")
'@Folder("VBAProject")
Sub ExportAllVBAComponents()
    Dim vbComp As Object
    Dim vbProj As Object
    Dim exportPath As String
    Dim fileName As String
    
    ' Set the export path
    exportPath = "C:\Users\lucasleung\OneDrive\12 Engineering\01 Resource Lib\04 Useful Excel Macro\Tools Addin\Useful Tools VBA Modules" ' Change this to your desired folder path
    
    ' Ensure the path ends with a backslash
    If Right(exportPath, 1) <> "\" Then exportPath = exportPath & "\"
    
    ' Loop through each component in the VBA project
    Debug.Print ThisWorkbook.name
    Set vbProj = ThisWorkbook.VBProject
    For Each vbComp In vbProj.VBComponents
        Select Case vbComp.Type
            Case 1 ' Standard Module
                fileName = exportPath & vbComp.name & ".bas"
                vbComp.Export fileName
            Case 2 ' Class Module
                fileName = exportPath & vbComp.name & ".cls"
                vbComp.Export fileName
            Case 3 ' UserForm
                fileName = exportPath & vbComp.name & ".frm"
                vbComp.Export fileName
            Case Else
                ' Handle other types if necessary
        End Select
    Next vbComp

    MsgBox "Export completed!"
End Sub

Sub TestDebugger()
    ' Create an instance of the Debugger class
    Dim dbg As New Debugger
    
    ' Test with a 1D array
    Dim arr1D(1 To 3) As Integer
    arr1D(1) = 1
    arr1D(2) = 20
    arr1D(3) = 30
    dbg.Log arr1D, "arr1D"
    
    ' Test with a 2D array
    Dim arr2D(1 To 2, 1 To 2) As String
    arr2D(1, 1) = "A"
    arr2D(1, 2) = "B"
    arr2D(2, 1) = "C"
    arr2D(2, 2) = "D"
    dbg.Log arr2D, "arr2D"
    
    ' Test with a Collection
    Dim coll As Collection
    Set coll = New Collection
    coll.Add "First", "Key1"
    coll.Add "Second", "Key2"
    coll.Add "Third", "Key3"
    dbg.Log coll, "coll"
    
    ' Test with a Dictionary
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "A", 1
    dict.Add "B", 2
    dict.Add "C", 3
    dbg.Log dict, "dict"
    
    ' Test with a nested Collection
    Dim nestedColl As Collection
    Set nestedColl = New Collection
    nestedColl.Add coll
    nestedColl.Add dict
    dbg.Log nestedColl, "nestedColl"
    
    ' Test with nested Dictionary
    Dim nestedDict As Object
    Set nestedDict = CreateObject("Scripting.Dictionary")
    nestedDict.Add "SubColl", coll
    nestedDict.Add "SubDict", dict
    dbg.Log nestedDict, "nestedDict"
    
    ' Test with a simple variable
    Dim simpleVar As String
    simpleVar = "Hello, World!"
    dbg.Log simpleVar, "simpleVar"
End Sub

Sub TestCreateNewSheet()
    'InitializeSummaryTemplate "Summary_TemplateEmpty"
    CreateSteelMemberDesignSummary "Summary_TemplateStlMem"
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
    
    ' Initialize    the worksheet with the specified content
    With ws
        ' Set the whole sheet font to 'Aptos'
        .Cells.Font.name = "Aptos Narrow"
        
        ' Set the headers and content
        .Cells(1, 1).value = "targetWS"
        .Cells(1, 2).value = "user input 1"
        .Cells(1, 3).value = "user input 2"
        .Cells(1, 4).value = "user input 3"
        
        .Cells(3, 1).value = "Summary of Design"
        .Cells(3, 1).Font.size = 18  ' Set font size of cell A3 to 18
        
        .Cells(4, 1).value = "Design Input"
        .Cells(4, 2).value = "Design Input"
        
        .Cells(5, 1).value = "Design Worksheet Name"
        .Cells(5, 2).value = "user input 1"
        .Cells(5, 3).value = "user input 2"
        .Cells(5, 4).value = "user input 3"
        
        ' Format the cells
        .rows("3:5").Font.Bold = True  ' Bold Rows 3 to 5
        .Columns("A:D").AutoFit
        
        ' Center across selection for Cells B4 to D4
        .Range("B4:D4").HorizontalAlignment = xlCenterAcrossSelection
        
        ' Add specific colors to each cell in Row 4
        .Cells(4, 1).Interior.Color = RGB(189, 215, 238)
        .Cells(4, 2).Interior.Color = RGB(198, 224, 180)
        .Cells(4, 3).Interior.Color = RGB(198, 224, 180)
        .Cells(4, 4).Interior.Color = RGB(198, 224, 180)
        
        ' Add specific colors to each cell in Row 5
        .Cells(5, 1).Interior.Color = RGB(221, 235, 247)
        .Cells(5, 2).Interior.Color = RGB(226, 239, 218)
        .Cells(5, 3).Interior.Color = RGB(226, 239, 218)
        .Cells(5, 4).Interior.Color = RGB(226, 239, 218)
        
        ' Add border to range A4:D5
        .Range("A4:D5").Borders.LineStyle = xlContinuous
    End With
    
    ' Set page setup options
    With ws.PageSetup
        On Error Resume Next
        .PaperSize = xlPaperA3
        If Err.Number <> 0 Then
            ' Default to A4 if A3 is not supported
            .PaperSize = xlPaperA4
        End If
        On Error GoTo 0
        .Orientation = xlLandscape
        .PrintTitleRows = "$3:$5"
        
        ' Set headers and footers
        .RightHeader = "&""Aptos,Regular""&11Page &P of &N"
        .RightFooter = "&""Aptos Narrow,Regular""&8Printed at &D &T" & Chr(10) & "&[Path]&[File]"
    End With
End Sub

Sub CreateSteelMemberDesignSummary(sheetName As String)
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
        ' Set the whole sheet font to 'Aptos'
        .Cells.Font.name = "Aptos Narrow"
        
        ' Set the headers and content
        .Cells(1, 1).value = "section"
        .Cells(1, 2).value = "eleName"
        .Cells(1, 3).value = "loadComb"
        .Cells(1, 4).value = "caseName"
        .Cells(1, 5).value = "P"
        .Cells(1, 6).value = "V2"
        .Cells(1, 7).value = "V3"
        .Cells(1, 8).value = "T"
        .Cells(1, 9).value = "M2"
        .Cells(1, 10).value = "M3"
        .Cells(1, 11).value = "targetWS"
        .Cells(1, 12).value = "Section Type"
        .Cells(1, 13).value = "Section Size"
        .Cells(1, 14).value = "Rolled/ Weld"
        .Cells(1, 15).value = "Steel Grade"
        .Cells(1, 16).value = "Eff Length (Axial, Major)"
        .Cells(1, 17).value = "Eff Length (Axial, Minor)"
        .Cells(1, 18).value = "Eff Length (LTB)"
        .Cells(1, 19).value = "Axial Uti"
        .Cells(1, 20).value = "Major Bend Uti"
        .Cells(1, 21).value = "Minor Bend Uti"
        .Cells(1, 22).value = "Overall Uti"
        .Cells(1, 23).value = "Slenderness"
        .Cells(1, 24).value = "Overall"
        .Cells(1, 25).value = "Calculation Title"

        .Cells(3, 1).value = "Steel Member Design Summary"
        .Cells(3, 1).Font.size = 18  ' Set font size of cell A3 to 18
        
        ' Merge and set titles
        .Range("A4").value = "Element Information"
        .Range("A4:J4").Interior.Color = RGB(255, 230, 153)
        .Range("A5:J5").Interior.Color = RGB(255, 242, 204)
        
        .Range("K4").value = "Design Input"
        .Range("K4:R4").Interior.Color = RGB(189, 215, 238)
        .Range("K5:R5").Interior.Color = RGB(221, 235, 247)
        
        .Range("S4").value = "Design Output"
        .Range("S4:X4").Interior.Color = RGB(198, 224, 180)
        .Range("S5:X5").Interior.Color = RGB(226, 239, 218)
        
        
        ' Set Row 5 values
        .Cells(5, 1).value = "Section"
        .Cells(5, 2).value = "Element Name"
        .Cells(5, 3).value = "Load Combination"
        .Cells(5, 4).value = "Correspondence Case"
        .Cells(5, 5).value = "Axial Force (kN) (+ve Tension)"
        .Cells(5, 6).value = "Shear Along y Axis (kN)"
        .Cells(5, 7).value = "Shear Along z Axis (kN)"
        .Cells(5, 8).value = "Torsion (kNm)"
        .Cells(5, 9).value = "Moment About y Axis (kNm)"
        .Cells(5, 10).value = "Moment About z Axis (kNm)"
        .Cells(5, 11).value = "Design Worksheet Name"
        .Cells(5, 12).value = "Design Section"
        .Cells(5, 13).value = "Design Size"
        .Cells(5, 14).value = "Rolled/ Welded"
        .Cells(5, 15).value = "Grade"
        .Cells(5, 16).value = "Eff. Length (for Buckling along y axis) (mm)"
        .Cells(5, 17).value = "Eff. Length (for Buckling along x axis) (mm)"
        .Cells(5, 18).value = "Eff. Length for LTB due to Moment Mx (mm)"
        .Cells(5, 19).value = "Axial Utilization (%)"
        .Cells(5, 20).value = "Bending Mx Utilization (%)"
        .Cells(5, 21).value = "Bending My Utilization (%)"
        .Cells(5, 22).value = "Overall Utilization (%)"
        .Cells(5, 23).value = "Slenderness Ratio"
        .Cells(5, 24).value = "Overall"
        .Cells(5, 25).value = "Calculation Title"

        ' Format the cells
        .rows("3:5").Font.Bold = True  ' Bold Rows 3 to 5
        .Columns("A:Y").AutoFit
        
        ' Center across selection for merged cells
        .Range("A4:J4").HorizontalAlignment = xlCenterAcrossSelection
        .Range("K4:R4").HorizontalAlignment = xlCenterAcrossSelection
        .Range("S4:X4").HorizontalAlignment = xlCenterAcrossSelection

        ' Add border to range A4:Y5
        .Range("A4:Y5").Borders.LineStyle = xlContinuous
    End With

    ' Set page setup options
    With ws.PageSetup
        On Error Resume Next
        .PaperSize = xlPaperA3
        If Err.Number <> 0 Then
            ' Default to A4 if A3 is not supported
            .PaperSize = xlPaperA4
        End If
        On Error GoTo 0
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
