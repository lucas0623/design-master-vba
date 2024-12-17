Attribute VB_Name = "Module1"
Function IsCellAddress(str As String) As Boolean
    Dim regex As Object
    Dim match As Object
    Dim columnPart As String
    Dim rowPart As String
    Dim columnNumber As Long
    Dim rowNumber As Long
    
    Dim var As Variant, CellAddress As String
    var = Split(str, "!")
    CellAddress = var(UBound(var))
    ' Create the regex object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Regular expression to match full cell addresses
    ' Example: '[Workbook.xlsx]Sheet'!$A$1 or just A1
    regex.Pattern = "(\$?[A-Z]+\$?\d+)"
    regex.IgnoreCase = True
    regex.Global = False
    
    ' Check if the cell address matches the pattern
    If Not regex.TEST(CellAddress) Then
        IsCellAddress = False
        Exit Function
    End If
    
    ' Extract the actual cell reference (e.g., A1 from $A$1)
    regex.Pattern = "\$?([A-Z]+)\$?(\d+)"
    Set match = regex.Execute(CellAddress)
    
    If match.count = 0 Then
        IsCellAddress = False
        Exit Function
    End If
    
    columnPart = match(0).SubMatches(0)
    rowPart = match(0).SubMatches(1)
    
    ' Convert column part to a column number
    columnNumber = 0
    Dim i As Long
    For i = 1 To Len(columnPart)
        columnNumber = columnNumber * 26 + (Asc(UCase(Mid(columnPart, i, 1))) - Asc("A") + 1)
    Next i
    
    ' Convert row part to a row number
    rowNumber = CLng(rowPart)
    
    ' Validate the column and row numbers against Excel's limits
    If columnNumber >= 1 And columnNumber <= 16384 And rowNumber >= 1 And rowNumber <= 1048576 Then
        IsCellAddress = True
    Else
        IsCellAddress = False
    End If
End Function

Function GetAddressWorksheetName(str As String) As String
    Dim regex As Object
    Dim match As Object
    Dim lastExclaimation As Integer, lastBracket As Integer
    Dim wsName As String
    
    If Not IsCellAddress(str) Then Exit Function

    lastExclaimation = FindLastCharInStr(str, "!")
    If lastExclaimation = 0 Then Exit Function 'no worksheet name
    
    lastBracket = FindLastCharInStr(str, "]")
    
    wsName = Mid(str, lastBracket + 1, lastExclaimation - lastBracket - 1)
    
    If left(wsName, 1) = "'" Then
        wsName = Mid(wsName, 2, Len(wsName) - 1)
    End If
    
    If Right(wsName, 1) = "'" Then
        wsName = Mid(wsName, 1, Len(wsName) - 1)
    End If
    
    GetAddressWorksheetName = wsName
End Function

Sub Test_isCellAddress()
    Debug.Print "Test 1 - A1: "; IsCellAddress("A1") ' Expected: True
    Debug.Print "Test 2 - AA100: "; IsCellAddress("AA100") ' Expected: True
    Debug.Print "Test 3 - XFD1048576: "; IsCellAddress("XFD1048576") ' Expected: True
    Debug.Print "Test 4 - Sheet1!A1: "; IsCellAddress("Sheet1!A1") ' Expected: True
    Debug.Print "Test 5 - 'Sheet1'!A1: "; IsCellAddress("'Sheet1'!A1") ' Expected: True
    Debug.Print "Test 6 - Sheet1!A: "; IsCellAddress("Sheet1!A") ' Expected: False
    Debug.Print "Test 7 - A1048577: "; IsCellAddress("A1048577") ' Expected: False
    Debug.Print "Test 8 - 123: "; IsCellAddress("123") ' Expected: False
    Debug.Print "Test 9 - ABCDE: "; IsCellAddress("ABCDE") ' Expected: False
    Debug.Print "Test 10 - 'Sheet!Name'!B12: "; IsCellAddress("'Sheet!Name'!B12") ' Expected: True
    Debug.Print "Test 11 - $A1: "; IsCellAddress("$A1") ' Expected: True
End Sub
Sub Test_GetAddressWorksheet()
    Debug.Print "Test 1 - A1: "; GetAddressWorksheetName("A1") ' Expected: ""
    Debug.Print "Test 2 - Sheet1!A1: "; GetAddressWorksheetName("Sheet1!A1") ' Expected: "Sheet1"
    Debug.Print "Test 3 - 'Sheet1'!A1: "; GetAddressWorksheetName("'Sheet1'!A1") ' Expected: "Sheet1"
    Debug.Print "Test 4 - 'Long Sheet Name'!B2: "; GetAddressWorksheetName("'Long Sheet Name'!B2") ' Expected: "Long Sheet Name"
    Debug.Print "Test 5 - Invalid Address: "; GetAddressWorksheetName("Invalid!A1") ' Expected: ""
    Debug.Print "Test 6 - 'Sheet!Name'!B12: "; GetAddressWorksheetName("'Sheet!Name'!B12") ' Expected: "Sheet!Name"
    Debug.Print "Test 7 - '[Connection Summary 20241113c.xlsx]Summary_BasePla''''''te!'!$A$1: "; GetAddressWorksheetName("'[Connection Summary 20241113c.xlsx]Summary_BasePla''''''te!'!$A$1") ' Expected: "Summary_BasePla''''''te!"
End Sub

Sub TEST()
    Dim var As Variant
    var = Split("gfs!o2!1fm!d", "!")
    Debug.Print "array size = " & UBound(var) + 1
    
    Debug.Print var(UBound(var))
End Sub

Function FindLastCharInStr(str As String, findChar As String) As Long
    Dim i As Long

    ' Initialize the function to return 0 if no '!' is found
    FindLastExclamation = 0

    ' Loop through the string backwards
    For i = Len(str) To 1 Step -1
        If Mid(str, i, 1) = findChar Then
            FindLastCharInStr = i
            Exit Function
        End If
    Next i
End Function

Sub Test_FindLastExclamation()
    Debug.Print "Test 1 - Sheet1!A1!B2!: " & FindLastCharInStr("Sheet1!A1!B2!", "!") ' Expected: 11
    Debug.Print "Test 2 - No exclamation: " & FindLastCharInStr("No exclamation", "!") ' Expected: 0
    Debug.Print "Test 3 - Single exclamation!: " & FindLastCharInStr("Single exclamation!", "!") ' Expected: 19
End Sub
