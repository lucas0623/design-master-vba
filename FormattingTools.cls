VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormattingTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Operation")
Option Explicit
Private wsInteract As clsWorksheetsInteraction

Public Sub AddBorder_Hori()

    Dim myRow As Range
'    Dim num_lastRow As Long
'    Dim str_lastRow As String
'    Dim num_currentRow As Long
'    Dim str_currentRow As String
    Dim fRow As Long, lRow As Long, cRow As Long
    Dim lastRowOfSameGroup As Long
    Dim colNum As Long, colRef As Range
    Dim count As Long
    Dim inputBox As Variant
    Dim lineType As Long
    
    lineType = Application.inputBox( _
                                        prompt:="PLEASE INPUT THE LINE TYPE NUMBER. " & vbNewLine & vbNewLine _
                                        & "NOTE" & vbNewLine _
                                        & "Input 1 for thin line  " & vbNewLine _
                                        & "Input 2 for thick line" _
                                        & "Input 3 for double line", _
                                        default:=1, _
                                        Type:=1)
    If lineType = 0 Then Exit Sub
    Set colRef = AddInputBox_rng( _
                                        prompt:="PLEASE SELECT THE COLUMN. " & vbNewLine & vbNewLine _
                                        & "NOTE" & vbNewLine _
                                        & "1. The column must be selected. " & vbNewLine _
                                        & "2. The bottom border will be added whenever the data in the current row of the SELECTED COLUMN is different from that in the last row")
    If colRef Is Nothing Then Exit Sub
    count = 0

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    colNum = colRef.column
    fRow = Selection.row
    lRow = Selection.row + Selection.count - 1
    
    For cRow = fRow To lRow
        lastRowOfSameGroup = FindLastRowOfSameGrp(cRow, colNum, maxRow:=lRow)
        If lastRowOfSameGroup = -1 Then Exit For
        cRow = lastRowOfSameGroup
        If lineType = 1 Then
            With rows(lastRowOfSameGroup).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = 1
            End With
        ElseIf lineType = 2 Then
            With rows(lastRowOfSameGroup).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = 1
            End With
        ElseIf lineType = 3 Then
            With rows(lastRowOfSameGroup).Borders(xlEdgeBottom)
                .LineStyle = xlDouble
                .Weight = xlThick
                .ColorIndex = 1
            End With
        Else:
            MsgBox ("Please Input Valid Line Type Number")
            End
        End If

    Next cRow
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
End Sub

Sub BoldMaxOfSameGroup()

    Dim myRow As Range
    Dim lRow As Long
    'Dim str_lastRow As String
    Dim cRow As Long
    Dim lRowOfSameGrp As Long
    'Dim str_currentRow As String
    Dim colNum As Long
    Dim colNum2 As Long 'the column number to determine the maximum value
    Dim count As Long
    Dim inputBox As Variant
    Dim i As Long, j As Long
    
    Dim maxValue As Double 'a temp variable to compare value
    Dim cValue As Double
    Dim row_maxValue As Long
    Dim colRef As Range, colRef2 As Range
    'Dim lineType As Integer
    
    
    Set colRef = AddInputBox_rng(prompt:="PLEASE SELECT THE COLUMN FOR DETERMINING GROUP. ")
    If colRef Is Nothing Then Exit Sub
    Set colRef2 = AddInputBox_rng(prompt:="PLEASE SELECT THE COLUMN FOR IDENTIFYING MAXIMUM VALUE. ")
    If colRef2 Is Nothing Then Exit Sub
    colNum = colRef.column
    colNum2 = colRef2.column
    count = 0
    
    cRow = Selection.row
    lRow = Selection.row + Selection.count - 1
    On Error GoTo NextIteration:
    For i = cRow To lRow
    
        lRowOfSameGrp = FindLastRowOfSameGrp(i, colNum, maxRow:=lRow)
        If lRowOfSameGrp = -1 Then GoTo NextIteration
        
        If Not IsNumeric(Cells(i, colNum2).value2) Then
            i = lRowOfSameGrp
            GoTo NextIteration
        End If
        cValue = Cells(i, colNum2).value2
        maxValue = cValue
        row_maxValue = i
        
        For j = i To lRowOfSameGrp
            cValue = Cells(j, colNum2)
            If cValue > maxValue Then
                maxValue = cValue
                row_maxValue = j
            End If
        Next j
        rows(row_maxValue).Font.Bold = True
        i = lRowOfSameGrp
        
NextIteration:
    Next i

End Sub

Sub ClearBorder_Hori()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Dim myRow As Range
    
    For Each myRow In Selection.rows
    
        With myRow.Borders(xlEdgeBottom)
            .LineStyle = xlLineStyleNone
        End With
    Next
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
End Sub

Sub ClearBorder_Right()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Dim myCol As Range
    
    For Each myCol In Selection.Columns
    
        With myCol.Borders(xlEdgeRight)
            .LineStyle = xlLineStyleNone
        End With
    Next
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
End Sub
Sub GetMaxOfSameGroup()

    Dim myRow As Range
    Dim lRow As Long
    'Dim str_lastRow As String
    Dim cRow As Long
    Dim lRowOfSameGrp As Long
    'Dim str_currentRow As String
    Dim colNum As Long
    Dim colNum2 As Long 'the column number to determine the maximum value
    Dim count As Long
    Dim inputBox As Variant
    Dim i As Long, j As Long
    
    Dim maxValue As Double 'a temp variable to compare value
    Dim cValue As Double
    Dim row_maxValue As Long
    Dim colRef As Range, colRef2 As Range, outputRng As Range
    Dim oWS As Worksheet
    'Dim lineType As Integer
    Set oWS = ActiveSheet
    
    Set colRef = AddInputBox_rng(prompt:="PLEASE SELECT THE COLUMN FOR DETERMINING GROUP. ")
    If colRef Is Nothing Then Exit Sub
    Set colRef2 = AddInputBox_rng(prompt:="PLEASE SELECT THE COLUMN FOR IDENTIFYING MAXIMUM VALUE. ")
    If colRef2 Is Nothing Then Exit Sub
    Set outputRng = AddInputBox_rng(prompt:="PLEASE SELECT THE TARGET OUTPUT LOCATION. ")
    If outputRng Is Nothing Then Exit Sub
    colNum = colRef.column
    colNum2 = colRef2.column
    count = 0
    
    cRow = Selection.row
    lRow = Selection.row + Selection.count - 1
    
    For i = cRow To lRow
    
        lRowOfSameGrp = FindLastRowOfSameGrp(i, colNum)
        
        cValue = Cells(i, colNum2)
        maxValue = cValue
        row_maxValue = i
        
        For j = i To lRowOfSameGrp
            cValue = Cells(j, colNum2)
            If cValue > maxValue Then
                maxValue = cValue
                row_maxValue = j
            End If
        Next j
        outputRng.Worksheet.Cells(outputRng.row + count, outputRng.column) = oWS.Cells(row_maxValue, colNum)
        outputRng.Worksheet.Cells(outputRng.row + count, outputRng.column + 1) = oWS.Cells(row_maxValue, colNum2)
        count = count + 1
        i = j
    Next i

End Sub
Public Sub AddThickBottomLineAtPageBreak()
    Dim myWS As Worksheet, i As Long
    Set myWS = ActiveSheet
    Dim myRow As Long
    Debug.Print myWS.PageSetup.PrintArea

    For i = 1 To myWS.HPageBreaks.count
        myRow = myWS.HPageBreaks(i).Location.row - 1
        With rows(myRow).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                        .ColorIndex = 1
        End With
    Next
End Sub
Public Sub SetFooter()
    ActiveSheet.PageSetup.RightHeader = "Page &P of &N"
    ActiveSheet.PageSetup.RightFooter = "&""Arial""&8" & "Printed at &D &T" & Chr(10) & "  &Z&F"
 'ActiveSheet.PageSetup.RightFooter.Text.Font = 8
End Sub
Private Function AddInputBox_rng(prompt As String, Optional title As String, Optional default As Range) As Range
    If default Is Nothing Then
        Set default = Range("D1")
    End If
    On Error GoTo ErrHandler:
        Set AddInputBox_rng = Application.inputBox(prompt:=prompt, title:=title, default:=default.address, Type:=8)
    On Error GoTo 0
    Exit Function
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    End
End Function

Private Function FindLastRowOfSameGrp(ByVal rRow As Long, ByVal rCol As Long, Optional ws As Worksheet, Optional maxRow As Long) As Long
    Dim cRow As Long
    Dim savedText As String, cText As String
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    cRow = rRow + 1
    
'find lRow_SchSameGrp
    With ws
        savedText = .Cells(rRow, rCol).Text
        cText = .Cells(cRow, rCol).Text
        Do Until .Cells(cRow, rCol).Text <> savedText And .Cells(cRow, rCol).Text <> vbNullString Or cRow > maxRow
            cRow = cRow + 1
            cText = .Cells(cRow, rCol).Text
        Loop
        If cRow > maxRow + 1 Then cRow = 0
        FindLastRowOfSameGrp = cRow - 1
    End With
End Function


