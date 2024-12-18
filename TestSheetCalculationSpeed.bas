Attribute VB_Name = "TestSheetCalculationSpeed"
'@Folder("Operation")
Sub TestCalculationSpeed()
    Dim startTime As Double
    Dim endTime As Double
    Dim i As Integer
    Dim ws As Worksheet
    
    Set ws = Worksheets("CH")
    ' Set calculation mode to manual
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.FormatStaleValues = False
    
    ' Record the start time
    startTime = Timer
    
    ' Perform the calculation 100 times
    For i = 1 To 100
        With ws
            If i Mod 2 = 1 Then
                .Range("E8") = "125x65x15"
            Else
                .Range("E8") = "100x50x10"
            End If
            .Range("M15") = i
            .Range("M16") = i
            .Range("M17") = i
            .Range("M18") = i
            .Range("M19") = i
            .Calculate
        End With
    Next i
    
    ' Record the end time
    endTime = Timer
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.FormatStaleValues = True
    Application.StatusBar = False
    ' Calculate the time taken and display
    MsgBox "Time taken for 100 calculations: " & Format(endTime - startTime, "0.00") & " seconds"
    
    ' Optionally, set calculation mode back to automatic
    ' Application.Calculation = xlCalculationAutomatic
End Sub

