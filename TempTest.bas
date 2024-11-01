Attribute VB_Name = "TempTest"
'@Folder("Tests")
Sub ShowUF()
    'Dim vm As VMForceExtraction

    
End Sub
Sub showUF2()
    Dim vm As VMForceExtraction
    Dim uf As New UFExtractForceMemByPos
    
    Set vm = New VMForceExtraction
    
    uf.IView_Initialize vm
    vm.Initialize "TEST"
    
    uf.IView_Show
    
End Sub

Sub showUF3()
    Dim vm As VMForceExtraction
    Dim uf As New UFExtractForceMemForGraph
    
    Set vm = New VMForceExtraction
    
    uf.IView_Initialize vm
    vm.Initialize "TEST"
    
    uf.IView_Show
    
End Sub

Sub showUF4()
    Dim vm As VMForceExtraction
    Dim uf As New UFExtractConnForce
    
    Set vm = New VMForceExtraction
    
    uf.IView_Initialize vm
    vm.Initialize "TEST"
    
    uf.IView_Show
    
End Sub

Sub TestCopyVal()
    Dim var As Variant
    var = Range("J11:J13")
    'Range("K11:K12").value2 = Range("J11:J13").value2
    Range(Range("K11"), Range("K15")).value2 = Range("J11:J13").value2
    'Range("K11:K13") = var
End Sub

Sub TestNextEmptyCell()
    Dim wsInteract As New clsWorksheetsInteraction
    Dim rng As Range
    Set rng = wsInteract.FindNextNonEmptyCell(Range("E6"), 10)
    Debug.Print "Next Non Empty Row = " & rng.row
End Sub

Sub TestDictFunction()
    Dim dict1 As Object, dict2 As Object, dict3 As Object
    Dim duplicateItems As Object
    Dim DictFunc As New libDictionaryFunctions
    Set dict1 = CreateObject("Scripting.Dictionary")
    Set dict2 = CreateObject("Scripting.Dictionary")
    Set dict3 = CreateObject("Scripting.Dictionary")
    'Set duplicateItems = CreateObject("Scripting.Dictionary")
    dict1.Add "A", "A"
    dict1.Add "B", "B"
    dict1.Add "C", "CCC"
    
    dict2.Add "A", "A"
    dict2.Add "E", "E"
    dict2.Add "F", "CCC"
    
'    dict3.Add "D", "D"
'    dict3.Add "E", "E"
'    dict3.Add "C", "CCC"
    Set duplicateItems = DictFunc.GetDuplicateItemsInDicts(dict1, dict3)
    Dim key As Variant
    For Each key In duplicateItems.keys
        Debug.Print duplicateItems(key)
    Next key
    Debug.Print duplicateItems.count
End Sub
Sub TestCollFunction()
    Dim coll As New Collection
    Dim jt As New StrJoint
    jt.Init "AA"
    coll.Add jt, "a"
    coll.Add jt, "b"
    coll.Add jt, "c"
    
    Dim CollFunc As New LibCollectionMethod
    Dim ret As Boolean
    
    ret = CollFunc.isKeyExist(coll, "F")
    Debug.Print ret
End Sub

Sub TestStrLimit()
    Dim str As String, i As Long
    On Error GoTo Err:
    str = "ABCDE"
    For i = 1 To 100
        str = str & "," & "ABCDE"
    Next i
    
    Debug.Print "COMPLETE"
    Exit Sub
Err:
    Debug.Print "FAIL"
End Sub

Sub testInt()
    Dim x As Long, y As Long
    Dim z As Long
    x = 4596
    y = 500
    z = Int(x / y) * y + x Mod y
    Debug.Print z
    
End Sub

Sub TestCondenseProp()
    Dim i As Long, count As Long
    Dim data As Variant, dataSize As Long
    Dim dataReturn As Variant
    Dim DsSys As New DataSheetSystem
    
    dataSize = 100000
    
    ReDim data(dataSize - 1)
    For i = 0 To dataSize - 1
        data(i) = "ABC-" & i
    Next i
    DsSys.propCondense("CondenseDataTest", "CondenseData") = data
    dataReturn = DsSys.propCondense("CondenseDataTest", "CondenseData")
    Debug.Print "ArrSize: = " & UBound(dataReturn) - LBound(dataReturn) + 1
    Debug.Print "First Item: = " & dataReturn(0)
    Debug.Print "Second Item: = " & dataReturn(1)
    Debug.Print "Last Item: = " & dataReturn(UBound(dataReturn))
End Sub

Sub TestReadTable()
    Dim wsInteract As New clsWorksheetsInteraction
    Dim fCol As Long, lCol As Long
    Dim df As New clsDataFrame, ws As Worksheet
    Set ws = Worksheets("Data_FrameForce (2)")
    df.Init_ReadWorksheet2 ws, headTags:=Split("eleID,loadComb", ",")
    
    Debug.Print df.idata(1, 2)
    'Debug.Print df.idata(100000, 2)
    Debug.Print df.CountRows
    
    Dim dsManager As New DataSheetManager
    Dim dsForce As oDataSheet
    Set dsForce = dsManager.DSFrameForce
    With dsForce.tagSelector
        dsForce.WriteDataframe df, False, True, .eleID, .loadComb
    End With
End Sub
Sub TestReadTable2()
    Dim wsInteract As New clsWorksheetsInteraction
    Dim fCol As Long, lCol As Long
    Dim df As New clsDataFrame, ws As Worksheet
    Set ws = Worksheets("Data_FrameForce")
    df.Init_ReadWorksheet2 ws, headTags:=Split("eleID,loadComb", ",")
    
    Debug.Print df.idata(1, 2)
    'Debug.Print df.idata(100000, 2)
    Debug.Print df.CountRows

End Sub
Sub TestSplitArr()
    Dim arr As Variant
    ReDim arr(0 To 53)
    Dim arr2
    For i = 0 To 53
        arr(i) = i
    Next i
    arr2 = SplitArr(arr, 3, 12)
    Debug.Print arr2(0)
    Debug.Print arr2(9)
End Sub

Private Function SplitArr(arr As Variant, startIndex As Long, endIndex As Long) As Variant
    Dim var As Variant, i As Long
    ReDim var(0 To endIndex - startIndex)
    For i = 0 To endIndex - startIndex
        var(i) = arr(i + startIndex)
    Next i
    SplitArr = var
End Function

Sub TestCopyFormula()
    Dim formula As String
     Cells(4, 13).Copy Cells(4, 26)
    'mWS.Range(mWS.Cells(iRow + 1, mCollProps(propName).Loc)).Copy mWS.Cells(startRow, rCol)
    'Range("P4").formula = formula
End Sub

Sub TestFilter()
    Dim model As New StrModel
    Dim dsManager As New DataSheetManager
    Dim genFunc As New clsGeneralFunctions
    Dim filter() As String
    'Dim g_log As New clsLog
    'g_log.CreateNewFile True
    model.Constructor.FormJointObj
    model.Constructor.FormFrmObj
    Dim frms As Collection
    Set frms = model.frames
    Debug.Print frms.count
    filter = genFunc.CStr_arr(Split("B0,B1", ","))
    Set frms = model.frmsBySection("B1")
    'Set frms = model.FilterCollOfObj(frms, "section", filter)
    'Set frms = model.frames("B0")
    Debug.Print frms.count
    'g_log.CloseFile
End Sub
