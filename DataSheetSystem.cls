VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataSheetSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("DataSheet.System")

Private ws_sys As Worksheet
Private wsInteract As clsWorksheetsInteraction
Private genFuncs As clsGeneralFunctions
Private Const DataDeliminator As String = "##"

Sub Class_Initialize()
    Set ws_sys = Worksheets("Data_System")
    Set wsInteract = New clsWorksheetsInteraction
    Set genFuncs = New clsGeneralFunctions
End Sub

'Function SetProp(propGrp As String, propName As String) As Variant
'    SetProp = FindPropVal(FindTableRange(propGrp), propName)
'End Function
'
'Function GetProp(propGrp As String, propName As String) As Variant
'    GetProp = FindPropRng(FindTableRange(propGrp), propName).Value
'End Function



Property Let prop(propGrp As String, propName As String, val As Variant)
    Dim rng1 As Range
    Set rng1 = FindPropRng(FindTableRange(propGrp), propName)
    
    If IsArray(val) Then
        wsInteract.WriteArrToRow val, rng1.row, rng1.column, rng1.Worksheet
    Else
        rng1 = val
    End If
End Property

Property Get prop(propGrp As String, propName As String) As Variant
    'data saved in a row. 1 item in array = 1 cell. Hence, data size is limited to ~10000 (Max. column in excel)
    prop = FindPropVal(FindTableRange(propGrp), propName)
End Property

Property Let propCondense(propGrp As String, propName As String, val As Variant)
    'data saved in a row. Every 500 items in array = 1 cell, split by const DataDeliminator = "##". Hence, data size is limited to ~10000*500 (Max. column in excel)
    Dim rng1 As Range
    Set rng1 = FindPropRng(FindTableRange(propGrp), propName)
    
    If Not genFuncs.isInitialised(val) Then Exit Property
    If IsArray(val) Then
        val = CondenseArrToString(val)
        wsInteract.WriteArrToRow val, rng1.row, rng1.column, rng1.Worksheet
    Else
        rng1 = val
    End If
End Property

Property Get propCondense(propGrp As String, propName As String) As Variant
    'data saved in a row. 1 item in array = 1 cell. Hence, data size is limited to ~10000 (Max. column in excel)
    propCondense = FindPropValCondense(FindTableRange(propGrp), propName)
End Property

Property Get propNames(propGrp As String) As String()
    Dim rng_tbl As Range, var As Variant
    Set rng_tbl = FindTableRange(propGrp)
    var = wsInteract.SetArrayOneDim(rng_tbl.row, rng_tbl.column, True, arrSize:=rng_tbl.count, ws:=rng_tbl.Worksheet)
    propNames = genFuncs.CStr_arr(var)
End Property

Private Function FindTableRange(propGrp As String, Optional isIncludeHeader As Boolean = False) As Range
'Find the table range as per table name input (table name = saved property group name)
'The <table=xxx> and <End Table> row is not considered in the range
    Dim str1 As String, str2 As String
    Dim rng1 As Range, rng2 As Range
    str1 = "<" & propGrp & ">"
    str2 = "<End Property>"
    With ws_sys

        Set rng1 = .Range("A:A").Find(str1)
        Set rng2 = .Range("A:A").Find(str2, After:=rng1)
        If isIncludeHeader Then
            Set FindTableRange = .Range(.Cells(rng1.row, rng1.column), .Cells(rng2.row, rng2.column))
        Else
            Set FindTableRange = .Range(.Cells(rng1.row + 1, rng1.column), .Cells(rng2.row - 1, rng2.column))
        End If
        'Debug.Print FindTableRange.row
        'Debug.Print FindTableRange.count
        
    End With
End Function

Private Function FindPropRng(rng_table As Range, prop As String) As Range
    'return the value of the property
    Dim rng1 As Range, lCol As Long
    With ws_sys
        Set rng1 = rng_table.Find(prop)
        If rng1 Is Nothing Then
            g_log.WriteLog "Cannot Find Prop '" & prop & "' in 'Data_System' Sheet."
            MsgBox "Error: Cannot Find Prop '" & prop & "' in 'Data_System' Sheet."
            End
        End If
        lCol = wsInteract.FindLastCol(rng1.row, rng1.column, ws_sys)
    End With
    If lCol = 1 Then
        Set FindPropRng = ws_sys.Cells(rng1.row, rng1.column + 1)
    Else
        Set FindPropRng = ws_sys.Range(ws_sys.Cells(rng1.row, rng1.column + 1), ws_sys.Cells(rng1.row, lCol))
    End If
End Function

Private Function FindPropVal(rng_table As Range, prop As String) As Variant
    'return the value of the property
    Dim rng1 As Range, var As Variant
    With ws_sys
        Set rng1 = rng_table.Find(prop)
        var = wsInteract.SetArrayOneDim(rng1.row, rng1.column + 1, False, ws:=rng1.Worksheet)
    End With
    FindPropVal = var
End Function

Private Function FindPropValCondense(rng_table As Range, prop As String) As Variant
    'return the value of the property
    Dim rng1 As Range, var As Variant, str As String
    Dim i As Long
    Dim arrSize As Long
    With ws_sys
        Set rng1 = rng_table.Find(prop)
        If rng1 Is Nothing Then
            g_log.WriteLog "Cannot Find Prop '" & prop & "' in 'Data_System' Sheet."
            MsgBox "Error: Cannot Find Prop '" & prop & "' in 'Data_System' Sheet."
            End
        End If
        var = wsInteract.SetArrayOneDim(rng1.row, rng1.column + 1, False, ws:=rng1.Worksheet)
    End With
    
    
    If IsArray(var) Then
        arrSize = UBound(var) - LBound(var) + 1
        str = var(LBound(var))
        For i = LBound(var) + 1 To UBound(var)
            str = str & DataDeliminator & var(i)
        Next i
    Else
        str = var
    End If
    
ExitFunc:
    If Not str = vbNullString Then
        FindPropValCondense = Split(str, DataDeliminator)
    Else
        FindPropValCondense = str
    End If
End Function

Private Function CondenseArrToString(arr As Variant, Optional n As Integer = 500) As Variant
    Dim newArr As Variant, str As String, newArrSize As Long
    Dim oldArrSize As Long
    Dim i As Long, count1 As Long, count2 As Long
    
    oldArrSize = UBound(arr) - LBound(arr) + 1
    newArrSize = Int(oldArrSize / n) + 1
    ReDim newArr(newArrSize - 1)
    
    count2 = 0
    str = arr(0)
    For i = 1 To UBound(arr)
        count1 = count1 + 1
        str = str & DataDeliminator & arr(i)
        If count1 = n Then
            newArr(count2) = str
            i = i + 1
            If Not i = UBound(arr) Then
                str = arr(i)
            Else
                str = vbNullString
            End If
            count1 = 1
            count2 = count2 + 1
        End If
    Next i
    
    If Not str = vbNullString Then
        newArr(count2) = str
    End If
    CondenseArrToString = newArr
End Function

Private Function isPropValEmpty(rng_table As Range, prop As String) As Variant
    Dim var As Variant
    var = FindPropVal(rng_table, prop)
    If Not IsArray(var) Then
        If CStr(FindPropVal(rng_table, prop)) = vbNullString Then isPropValEmpty = True
    End If
End Function

Public Sub CreateCustomPropGrp(propGrpName As String, ParamArray props() As Variant)
'0 for sucess, -1 if tableName already exists
    Dim lRow As Long, i As Long
    Dim count As Long
    Dim ret As Integer
    
    'check if table already exist
    If isPropGrpExist(propGrpName) Then
        Exit Sub
    End If
    
    'Find Last Row of Sheet and add table
    lRow = wsInteract.FindLastRowOfWS(1, ws_sys)
    count = 1
    With ws_sys
        .Cells(lRow + 2, 1) = "<" & propGrpName & ">"
        For i = LBound(props) To UBound(props)
            .Cells(lRow + 2 + count, 1) = props(i)
            count = count + 1
        Next i
        .Cells(lRow + 2 + count, 1) = "<End Property>"
    End With
    
    
    'register the CustomPropGrp to the 'CustomPropGroupList.Name' Property
    AddPropVal "CustomPropGroupList", "Name", propGrpName
    
'SetReturnVal:
    'CreateCustomPropGrp = ret
End Sub

'Public Function AddPropToPropGrp(propGrpName As String, ParamArray props() As Variant)
'
'End Function
Private Sub AddPropVal(propGrp As String, propName As String, val As Variant)
'add single prop val to the existing prop grp
    Dim rng1 As Range, rng_tbl As Range
    Set rng_tbl = FindTableRange(propGrp)
    Set rng1 = FindPropRng(rng_tbl, propName)
    If isPropValEmpty(rng_tbl, propName) Then
        rng1.value2 = val
    Else
        ws_sys.Cells(rng1.row, rng1.column + rng1.count) = val
    End If
End Sub

Public Sub DeleteAllCustomPropGrp()
    Dim propGrps As Variant, i As Integer
    
    propGrps = Me.prop("CustomPropGroupList", "Name")
'    If Not IsArray(propGrps) Then
'        Dim temp As Variant
'        temp = propGrps
'        ReDim propGrps(0)
'        propGrps(0) = temp
'    End If
    If genFuncs.isInitialised(propGrps) Then
        For i = LBound(propGrps) To UBound(propGrps)
            DeletePropGrp CStr(propGrps(i))
        Next i
        ClearPropGrpVal "CustomPropGroupList"
    End If
    


End Sub

Public Sub DeleteCustomPropGrp(propGrp As String)
    Dim var1 As Variant, var2 As Variant
    Dim i As Integer, count As Integer
    'Delete the table
    DeletePropGrp propGrp
    
    'Modify the value in 'CustomPropGroupList.Name'
    var1 = prop("CustomPropGroupList", "Name")
    ReDim var2(LBound(var1) To UBound(var1) - 1)
    count = LBound(var1)
    For i = LBound(var1) To UBound(var1)
        If Not var1(i) = propGrp Then
            var2(count) = var1(i)
            count = count + 1
        End If
    Next i
    ClearPropGrpVal "CustomPropGroupList"
    prop("CustomPropGroupList", "Name") = var2
End Sub

Public Sub DeletePropGrp(propGrp As String)
    ClearPropGrpVal propGrp
    FindTableRange(propGrp, True).Clear
End Sub

Public Sub ClearPropGrpVal(propGrp As String)
    Dim propNames() As String, propName As Variant
    propNames = Me.propNames(propGrp)
    
    For Each propName In propNames
        FindPropRng(FindTableRange(propGrp), CStr(propName)).Clear
    Next
End Sub

Public Function isPropGrpExist(propGrpName As String) As Boolean
    Dim rng1 As Range
    Set rng1 = ws_sys.Range("A:A").Find("<" & propGrpName & ">")
    If Not rng1 Is Nothing Then
        isPropGrpExist = True
    End If
End Function

Public Sub RestoreToDefault()
    prop("isCreated", "member") = False
    prop("isCreated", "frameOrientationType") = False
    prop("isCreated", "isIdentifiedConn") = False
    prop("isCreated", "isMappedConn") = False
End Sub
