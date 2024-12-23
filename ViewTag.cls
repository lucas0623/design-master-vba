VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ViewTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Operation"

Private ds_sys As DataSheetSystem ', model As StrModel
Private wsInteract As clsWorksheetsInteraction
Private genFunc As clsGeneralFunctions
Private UI As clsUIManager
Private dsSum As DataSheetSummary

Private Sub Initialize()

    Set ds_sys = New DataSheetSystem
    Set wsInteract = New clsWorksheetsInteraction
    Set genFunc = New clsGeneralFunctions
    Set UI = New clsUIManager
    'Set model = globalControl.model
    Set dsSum = New DataSheetSummary
End Sub

Sub ViewDesignWSTag()
    Dim form1 As UFBasic
    Dim outputSheetName As Variant, cb_outputSheetName As msForms.ComboBox
    Dim designWSName As Variant, cb_designWSName As msForms.ComboBox, dataWS As Worksheet
    Dim arr As Variant, arr_header As Variant
    Dim arr2 As Variant, arr_header2 As Variant
    Dim dict As Object, dict2 As Object
    Dim i As Long, j As Long

    Initialize
    Set dataWS = wsInteract.setWorksheet("WSData", errText:="that contains the target worksheet input/output data")

    Set form1 = New UFBasic
    form1.Initialize 350, True
    form1.TitleBarCaption = "View Tag of Design Worksheets"

    'The Combo Boxes
    outputSheetName = GetSummmarySheetsName
    designWSName = GetDesignWSName
    designWSName = genFunc.DeDupeOneDimArray(designWSName)


    'Design Worksheet Input Data

    Set dict = CreateObject("Scripting.Dictionary")
    Set dict2 = CreateObject("Scripting.Dictionary")
    For i = 0 To UBound(designWSName)
        For j = 0 To UBound(outputSheetName)
            'dict->WS Input. dict2->WS Output
            dict(designWSName(i) & outputSheetName(j)) = GetDesignTagArr(CStr(designWSName(i)), dataWS, Worksheets(outputSheetName(j)), True)
            dict2(designWSName(i) & outputSheetName(j)) = GetDesignTagArr(CStr(designWSName(i)), dataWS, Worksheets(outputSheetName(j)), False)
        Next j
    Next i
    arr = dict(designWSName(0) & outputSheetName(0))
    arr_header = Array("Tag Text", "Location", "Column")
    arr2 = dict2(designWSName(0) & outputSheetName(0))
    arr_header2 = Array("Tag Text", "Location", "Column")

'    'Design Worksheet Input Data
'    Set dict = CreateObject("Scripting.Dictionary")
'    For i = 0 To UBound(designWSName)
'        dict(designWSName(i)) = GetDesignWSTag(CStr(designWSName(i)), dataWS, True, False)
'    Next i
'    designWSName = dict.keys 'remove duplicate
'    arr = GetDesignWSTag(CStr(designWSName(0)), dataWS, True)
'    arr_header = Array("Tag Text", "Location")

'    'Tag Location in Summary Sheet
'    Set dict2 = CreateObject("Scripting.Dictionary")
'    For i = 0 To UBound(designWSName)
'        If Not dict.exists(designWSName(i)) Then
'            dict(designWSName(i)) = GetDesignWSTag(CStr(designWSName(i)), dataWS, True, False)
'        End If
'    Next i


    form1.AddComboBox designWSName, cb_designWSName, "Design Worksheet Name"
    form1.AddComboBox outputSheetName, cb_outputSheetName, "Summary Worksheet Name"
    form1.AddMultiColumnsListBox arr, arr_header, dict:=dict, ControlCB:=cb_designWSName, _
                                ControlCB2:=cb_outputSheetName, width_LB:=360, colWidth:="180;100;60", title:="Design Worksheet INPUT", height_LB:=200
    form1.AddMultiColumnsListBox arr2, arr_header2, dict:=dict2, ControlCB:=cb_designWSName, _
                                ControlCB2:=cb_outputSheetName, width_LB:=360, colWidth:="180;100;60", title:="Design Worksheet OUTPUT", height_LB:=150
    'Design Worksheet Output Data
    form1.AdjustHeight
    form1.Show
End Sub
Sub ViewExtractionTag()

    Dim form1 As UFBasic
    Dim outputSheetName As Variant, cb_outputSheetName As msForms.ComboBox
    Dim arr As Variant, arr_header As Variant
    Dim dict As Object

    Initialize

    Set form1 = New UFBasic
    form1.Initialize 350, True
    form1.TitleBarCaption = "Setting - Forces Extraction"

    outputSheetName = GetSummmarySheetsName
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 0 To UBound(outputSheetName)
        dict(outputSheetName(i)) = GetDataExtractionTagInSummary(Worksheets(outputSheetName(i)))
    Next i
    arr = GetDataExtractionTagInSummary(Worksheets(outputSheetName(0)))
    arr_header = Array("Tag Description", "Tag Text", "Column")

    form1.AddComboBox outputSheetName, cb_outputSheetName, "Summary Worksheet Name:"
    form1.AddMultiColumnsListBox arr, arr_header, dict:=dict, ControlCB:=cb_outputSheetName, width_LB:=360, colWidth:="180;100;60", title:="", height_LB:=280
    form1.AdjustHeight
    form1.Show
End Sub
Private Function GetDesignTagArr(tarWsName As String, dataWS As Worksheet, sumWS As Worksheet, isInputData As Boolean) As Variant
    Dim arr As Variant, arr1 As Variant, arr2 As Variant
    Dim tags() As String

    arr1 = GetDesignWSTag(tarWsName, dataWS, isInputData, False)
    
    If genFunc.isInitialised(arr1) Then
        tags = genFunc.CStr_arr(GetDesignWSTag(tarWsName, dataWS, isInputData, True))
        arr2 = GetDesignWSTagInSummary(tags, sumWS)
    Else
        ReDim arr1(1, 2)
        arr1(1, 0) = "-"
        arr1(1, 1) = "-"
        arr1(1, 2) = "-"
        ReDim arr2(0)
        arr2(0) = "-"
    End If
    ReDim arr(UBound(arr2), 2)

    For i = 0 To UBound(arr2)
        arr(i, 0) = arr1(i + 1, 0 + 1)
        arr(i, 1) = arr1(i + 1, 1 + 1)
        arr(i, 2) = arr2(i)
    Next i
    GetDesignTagArr = arr
End Function

Private Function GetDesignWSTag(tarWsName As String, dataWS As Worksheet, isInputData As Boolean, isTagNameOnly As Boolean) As Variant
    Dim col_tarWSData As Long, fRow_tarWSData As Long
    Dim arrSizeX As Long
    Dim arr As Variant

    fRow_tarWSData = wsInteract.getLocVar(dataWS, "rRow_tarWSData", False, _
                               "'fRow_tarWSData' text (used for locating the first row of input/output data of the target Worksheets)") + 1
    If isInputData Then
        col_tarWSData = wsInteract.getLocVar(dataWS, tarWsName, True, "target worksheet name")
    Else
        col_tarWSData = wsInteract.getLocVar(dataWS, tarWsName, True, "target worksheet name") + 2
    End If
    
    If isTagNameOnly Then
        arr = wsInteract.SetArrayOneDim(fRow_tarWSData, col_tarWSData, ws:=dataWS)
    Else
        arr = wsInteract.SetArrayTwoDim(fRow_tarWSData, col_tarWSData, arrSizeX:=2, ws:=dataWS)
    End If


    GetDesignWSTag = arr
End Function


Private Function GetDesignWSTagInSummary(tags() As String, sumWS As Worksheet) As Variant
    Dim i As Long, arr As Variant, col As Long
    ReDim arr(LBound(tags) To UBound(tags))
    With sumWS
        For i = LBound(tags) To UBound(tags)
            col = wsInteract.getLocVar(sumWS, tags(i), True, isMustMatch:=False)
            arr(i) = Number2Letter(col)
            
        Next i
    End With
    GetDesignWSTagInSummary = arr
End Function
Private Function GetDataExtractionTagInSummary(ws As Worksheet) As Variant
    Dim arr1() As String, arr2() As String, arr3() As Long
    Dim arr As Variant

    arr1 = dsSum.GetTagsDesc
    arr2 = dsSum.GetTagsText
    arr3 = dsSum.GetTagColumns(ws)
'    If Not LBound(col_summary) = 0 Then
'        genFunc.ChangeArrBase col_summary, 0
'    End If

    ReDim arr(UBound(arr1), 2)
    For i = 0 To UBound(arr1)
        arr(i, 0) = arr1(i)
        arr(i, 1) = arr2(i)
        arr(i, 2) = Number2Letter(arr3(i))
    Next i

    GetDataExtractionTagInSummary = arr
    'This function return the arr/list of the existing tag

End Function

Private Function GetDesignWSName() As String()
    Dim dataWS As Worksheet
    Dim rng_wsName() As Range, arr() As String, i As Long
    Set dataWS = wsInteract.setWorksheet("WSData", errText:="that contains the target worksheet input/output data")
    rng_wsName = wsInteract.FindNonEmptyCell(dataWS.rows(1))
    ReDim arr((UBound(rng_wsName) + 1) / 4 - 1)
    For i = 0 To UBound(arr)
        arr(i) = rng_wsName(i).Text
    Next i
    GetDesignWSName = arr
End Function


Private Function GetSummmarySheetsName() As Variant
    Dim ws_count As Long, arrL As Object
    Dim ws As Worksheet
    Set arrL = CreateObject("System.Collections.ArrayList")

    'Set ws = ActiveSheet
    'With ActiveSheet
        If InStr(1, ActiveSheet.Name, "Sum") > 0 Then
            arrL.Add ActiveSheet.Name
        End If
    'End With
    With ActiveWorkbook
        ws_count = .Worksheets.count
        'Set the activesheet as the first item

        For i = 1 To ws_count
            If InStr(1, .Worksheets(i).Name, "Sum") > 0 And Not .Worksheets(i).Name = ActiveSheet.Name Then
                arrL.Add .Worksheets(i).Name
            End If
        Next i
    End With

    GetSummmarySheetsName = arrL.ToArray
End Function

Private Function Number2Letter(colNum As Long) As String
'PURPOSE: Convert a given number into it's corresponding Letter Reference
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim ColumnLetter As String


'Convert To Column Letter
    If colNum = -1 Then
        Number2Letter = "-"
    Else
        Number2Letter = Split(Cells(1, colNum).address, "$")(1)
    End If

End Function


