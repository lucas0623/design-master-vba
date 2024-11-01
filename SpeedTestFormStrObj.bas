Attribute VB_Name = "SpeedTestFormStrObj"
'@Folder("Tests")
Private mModel As StrModel
Private mDsManager As DataSheetManager

Sub Init()
    Set mModel = New StrModel
    Set mDsManager = New DataSheetManager
End Sub

Sub ReadFrmTableToDf()
    Init
    Dim time1 As Variant, timeUsed As Variant
    
    Dim df As clsDataFrame
    Set df = New clsDataFrame
    Dim DS As oDataSheet
    Set DS = mDsManager.DSFrameData
    time1 = Timer
    With DS.tagSelector
        Set df = DS.GetDataframe(.eleID, .section, .jtI, .jtJ, .eleLen, .memID)
    End With
    timeUsed = Timer - time1
    Debug.Print "Time Used: " & timeUsed
    Debug.Print "Num of Row: " & df.CountRows
End Sub

Sub ReadFrmForceTableToDf()
    Init
    Dim time1 As Variant, timeUsed As Variant
    
    Dim df As clsDataFrame
    Set df = New clsDataFrame
    Dim DS As oDataSheet
    Set DS = mDsManager.DSFrameForce
    time1 = Timer
    With DS.tagSelector
        Set df = DS.GetDataframe(.eleID, .station, .loadComb, .stepType, .p, .V2, .V3, .t, .M2, .M3, .section, .memID)
    End With
    timeUsed = Timer - time1
    Debug.Print "Time Used: " & timeUsed
    Debug.Print "Num of Row: " & df.CountRows
End Sub

Sub FormFrmObjAll()
    Init
    Dim time1 As Variant, timeUsed As Variant
    Dim ret As Long
    
    time1 = Timer

    ret = mModel.Constructor.FormFrmObj
    timeUsed = Timer - time1
    Debug.Print "Time Used: " & timeUsed
    Debug.Print "Num of Frame: " & mModel.frames.count
End Sub

Sub FormFrmForceObjAll()
    Init
    Dim time1 As Variant, timeUsed As Variant
    Dim ret As Long
    ret = mModel.Constructor.FormFrmObj
    time1 = Timer
    ret = mModel.Constructor.FormFrmForceObj
    timeUsed = Timer - time1
    Debug.Print "Time Used: " & timeUsed
    Debug.Print "Num of FrameForce: " & mModel.frmForces.count
End Sub

Sub LoopFrmForceObjAll()
    
    Dim time1 As Variant, timeUsed As Variant
    Dim ret As Long
    Dim frmForce As StrFrameForce, frmForces As Collection
    Dim x As Double, maxP As Double, cForce As StrFrameForce
    Dim i As Long
    FormFrmForceObjAll
    Set frmForces = mModel.frmForces
    time1 = Timer
    Set cForce = frmForces(1)
    For Each frmForce In frmForces
        x = frmForce.force(1)
        If frmForce.force(1) > cForce.force(1) Then Set cForce = frmForce
    Next
'    Set cForce = frmForces(1)
'    For i = 1 To frmForces.count
'        x = frmForces(i).force(1)
'        If frmForces(i).force(1) > cForce.force(1) Then Set cForce = frmForces(i)
'    Next
    timeUsed = Timer - time1
    Debug.Print "Time Used: " & timeUsed
    Debug.Print "Num of FrameForce: " & mModel.frmForces.count
End Sub
